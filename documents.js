import express from 'express';
import path from 'path';
import fs from 'fs';
import { getFrenchFormattedDate, fetchTemplate, generateReport, updateAirtableRecord, updateAirtableRecords, getAirtableSchema, processFieldsForDocx, getAirtableRecords, getAirtableRecord, ymd } from './utils.js';
import { Readable } from 'stream';
import archiver from 'archiver';
// import { Stream } from 'stream';
import { GITHUBTEMPLATES } from './constants.js';

const bufferToStream = (buffer) => {
    const stream = new Readable();
    stream.push(buffer);
    stream.push(null); // Signal the end of the stream
    return stream;
};

export const sanitizeFileName = (fileName) => {
    var newFileName = fileName.replace(/[^a-zA-Z0-9-_.\s'éèêëàâîôùçãáÁÉÈÊËÀÂÎÔÙÇ]/g, '');
    newFileName = newFileName.replace(/  /g, ' '); // Sanitize the filename
    // get extension
    var ext = path.extname(newFileName);
    // if extension duplicated, remove it
    newFileName = newFileName.replace(ext+ext, ext);
}

export const documents = [
    {
        name: 'catalogue',
        multipleRecords: true,
        formula: `AND(OR({année}=2025,{année}=""), OR(FIND(lieuxdemij_cumul,"iège"),FIND(lieuxdemij_cumul,"visio")))`,
        titleForming: function(data) {
            return `Catalogue des formations FSH ${data["année"]}`;
        },
        template: 'catalogue.docx',
        view: 'Grid view',
        table: 'Sessions',
        sortField: 'du',
        sortOrder: 'asc',
        // queriedField: null,
    },
    {
        name: 'programme',
        multipleRecords: false,
        titleForming: function(data) {
            let newTitle = data["titre_fromprog"]
            if(data["du"] && data["au"]) { newTitle+= `${ymd(data["du"])}-${data["au"] && ymd(data["au"])}`}
            return newTitle;
            // return `${data["titre_fromprog"]} ${ymd(data["du"])}-${data["au"] && ymd(data["au"])}`;
        },
        template: 'programme.docx',
        table: 'Sessions',
        queriedField: 'recordId',
    },
    {
        name: 'devis',
        multipleRecords: false,
        titleForming: function(data) {
            return `DEVIS FSH ${data["id"]}`;
        },
        template: 'devis.docx',
        table: 'Devis',
        queriedField: 'recordId',
    },
    {
        name: 'facture',
        multipleRecords: false,
        titleForming: function(data) {
            return `Facture ${data["id"]} ${data["nom"]} ${data["prenom"]}`;
        },
        template: 'facture.docx',
        table: 'Inscriptions',
        queriedField: 'recordId',
        dataPreprocessing: function(data) {
            if(data["date_facture"]) {
                data["today"] = new Date(data["date_facture"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' });
            }
            function calculateCost(data) {
                let cost;
                
                if (data["tarif_special"]) {
                    // If "tarif_special" is available, use it
                    cost = data["tarif_special"];
                } else {
                    // Calculate the base cost, considering whether the person is accompanied
                    let baseCost;
                    if (data["accomp"]) {
                        baseCost = (data["Coût adhérent TTC (from Programme) (from Session)"] || 0) / 2;
                    } else {
                        if (data["Adhérent? (from Participant.e)"]) {
                            baseCost = data["Coût adhérent TTC (from Programme) (from Session)"] || 0;
                        } else {
                            baseCost = data["Coût non adhérent TTC (from Programme) (from Session)"] || 0;
                        }
                    }
                    
                    // Apply "rabais" if available
                    if (data["rabais"]) {
                        cost = baseCost * (1 - data["rabais"]);
                    } else {
                        cost = baseCost;
                    }
                }
                
                return cost;
            }
            
            data["Montant"] = calculateCost(data)
            console.log("Montant calc", data["Montant"])
            data['montant'] = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(
                parseFloat(data["Montant"]),
            );  
            
        },
        airtableUpdatedData: function(data) {
            var updatedInvoiceDate = false;
            if(data["date_facture"]) { updatedInvoiceDate = true; }
            var updatedData = { 
                total: data['Montant'].toString()
            }
            if(!updatedInvoiceDate) {
                updatedData["date_facture"] = new Date().toLocaleDateString('fr-CA');
            }
            return updatedData;
        }
    },
    {
        name: 'certif_realisation',
        multipleRecords: false,
        titleForming: function(data) {
            return `Certificat de réalisation ${data["code_fromprog"]} ${ymd(data["au"])} - ${data["nom"]} ${data["prenom"]}`;
        },
        template: 'certif_realisation.docx',
        table: 'Inscriptions',
        queriedField: 'recordId',
        dataPreprocessing: function(data) {
            data['du'] = new Date(data["du"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            data['au'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            data['dureeh_fromprog'] = data["dureeh_fromprog"]/3600;
            data["assiduite"] = data["assiduite"] * 100;
            data['nom'] = Array.isArray(data["nom"]) ? data["nom"][0].toUpperCase() : data["nom"].toUpperCase();
            data['today'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            data['apaye'] = data.moyen_paiement && data.date_paiement;
            data['acquit'] = data["paye"].includes("Payé")
            ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
            : "";
        }
    },
    {
        name: 'attestation',
        multipleRecords: false,
        titleForming: function(data) {
            return `Attestation de formation ${data["code_fromprog"]} ${ymd(data["au"])} - ${data["nom"]} ${data["prenom"]}`;
        },
        template: 'attestation.docx',
        table: 'Inscriptions',
        queriedField: 'recordId',
        dataPreprocessing: function(data) {
            data['du'] = new Date(data["du"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            data['au'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            data['dureeh_fromprog'] = data["dureeh_fromprog"]/3600;
            data["assiduite"] = data["assiduite"] * 100;
            data['nom'] = Array.isArray(data["nom"]) ? data["nom"][0].toUpperCase() : data["nom"].toUpperCase();
            data['today'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            data['apaye'] = data.moyen_paiement && data.date_paiement;
            data['acquit'] = data["paye"].includes("Payé")
            ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
            : "";
        }
    },
    {
        name: 'factures',
        multipleRecords: true,
        titleForming: function(data) {
            return `Facture ${data["id"]} ${data["nom"]} ${data["prenom"]}`;
        },
        template: 'facture.docx',
        table: 'Inscriptions',
        queriedField: 'recordId',
    }
]

/**
 * Fetches and processes data from Airtable based on the provided document configuration.
 * 
 * @param {Object} req - The request object.
 * @param {Object} req.query - The query parameters from the request.
 * @param {string} req.query.recordId - The record ID to fetch when multipleRecords is false.
 * @param {Object} res - The response object.
 * @param {Object} document - The document configuration object.
 * @param {boolean} document.multipleRecords - Flag indicating whether to fetch multiple records.
 * @param {string} document.table - The Airtable table name.
 * @param {string} document.view - The Airtable view name.
 * @param {string} document.formula - The Airtable formula for filtering records.
 * @param {string} document.sortField - The field to sort records by.
 * @param {string} document.sortOrder - The order to sort records (e.g., "asc" or "desc").
 * @param {Function} [document.dataPreprocessing] - Optional function to preprocess the data.
 * 
 * @returns {Promise<Object|Object[]>} The fetched data from Airtable.
 * 
 * @throws {Error} If there is an error during data fetching or processing.
 */
export const getDataAndProcess = async (req, res, document) => {
    try {
        // only fetch the recordId from the query if document.multipleRecords is false
        if(document.multipleRecords) {
            console.log(`Fetching multiple records from table: ${document.table}, view: ${document.view}, formula: ${document.formula}, sortField: ${document.sortField}, sortOrder: ${document.sortOrder}`);
            const data = await getAirtableRecords(document.table, document.view, document.formula, document.sortField, document.sortOrder);
            if (data) {
                console.log('Data successfully retrieved:', `${data.length} records`);
            } else {
                console.error('Failed to retrieve data.');
            }
            if(document.dataPreprocessing) {
                console.log('Preprocessing data...');
                document.dataPreprocessing(data);
            }
            return data;
        }
        
        const { recordId } = req.query;
        console.log(`Fetching single record from table: ${document.table}, recordId: ${recordId}`);
        const data = await getAirtableRecord(document.table, recordId);
        if (data) {
            console.log('Data successfully retrieved:', data);
        } else {
            console.error('Failed to retrieve data.');
        }
        if(document.dataPreprocessing) {
            console.log('Preprocessing data...');
            document.dataPreprocessing(data);
        }
        return data;
    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ success: false, error: error.message });
    }
}


export const generateReportBuffer = async (templateName, data) => {
    try {
        console.log('Creating buffer file...');
        const buffer = await generateReport({
            output: 'buffer',
            template: await fetchTemplate(GITHUBTEMPLATES+templateName),
            data: data,
        });
        return buffer;
    } catch (error) {
        console.error('Error:', error);
    }
}


/**
 * Generates a ZIP archive from provided buffers and sends it as a response (= instant download).
 *
 * @param {Object} res - The HTTP response object.
 * @param {Array} buffers - An array of objects containing filename and content to be included in the ZIP archive.
 * @param {string} zipFileName - The name of the resulting ZIP file.
 * @returns {Promise<void>} - A promise that resolves when the ZIP archive is successfully created and sent.
 */
export const generateAndSendZipReport = async (res, buffers, zipFileName) => {
    console.log('Setting response headers for ZIP file download...');
    console.log('buffers', buffers)
    if(buffers.length === 0) {
        console.error('No buffers to append to ZIP archive');
        return;
    }
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename=${zipFileName}`);

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(res);
    try {
        console.log('Appending files to ZIP archive...');
        for (let i = 0; i < buffers.length; i++) {
            const { filename, content } = buffers[i];
            console.log(`Appending file: ${filename}`);
            const stream = bufferToStream(content);
            archive.append(stream, { name: filename });
        }
        console.log('Finalizing ZIP archive...');
        await archive.finalize();
        console.log('ZIP archive successfully created and sent.');
    } catch (error) {
        console.error('Error creating zip archive:', error);
        res.status(500).send('Error creating zip archive');
    }
}


export const makeSessionFactures = async (res, sessionId) => {
    console.log(`Fetching inscriptions for session: ${sessionId}`);
    const inscriptions = await getAirtableRecords('Inscriptions', 'Grid view', 
        `AND(sessId='${sessionId}',{Statut}="Enregistrée")`
    );
    // console.log("insc", inscriptions)
    if(!inscriptions || inscriptions.length === 0) {
        console.error('Failed to fetch inscriptions');
        return;
    }
    console.log(`Fetched ${inscriptions.records.length} inscriptions`);
    // console.log(inscriptions[0])

    const document = documents.find(doc => doc.name === 'facture');
    console.log(document.dataPreprocessing)

    let buffers = [];
    for (let i = 0; i < inscriptions.records.length; i++) {
        const inscription = inscriptions.records[i];
        console.log(`Processing inscription: ${inscription.id}`);
        
        let data = inscription;
        
        if (document.dataPreprocessing) {
            console.log('Preprocessing data...');
            data = document.dataPreprocessing(data);
        }
        
        const buffer = await generateReportBuffer('facture.docx', data);
        const filename = document.titleForming(data);
        console.log(`Generated report for: ${filename}`);
        
        buffers.push({ filename, content: buffer });
    }

    const zipFileName = `factures_${sessionId}.zip`;
    console.log(`Generating ZIP file: ${zipFileName}`);
    await generateAndSendZipReport(res, buffers, zipFileName);
    console.log('ZIP file generated and sent successfully');
}

