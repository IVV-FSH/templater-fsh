import express from 'express';
import path from 'path';
import fs from 'fs';
import { getFrenchFormattedDate, fetchTemplate, generateReport, updateAirtableRecord, updateAirtableRecords, getAirtableSchema, processFieldsForDocx, getAirtableRecords, getAirtableRecord, ymd, addMissingFields } from './utils.js';
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

export const logIfNotVercel = (message) => {
    if (process.env.VERCEL !== '1') {
      console.log(message);
    }
  };

export function calculTotalPrixInscription(data) {
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


export const sanitizeFileName = (fileName) => {
    var newFileName = fileName.replace(/[^a-zA-Z0-9-_.\s'éèêëàâîôùçãáÁÉÈÊËÀÂÎÔÙÇ]/g, '');
    newFileName = newFileName.replace(/  /g, ' '); // Sanitize the filename
    // get extension
    var ext = path.extname(newFileName);
    // if extension duplicated, remove it
    newFileName = newFileName.replace(ext+ext, ext);
    if(newFileName.length > 150) {
        newFileName = newFileName.slice(0, 150);
    }

    return newFileName;
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
            data['acquit'] = data["paye"].includes("Payé")
            ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
            : "";

            data["Montant"] = calculTotalPrixInscription(data)
            logIfNotVercel("Montant calc", data["Montant"])
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
        name: 'facture_grp',
        multipleRecords: false,
        titleForming: function(data) {
            return `${data["Name"]}`;
        },
        template: 'facture_grp.docx',
        table: 'Factures',

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
            data['Formateurice'] = Array.isArray(data["Formateurice"]) ? data["Formateurice"].join(", ").replace(/"/g, '') : data["Formateurice"].replace(/"/g, '');
            data['du'] = new Date(data["du"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            data['au'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            // data['dureeh_fromprog'] = data["dureeh_fromprog"]/3600;
            // data["assiduite"] = data["assiduite"] * 100;
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
            data['Formateurice'] = Array.isArray(data["Formateurice"]) ? data["Formateurice"].join(", ").replace(/"/g, '') : data["Formateurice"].replace(/"/g, '');
            data['du'] = new Date(data["du"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            data['au'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
            // data['dureeh_fromprog'] = data["dureeh_fromprog"]/3600;
            // data["assiduite"] = data["assiduite"] * 100;
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
        console.log(`Creating buffer file from ${templateName}...`);
        const template = await fetchTemplate(GITHUBTEMPLATES + templateName);
        // const template = fs.readFileSync(path.join('templates', templateName));
        const buffer = await generateReport(
            template,
            data,
        );
        // const buffer = await generateReport({
        //     output: 'buffer',
        //     template,
        //     data,
        // });
        return buffer;
    } catch (error) {
        console.error('Error generating report buffer:', error.message);
        console.error('Stack trace:', error.stack);
        throw new Error('Failed to generate report buffer', error);
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
    logIfNotVercel('Setting response headers for ZIP file download...');
    // console.log('buffers', buffers)

    if(buffers.length === 0) {
        console.error('No buffers to append to ZIP archive');
        return;
    } else {
        logIfNotVercel(`Appending ${buffers.length} files to ZIP archive...`);
    }
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename=${zipFileName}`);

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(res);
    try {
        logIfNotVercel('Appending files to ZIP archive...');
        for (let i = 0; i < buffers.length; i++) {
            const { filename, content } = buffers[i];
            logIfNotVercel(`Appending file: ${filename}`);
            const stream = bufferToStream(content);
            archive.append(stream, { name: filename });
        }
        logIfNotVercel('Finalizing ZIP archive...');
        await archive.finalize();
        console.log('ZIP archive successfully created and sent.');
    } catch (error) {
        console.error('Error creating zip archive:', error);
        res.status(500).send('Error creating zip archive');
    }
}



export const makeGroupFacture = async (factureId) => {
    const factureGrpParams = documents.find(doc => doc.name === 'facture_grp');
    console.log(`Fetching facture: ${factureId}`);
    let data = await getAirtableRecord(factureGrpParams.table, factureId);
    // console.log("inscrits", inscrits)
    if(!data) {
        console.error('Failed to fetch facture');
        return;
    }
    if(data.unicite != "ok") {
        throw new Error(`Erreur dans le groupe pour la facture ${factureId}, l'entité ET la session concernée doivent être uniques`);
    }
    const inscrits = await getAirtableRecords(
        'Inscriptions', 'Grid view', 
        `AND({factGroupId}="${factureId}",{Statut}="Enregistrée")`,
        'nom',
        'asc'
    );
    if(!inscrits || inscrits.length === 0) {
        console.error('Failed to fetch inscrits');
        return;
    }
    // const session = await getAirtableRecord('Sessions', data.sessId);
    // merge data and inscrits, but if there are conflicts, data wins
    // logIfNotVercel('inscr0', inscrits.records[0])
    data = {...inscrits.records[0], ...data};
    // logIfNotVercel("data", data)

    var total = 0.0;

    var stagiaires = inscrits.records.map(inscrit => {
        let stagiaire = {
            nom: inscrit.nom[0],
            prenom: inscrit.prenom[0],
            poste: (inscrit.poste && ", "+inscrit.poste[0]) || "",
            // nom_poste: `${inscrit.prenom[0]} ${inscrit.nom[0] && inscrit.nom[0].toUpperCase()} ${inscrit.poste && ", "+inscrit.poste[0]}`,
            paye: inscrit.paye.includes("✅"),
        }
        const montant = calculTotalPrixInscription(inscrit);
        // total += stagiaire.paye?parseFloat(montant):0;
        if(!stagiaire.paye) {
            total += parseFloat(montant);
        }

        stagiaire.montant = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(
            parseFloat(montant),
        );
        // logIfNotVercel("stagiaire", stagiaire)
        return stagiaire;
    });

    // logIfNotVercel("montant total", total)

    data.stagiaires = stagiaires;
    data.montant = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(
        parseFloat(total),
    );

    // if (factureGrpParams.dataPreprocessing) {
    //     factureGrpParams.dataPreprocessing(data);
    // }
    
    const buffer = await generateReportBuffer(factureGrpParams.template, data);
    const filename = sanitizeFileName(getFrenchFormattedDate()+" "+factureGrpParams.titleForming(data)+".docx");
    console.log(`Generated report for: ${filename}`);
    // downloadDocxBuffer(res, filename, buffer);
    // res.setHeader('Content-Disposition', `attachment; filename="${filename}.docx"`);
    // res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    // res.setHeader('Content-Length', buffer.length); // Ensure the buffer length is correctly sent

    // // Send the buffer as a binary response
    // res.end(buffer, 'binary');
    return { filename:filename, content: buffer };
};



export const downloadDocxBuffer = (res, filename, buffer) => {
    const encodedFileName = encodeURIComponent(filename);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename=${encodedFileName}`);
    res.setHeader('Content-Length', buffer.length); // Ensure the buffer length is correctly sent
    res.send(buffer);
}

export const makeSessionDocuments = async (res, sessionId) => {
    // get schema from table Inscriptions, and list the fields
    const schema = await getAirtableSchema('Inscriptions');
    const inscTables = schema.filter(s => s.name === 'Inscriptions');
    const fields = inscTables[0].fields.map(f => f.name);
    logIfNotVercel("fields", fields)


    console.log(`Fetching inscriptions for session: ${sessionId}`);
    const inscriptions = await getAirtableRecords('Inscriptions', 'Grid view', 
        `AND(sessId='${sessionId}',{Statut}="Enregistrée")`
    );
    // logIfNotVercel("insc", inscriptions)
    if(!inscriptions || inscriptions.length === 0) {
        console.error('Failed to fetch inscriptions');
        return;
    }
    logIfNotVercel(`Fetched ${inscriptions.records.length} inscriptions`);
    // logIfNotVercel(inscriptions[0])

    const factureParams = documents.find(doc => doc.name === 'facture');
    const attestParams = documents.find(doc => doc.name === 'attestation');
    const certifParams = documents.find(doc => doc.name === 'certif_realisation');

    // fetch templates for all ofthe 3
    const factureTemplate = await fetchTemplate(GITHUBTEMPLATES + factureParams.template);
    const attestTemplate = await fetchTemplate(GITHUBTEMPLATES + attestParams.template);
    const certifTemplate = await fetchTemplate(GITHUBTEMPLATES + certifParams.template);
    // console.log(factureParams.dataPreprocessing)
    logIfNotVercel(factureParams)

    const idSession = inscriptions.records[0].id;

    let buffers = [];
    for (let i = 0; i < inscriptions.records.length; i++) {
        const inscription = inscriptions.records[i];
        logIfNotVercel(`Processing inscription: ${inscription.nom}`);
        
        let data = {...inscription};
        data = addMissingFields(fields, data);
        // logIfNotVercel("data", data)
        // logIfNotVercel("keys", Object.keys(data))
        if (factureParams.dataPreprocessing) {
            // logIfNotVercel('Preprocessing data...');
            factureParams.dataPreprocessing(data);
        }
        
        const buffer = await generateReport(
            factureTemplate,
            data,
        );
        // const buffer = await generateReportBuffer('test.docx', { Titre: 'Hello'+i });
        const filename =  sanitizeFileName(factureParams.titleForming(data)+".docx");

        // const filename = `file${i + 1}.docx`;
        logIfNotVercel(`Generated report for: ${filename}`);

        // TODO: update record in Airtable with facture date
        
        buffers.push({ filename:filename, content: buffer });
        if (attestParams.dataPreprocessing) {
            attestParams.dataPreprocessing(data);
        }
        const bufferAttest = await generateReport(attestTemplate, data);
        // const buffer = await generateReportBuffer('test.docx', { Titre: 'Hello'+i });
        let attestFilename =  sanitizeFileName(attestParams.titleForming(data)+".docx");
        // if(!attestFilename) {
        //     attestFilename = `attestation ${data["nom"]}.docx`;
        // }

        // const filename = `file${i + 1}.docx`;
        logIfNotVercel(`Generated report for: ${attestFilename}`);
        
        buffers.push({ filename:attestFilename, content: bufferAttest });

        if (certifParams.dataPreprocessing) {
            certifParams.dataPreprocessing(data);
        }

        const bufferCertif = await generateReport(certifTemplate, data);
        let certifFileName =  sanitizeFileName(certifParams.titleForming(data)+".docx");
        // if(!certifFileName) {
        //     certifFileName = `certification ${data["nom"]}.docx`;
        // }
        logIfNotVercel(`Generated report for: ${certifFileName}`);
        buffers.push({ filename:certifFileName, content: bufferCertif });
    }

    // TODO: add programme?
    const zipFileName = sanitizeFileName(`${idSession} Factures Attestations Certificats.zip`);
    logIfNotVercel(`Generating ZIP file: ${zipFileName}`);
    await generateAndSendZipReport(res, buffers, zipFileName);
    console.log('ZIP file generated and sent successfully');
}

