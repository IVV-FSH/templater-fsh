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