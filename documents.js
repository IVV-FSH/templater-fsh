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
        name: 'catalogue', // OK
        multipleRecords: true,
        // formula: `AND(OR({année}=2025,{année}=""), OR(FIND(lieuxdemij_cumul,"iège"),FIND(lieuxdemij_cumul,"visio")))`,
        formula: `OR(
                        AND(
                            {année}="", 
                            FIND(lieuxdemij_cumul,"intra")
                        ),
                        AND(
                            {année}=${new Date().getFullYear() + 1}, 
                            FIND(lieuxdemij_cumul,"intra")=0
                        )
                    )`,
        sortField: 'du',
        sortOrder: 'asc',
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
        name: 'programme', // OK + testé en ligne
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
        examples: [
            {recordId:"recvP7YK3HUSxWZqV", desc:"sans dates"},
            {recordId:"recTgirR6imNqtFAs", desc:"avec dates, visio"},
            {recordId:"rec8XSuWVObX1R6pW", desc:"ouverte pers accomp, siège"}
        ],
    },
    {
        name: 'devis',
        multipleRecords: false,
        queriedField: 'recordId',
        titleForming: function(data) {
            return `${ymd(new Date(data["date_doc"]))} Devis FSH ${data['CODE'].replace("[","").replace("]","")} ${data["entite"]}`;
        },
        template: 'devis.docx',
        table: 'Factures-Devis-Conventions',
        dataPreprocessing: async function(data) {
            const prog = await getAirtableRecord("Programme", data.progId);
            // merge prog and data, but if there are conflicts, data wins
            data.entite = data.entite_sgtr || data.entite_dmdr || data.entite || "";
            data.rue = data.rue_sgtr || data.rue_dmdr || data.rue || "";
            data.cp = data.cp_sgtr || data.cp_dmdr || data.cp || "";
            data.ville = data.ville_sgtr || data.ville_dmdr || data.ville || "";
            data.poste = data.poste_sgtr || data.poste_dmdr || data.poste || "";

            data['duree_horaires'] = data["duree_h"]/3600;
            data = {...prog, ...data};
            data["id"] = `${ymd(new Date(data["date_doc"]))} ${data['CODE'].replace("[","").replace("]","").trim()}-${String(data.entite).substring(0, 3)}`;            // console.log("data", data)
            // for each Tarifs, retrieve record from Tarifs
            // Tarifs: [ 'recVLy02tLFlINZOo', 'recoAZnkTqeEVujSD' ],
            console.log("data.annee_envisagee:", data.annee_envisagee);

            let tarifs = [];
            for (let i = 0; i < data.Tarifs.length; i++) {
                const tarif = await getAirtableRecord("Tarifs", data.Tarifs[i]);
                // console.log("tarif", tarif) 
                tarifs.push(tarif);
            }
            data.tarifs = tarifs;

            let tarif = tarifs.length > 0 ? tarifs.filter(t => {
                console.log("Comparing:", parseInt(t['Année']), "with", parseInt(data.annee_envisagee));
                return parseInt(t['Année']) == parseInt(data.annee_envisagee);
            })[0] : null;

            if (tarif) {
                data['prixintra'] = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(tarif["prix_intra"]);
            } else {
                data['prixintra'] = NaN;
            }
            // console.log("tarifs", tarifs)
            return data;
        },
        examples: [{recordId:"recWMivigueMuraHe", desc:"guilchard"},],
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
                data["today"] = new Date(data["date_facture"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'});
            }
            data['acquit'] = data["paye"].includes("Payé")
            ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
            : "";

            data["Montant"] = calculTotalPrixInscription(data)
            logIfNotVercel("Montant calc", data["Montant"])
            data['montant'] = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(
                parseFloat(data["Montant"]),
            );  

            return data;
            
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
        },
        examples: [{recordId:"recOjSqWKF5VVudrL", desc:"payée"},{recordId:"rec1Mu1M2papdWAmq", desc:"non payée"},],
        documentField: "doc_facture",
    },
    {
        name: 'facture_grp',
        multipleRecords: false,
        titleForming: function(data) {
            return `${data["Name"]}`;
        },
        template: 'facture_grp.docx',
        table: 'Factures-Devis-Conventions',
        queriedField: 'factureId',
        dataPreprocessing: async function(data, factureId) {
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

            data.entite = data.entite_sgtr || data.entite_dmdr || data.entite || "";
            data.rue = data.rue_sgtr || data.rue_dmdr || data.rue || "";
            data.cp = data.cp_sgtr || data.cp_dmdr || data.cp || "";
            data.ville = data.ville_sgtr || data.ville_dmdr || data.ville || "";
        
            if(data.lieuxdemij_cumul.includes("En intra")) {
                // dont make stagiaires liste.
                // total = data.prixintra_fromsess
                data['stagiaires'] = [];
                data.montant = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(
                    parseFloat(data.prixintra_fromsess),
                );
                data.nb_pax = data.nb_pax > 0 ? data.nb_pax : "A déterminer";
            } else {
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
            
                data['stagiaires'] = [...stagiaires];
                data.montant = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(
                    parseFloat(total),
                );
            }
            return data;
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
            data['Formateurice'] = Array.isArray(data["Formateurice"]) ? data["Formateurice"].join(", ").replace(/"/g, '') : data["Formateurice"].replace(/"/g, '');
            data['du'] = new Date(data["du"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            data['au'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            // data['dureeh_fromprog'] = data["dureeh_fromprog"]/3600;
            // data["assiduite"] = data["assiduite"] * 100;
            data['nom'] = Array.isArray(data["nom"]) ? data["nom"][0].toUpperCase() : data["nom"].toUpperCase();
            data['today'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            data['apaye'] = data.moyen_paiement && data.date_paiement;
            data['acquit'] = data["paye"].includes("Payé")
            ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
            : "";

            return data;
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
            data['du'] = new Date(data["du"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            data['au'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            // data['dureeh_fromprog'] = data["dureeh_fromprog"]/3600;
            // data["assiduite"] = data["assiduite"] * 100;
            data['nom'] = Array.isArray(data["nom"]) ? data["nom"][0].toUpperCase() : data["nom"].toUpperCase();
            data['today'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            data['apaye'] = data.moyen_paiement && data.date_paiement;
            data['acquit'] = data["paye"].includes("Payé")
            ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
            : "";
            return data;
        }
    },
    {
        name: 'emargement',
        multipleRecords: false,
        titleForming: function(data) {
            return `Feuille d'émargement ${data["code_fromprog"]} ${ymd(data["du"])}`;
        },
        template: 'emargement.docx',
        table: 'Sessions',
        queriedField: 'recordId',
        dataPreprocessing: async function(data) {
            data['du'] = new Date(data["du"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            data['au'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            const demij = await getAirtableRecords("Demi-journées", "Grid view", `sessId="${data.recordid}"`);
            const stagiaires = await getAirtableRecords("Inscriptions", "Grid view", `AND(sessId="${data.recordid}",{Statut}="Enregistrée")`);
            // data['demijournees'] = demij.records.map(d => {
            //     const horaires = {
            //         date: new Date(d.debut).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'}),
            //         horaires: `de ${new Date(d.debut).toLocaleTimeString('fr-FR')} à ${new Date(d.fin).toLocaleTimeString('fr-FR')}`
            //     }
            //     return {dates:`${horaires.date}, ${horaires.horaires}`};
            // });
            // from demij, make journées
            // make a set of debut dates
            const debutDates = new Set(demij.records.map(d => new Date(d.debut).toLocaleDateString('fr-FR', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric',timeZone: 'Europe/Paris' })));
            // make an object that has, for each date, matin and apres-midi (if they exist in demij)
            const getTimeOfDay = (dateString) => {
                // console.log("dateString", dateString)
                const date = new Date(dateString);
                const options = { hour: 'numeric', hour12: false, timeZoneName: 'short' };
                const localHour = new Intl.DateTimeFormat('fr-FR', options).format(date);
                // console.log("localHour", localHour)
                return parseInt(localHour);
            };
            data['jrs'] = [...debutDates]
            const journees = [...debutDates].map(date => {
                const demijForDate = demij.records.filter(d => new Date(d.debut).toLocaleDateString('fr-FR', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris' }) === date).sort((a, b) => new Date(a.debut) - new Date(b.debut));
                const matin = demijForDate[0];
                const apresmidi = demijForDate[1];
                return {
                    date: date,
                    matin: matin ? `de ${new Date(matin.debut).toLocaleTimeString('fr-FR', { hour: 'numeric', minute: 'numeric' }).replace(':', 'h')} à ${new Date(matin.fin).toLocaleTimeString('fr-FR', { hour: 'numeric', minute: 'numeric' }).replace(':', 'h')}` : "",
                    apresmidi: apresmidi ? `de ${new Date(apresmidi.debut).toLocaleTimeString('fr-FR', { hour: 'numeric', minute: 'numeric' }).replace(':', 'h')} à ${new Date(apresmidi.fin).toLocaleTimeString('fr-FR', { hour: 'numeric', minute: 'numeric' }).replace(':', 'h')}` : "",
                }
            });
            data['journees'] = journees;
            data['stagiaires'] = stagiaires && stagiaires.records && stagiaires.records.length > 0 ? stagiaires.records.map(s => {
                return {
                    nom: (s.nom[0]).toUpperCase(),
                    prenom: s.prenom[0],
                }
            }).sort((a,b) => {
                if(a.nom < b.nom) { return -1; }
                if(a.nom > b.nom) { return 1; }
                return 0;
            }) : [];
            // console.log("data", data)

            return data;

        },
        examples: [{recordId:"recXcDJ6nYYJYsVmV", desc:"SM14JV que aprem"},],
    },
    {
        name: 'convention',
        multipleRecords: false,
        titleForming: function(data) {
            return `Convention professionnelle de formation ${data.nom_fromdesti} ${data.ville_fromlieu}`;
        },
        template: 'convention.docx',
        table: 'Factures-Devis-Conventions',
        queriedField: 'recordId',
        dataPreprocessing: async function(data) {
            const sessId = data["sessId"]
            const session = await getAirtableRecord("Sessions", sessId);
            data = {...session, ...data};
            data['SIRET'] = data["SIRET"] || null;
            data['du'] = new Date(data["du"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            data['au'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
            data['dureeh_fromprog'] = data["dureeh_fromprog"]/3600;
            data.entite = data.entite_sgtr || data.entite_dmdr || data.entite || "";
            data.rue = data.rue_sgtr || data.rue_dmdr || data.rue || "";
            data.cp = data.cp_sgtr || data.cp_dmdr || data.cp || "";
            data.ville = data.ville_sgtr || data.ville_dmdr || data.ville || "";
            data.poste = data.poste_sgtr || data.poste_dmdr || data.poste || "";
            data.montant = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(
                parseFloat(data.prixintra_fromsess),
            );
            // const demij = await getAirtableRecords("Demi-journées", "Grid view", `sessId="${data.recordid}"`);
            const stagiaires = await getAirtableRecords("Inscriptions", "Grid view", `AND(sessId="${sessId}",{Statut}="Enregistrée")`);
            // make a set of debut dates
            data['stagiaires'] = stagiaires && stagiaires.records && stagiaires.records.length > 0 ? stagiaires.records.map(s => {
                return {
                    nom: (s.nom[0]).toUpperCase(),
                    prenom: s.prenom[0],
                    poste: (s.poste && s.poste[0]) || "",
                }
            }).sort((a,b) => {
                if(a.nom < b.nom) { return -1; }
                if(a.nom > b.nom) { return 1; }
                return 0;
            }) : [];
            // console.log("data", data)

            return data;

        },
        examples: [{recordId:"recrzATVTK5NtDfDP", desc:"FMDC besoins remplis + liste pax"},],
    },
    // {
    //     name: 'factures',
    //     multipleRecords: true,
    //     titleForming: function(data) {
    //         return `Facture ${data["id"]} ${data["nom"]} ${data["prenom"]}`;
    //     },
    //     template: 'facture.docx',
    //     table: 'Inscriptions',
    //     queriedField: 'recordId',
    // }
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
    // add header for size
    // res.setHeader('Content-Length', archive.pointer());
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

export const makeFeuilleEmargement = async (sessionId) => {
    const emargementParams = documents.find(doc => doc.name === 'emargement');
    console.log(`Fetching session: ${sessionId}`);
    let data = await getAirtableRecord(emargementParams.table, sessionId);
    if(!data) {
        console.error('Failed to fetch session');
        return;
    }
    if (emargementParams.dataPreprocessing) {
        console.log('Preprocessing data...');
        data = await emargementParams.dataPreprocessing(data);
    }
    const buffer = await generateReportBuffer(emargementParams.template, data);
    const filename = sanitizeFileName(getFrenchFormattedDate()+" "+emargementParams.titleForming(data)+".docx");
    console.log(`Generated report for: ${filename}`);
    return { filename:filename, content: buffer };
}

export const makeConvention = async (devisId) => {
    const conventionParams = documents.find(doc => doc.name === 'convention');
    console.log(`Fetching session: ${devisId}`);
    let data = await getAirtableRecord(conventionParams.table, devisId);
    if(!data) {
        console.error('Failed to fetch session');
        return;
    }
    if (conventionParams.dataPreprocessing) {
        console.log('Preprocessing data...');
        data = await conventionParams.dataPreprocessing(data);
    }
    const buffer = await generateReportBuffer(conventionParams.template, data);
    const filename = sanitizeFileName(getFrenchFormattedDate()+" "+conventionParams.titleForming(data)+".docx");
    console.log(`Generated report for: ${filename}`);
    return { filename:filename, content: buffer };
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
        data["today"] = new Date().toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'});
        // logIfNotVercel("data", data)
        // logIfNotVercel("keys", Object.keys(data))
        if (factureParams.dataPreprocessing) {
            // logIfNotVercel('Preprocessing data...');
            data = factureParams.dataPreprocessing(data);
        }
        console.log("data for facture", data)
        
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
            data = attestParams.dataPreprocessing(data);
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
            data = certifParams.dataPreprocessing(data);
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

