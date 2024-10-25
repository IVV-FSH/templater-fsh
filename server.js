import express from 'express';
import path from 'path';
import fs from 'fs';
import { getFrenchFormattedDate, fetchTemplate, generateReport, updateAirtableRecord, updateAirtableRecords, getAirtableSchema, processFieldsForDocx, getAirtableRecords, getAirtableRecord, ymd } from './utils.js';
import { put } from "@vercel/blob";
import { PassThrough } from 'stream';
import archiver from 'archiver';
// import { Stream } from 'stream';
import { GITHUBTEMPLATES } from './constants.js';

const app = express();

app.use(express.json()); // For parsing application/json
app.use(express.urlencoded({ extended: true })); // For parsing application/x-www-form-urlencoded


app.get('/', (req, res) => {
  res.sendFile(path.join(process.cwd(), 'index.html'));
});


const documents = [
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
    name: 'realisation',
    multipleRecords: false,
    titleForming: function(data) {
      return `Attestation de réalisation ${data["code_fromprog"]} ${ymd(data["au (from Session)"])} - ${data["nom"]} ${data["prenom"]}`;
    },
    template: 'attestation.docx',
    table: 'Inscriptions',
    queriedField: 'recordId',
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

const handleReportGeneration = async (req, res, document) => {
  try {
    let data;
    if(document.multipleRecords) {
      console.log(`Fetching multiple records from table: ${document.table}, view: ${document.view}, formula: ${document.formula}, sortField: ${document.sortField}, sortOrder: ${document.sortOrder}`);
      data = await getAirtableRecords(document.table, document.view, document.formula, document.sortField, document.sortOrder);
    } else {
      const { recordId } = req.query;
      if (document.queriedField && !recordId) {
        console.error('Paramètre recordId manquant.');
        return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
      }
      console.log(`Fetching single record from table: ${document.table}, recordId: ${recordId}`);
      data = await getAirtableRecord(document.table, recordId);
    }

    if (data) {
      console.log('Data successfully retrieved:', `${data.length} records`);
    } else {
      console.error('Failed to retrieve data.');
    }

    if(document.dataPreprocessing) {
      console.log('Preprocessing data...');
      document.dataPreprocessing(data);
    }

    // Generate and send the report
    console.log(`Generating report using template: ${document.template}`);
    await generateAndDownloadReport(
      `${GITHUBTEMPLATES}${document.template}`,
      data,
      res,
      document.titleForming(data)
    );

    if(document.airtableUpdatedData) {
      console.log('Updating Airtable record...');
      const updatedRecord = await updateAirtableRecord(document.table, recordId, document.airtableUpdatedData(data));
      if (updatedRecord) {
        console.log('Facture date updated successfully:', updatedRecord.id);
      } else {
        console.error('Failed to update facture date.');
      }
    }
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
}

// // Iterate over the documents array and create routes dynamically
// documents.forEach(doc => {
//   app.get(`/${doc.name}`, (req, res) => {
//     handleReportGeneration(req, res, doc.templateUrl, doc.processData, doc.generateFileName);
//   });
// });

app.get('/schemas', async (req, res) => {
  try {
    const schema = await getAirtableSchema();
    if (!schema) {
      console.log('Failed to retrieve schema.');
      return res.status(500).json({ success: false, error: 'Failed to retrieve schema.' });
    }
    // Extract only the names of the fields
    let mdFieldsSession = schema.find(table => table.name === 'Sessions').fields
    // let mdFieldsSession = schema.find(table => table.name === 'Sessions').fields
    .filter(field => {
      if (field.type === 'richText') {
        return field.name;
      } else if (field.type === 'multipleLookupValues') {
        if (field.options.result.type === 'richText') return field.name;
      } else {
        return null;
      }
    })
    .map(field => field.name); // Map to only field names
    
    res.json({champsMarkdown: mdFieldsSession});
  } catch (error) {
    console.error('Error fetching schema:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});



app.get('/catalogue', async (req, res) => {
  // res.sendFile(path.join(process.cwd(), 'index.html'));
  const table = "Sessions";
  const view = "Grid view";
  // const view = "Catalogue";
  var { annee } = req.query;

  if(!annee) { annee = new Date().getFullYear() + 1; }

  // const formula = `AND(OR({année}=${annee},{année}=""), OR(FIND(lieuxdemij_cumul,"iège"),FIND(lieuxdemij_cumul,"visio")))`; 
  const formula = `
  OR(
    AND(
        {année}="", 
        FIND(lieuxdemij_cumul,"intra")
    ),
    AND(
        {année}=2025, 
        FIND(lieuxdemij_cumul,"intra")=0
    )
  )`

  try {
    const data = await getAirtableRecords(table, view, formula, "du", "asc");
    if (data) {
      console.log('Data successfully retrieved:', data.records.length, "records");
      // broadcastLog(`Data successfully retrieved: ${data.records.length} records`);
    } else {
      console.log('Failed to retrieve data.');
      // broadcastLog('Failed to retrieve data.');
    }
    
    // Generate and send the report
    await generateAndDownloadReport(
      GITHUBTEMPLATES + 'catalogue.docx', 
      data, 
      res,
      'Catalogue des formations FSH ' + annee
    );
    // res.render('index', { title: 'Catalogue', heading: `Catalogue : à partir de ${table}/${view}` });
  } catch (error) {
    console.error('Error:', error);
    // broadcastLog(`Error: ${error.message}`);
    res.status(500).json({ success: false, error: error.message });
  }
  
});

app.get('/programme', async (req, res) => {
  // res.sendFile(path.join(process.cwd(), 'index.html'));
  const table="Sessions";
  // const recordId="recAzC50Q7sCNzkcf";
  const { recordId } = req.query;
  
  if (!recordId) {
    return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
  }
  try {
    const data = await getAirtableRecord(table, recordId);
    if (data) {
      console.log('Data successfully retrieved:', data.length);
      // broadcastLog(`Data successfully retrieved: ${data.length} records`);
    } else {
      console.log('Failed to retrieve data.');
      // broadcastLog('Failed to retrieve data.');
    }

    let newTitle = data["titre_fromprog"]
    if(data["du"] && data["au"]) { newTitle+= `${ymd(data["du"])}-${data["au"] && ymd(data["au"])}`}

    // Generate and send the report
    await generateAndDownloadReport(
      GITHUBTEMPLATES + 'programme.docx', 
      data, 
      res,
      newTitle || "err titre prog"
    );
    // res.render('index', { title: `Générer un Programme pour ${recordId}`, heading: 'Programme' });
  } catch (error) {
    console.error('Error:', error);
    // broadcastLog(`Error: ${error.message}`);
    res.status(500).json({ success: false, error: error.message });
  }
  
});  


app.get('/devis', async (req, res) => {
  // res.sendFile(path.join(process.cwd(), 'index.html'));
  const table="Devis";
  // const recordId="recAzC50Q7sCNzkcf";
  const { recordId } = req.query;
  
  if (!recordId) {
    return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
  }
  try {
    const data = await getAirtableRecord(table, recordId);
    if (data) {
      console.log('Data successfully retrieved:', data.length);
      // broadcastLog(`Data successfully retrieved: ${data.length} records`);
    } else {
      console.log('Failed to retrieve data.');
      // broadcastLog('Failed to retrieve data.');
    }

    let newTitle = `DEVIS FSH ${data["id"]} `
    // if(data["du"] && data["au"]) { newTitle+= `${ymd(data["du"])}-${data["au"] && ymd(data["au"])}`}
    
    // Generate and send the report
    await generateAndDownloadReport(
      GITHUBTEMPLATES + 'devis.docx', 
      data, 
      res,
      newTitle
    );
    // res.render('index', { title: `Générer un Programme pour ${recordId}`, heading: 'Programme' });
  } catch (error) {
    console.error('Error:', error);
    // broadcastLog(`Error: ${error.message}`);
    res.status(500).json({ success: false, error: error.message });
  }
  
});  


app.get('/facture', async (req, res) => {
  // res.sendFile(path.join(process.cwd(), 'index.html'));
  const table="Inscriptions"; // Inscriptions
  // const recordId="recdVx9WSFFeX5GP7";
  const { recordId } = req.query;

  var updatedInvoiceDate = false;
  
  if (!recordId) {
    return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
  }
  try {
    var data = await getAirtableRecord(table, recordId);
    if (data) {
      console.log('Data successfully retrieved:', data.length);
      // broadcastLog(`Data successfully retrieved: ${data.length} records`);
      if(data["date_facture"]) {
        data["today"] = new Date(data["date_facture"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' });
        updatedInvoiceDate = true;
      }
    } else {
      console.log('Failed to retrieve data.');
      // broadcastLog('Failed to retrieve data.');
    }

    // data['apaye'] = data.moyen_paiement && data.date_paiement;
    data['acquit'] = data["paye"].includes("Payé")
    ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
    : "";

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
    console.log("montant", data["montant"])

    // Generate and send the report
    await generateAndDownloadReport(
    // await generateAndSendZipReport(
      GITHUBTEMPLATES + 'facture.docx', 
      data, 
      res,
      `Facture ${data["id"]} ${data["nom"]} ${data["prenom"]}`
    );
    var updatedData = { 
        total: data['Montant'].toString()
      }
    // TODO: update the record with the facture date
    if(!updatedInvoiceDate) {
      updatedData["date_facture"] = new Date().toLocaleDateString('fr-CA');
    }
    const updatedRecord = await updateAirtableRecord(table, recordId, updatedData);
    if (updatedRecord) {
      console.log('Facture date updated successfully:', updatedRecord.id);
      // broadcastLog(`Facture date updated successfully: ${updatedRecord.id}`);
    } else {
      console.log('Failed to update facture date.');
      // broadcastLog('Failed to update facture date.');
    }
    // res.render('index', { title: `Générer un Programme pour ${recordId}`, heading: 'Programme' });
  } catch (error) {
    console.error('Error:', error);
    // broadcastLog(`Error: ${error.message}`);
    res.status(500).json({ success: false, error: error.message });
  } 
  
});  

app.get('/attestformation', async (req, res) => {
  const table = "Inscriptions";
  const { recordId } = req.query;

  if (!recordId) {
    return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
  }

  try {
    var data = await getAirtableRecord(table, recordId);
    if (data) {
      console.log('Data successfully retrieved:', data.length);
      // broadcastLog(`Data successfully retrieved: ${data.length} records`);
    } else {
      console.log('Failed to retrieve data.');
      // broadcastLog('Failed to retrieve data.');
    }
    
    // date de l'attestation au dernier jour de la formation
    data['today'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' })
    data['apaye'] = data.moyen_paiement && data.date_paiement;
    data['acquit'] = data["paye"].includes("Payé")
    ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
    : "";
    const newName = `Attestation de formation ${data["code_fromprog"]} ${new Date(data["au"]).toLocaleDateString('fr-FR').replace(/\//g, '-')} - ${data["nom"]} ${data["prenom"]}` || "err nom fact";
    // console.log('Data retrieved in /realisation:', data);
    // console.log('Code from prog:', data["code_fromprog"]);
    // console.log('Date from session:', data["au (from Session)"]);
    // console.log('New report name:', newName);
    
    // Generate and send the report
    await generateAndDownloadReport(
    // await generateAndSendZipReport(
      GITHUBTEMPLATES + 'attestation.docx', 
      data, 
      res,
      newName
    );
    // http://localhost:3000/realisation?recordId=rec9ZMibFvLaRuTm7

    // res.render('index', { title: `Générer un Programme pour ${recordId}`, heading: 'Programme' });
  } catch (error) {
    console.error('Error:', error);
    // broadcastLog(`Error: ${error.message}`);
    res.status(500).json({ success: false, error: error.message });
  } 
});


app.get('/realisation', async (req, res) => {
  const table = "Inscriptions";
  const { recordId } = req.query;

  if (!recordId) {
    return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
  }

  try {
    var data = await getAirtableRecord(table, recordId);
    if (data) {
      console.log('Data successfully retrieved:', data.length);
      // broadcastLog(`Data successfully retrieved: ${data.length} records`);
    } else {
      console.log('Failed to retrieve data.');
      // broadcastLog('Failed to retrieve data.');
    }
    console.log("data:NOM", data.nom)
    // date de l'attestation au dernier jour de la formation
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
    const newName = `Certificat de réalisation ${data["code_fromprog"]} ${new Date(data["au"]).toLocaleDateString('fr-FR').replace(/\//g, '-')} - ${data["nom"]} ${data["prenom"]}` || "err nom fact";
    // console.log('Data retrieved in /realisation:', data);
    // console.log('Code from prog:', data["code_fromprog"]);
    // console.log('Date from session:', data["au (from Session)"]);
    // console.log('New report name:', newName);
    
    // Generate and send the report
    await generateAndDownloadReport(
    // await generateAndSendZipReport(
      GITHUBTEMPLATES + 'certif_realisation.docx', 
      data, 
      res,
      newName
    );
    // http://localhost:3000/realisation?recordId=rec9ZMibFvLaRuTm7

    // res.render('index', { title: `Générer un Programme pour ${recordId}`, heading: 'Programme' });
  } catch (error) {
    console.error('Error:', error);
    // broadcastLog(`Error: ${error.message}`);
    res.status(500).json({ success: false, error: error.message });
  } 
});

app.get('/factures', async (req, res) => {
  const table = "Sessions";
  const sessionId = "recxEooSpjiO0qbvQ";
  // const { sessionId } = req.query

  const session = await getAirtableRecord("Sessions", sessionId);
  const inscrits = session["Inscrits"];

  var files = [];

  await Promise.all(inscrits.map(async id => {
    const data = await getAirtableRecord("tblxLakvfLieKsRyH", id);
    const fileName = `Facture ${data["id"]} ${data["nom"]} ${data["prenom"]}.docx`;
    console.log(data);
    const buffer = await generateReportBuffer(
      GITHUBTEMPLATES+"facture.docx",
      data
    );

    if (buffer) {
      files.push({ fileName, buffer });
    } else {
      console.error(`Invalid buffer for file: ${fileName}`);
    }
  }));

  if (files.length > 0) {
    await createZipArchive(files, res, 'all_reports.zip');
  } else {
    res.status(500).json({ success: false, error: 'No valid files to archive.' });
  }

  // TODO: update all records with the facture date
  // const updateAirtableRecords = async (table, recordIds, data) => {
  
});


// Reusable function to generate and send report
async function generateAndDownloadReport(url, data, res, fileName = "") {
  try {
    console.log('Generating report...');
    
    // Fetch the template and generate the report buffer
    const template = await fetchTemplate(url);
    const buffer = await generateReport(template, data); // This should return the correct binary buffer

    // Determine the file name
    const originalFileName = path.basename(url);
    const fileNameWithoutExt = originalFileName.replace(path.extname(originalFileName), '');
    let newTitle = fileName || fileNameWithoutExt;

    var newFileName = `${getFrenchFormattedDate()} ${newTitle}`.replace(/[^a-zA-Z0-9-_.\s'éèêëàâîôùçãáÁÉÈÊËÀÂÎÔÙÇ]/g, '');
    var newFileName = newFileName.replace(/  /g, ' '); // Sanitize the filename
    // const newFileName = "file.docx"
    // Set the correct headers for file download and content type for .docx
    const encodedFileName = encodeURIComponent(newFileName);

    console.log("file name", newFileName)
    res.setHeader('Content-Disposition', `attachment; filename="${encodedFileName}.docx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Length', buffer.length); // Ensure the buffer length is correctly sent

    // Send the buffer as a binary response
    res.end(buffer, 'binary');
    
    console.log('Report generated and sent as a download.');

  } catch (error) {
    console.error('Error generating report:', error);
    res.status(500).json({ success: false, error: error.message });
  }
}

/**
 * Fetches the template from the provided URL and generates a report buffer using the given data.
 * 
 * @async
 * @function generateReportBuffer
 * 
 * @param {string} url - The URL from which to fetch the document template (e.g., a .docx template).
 * @param {object} data - The data used to fill the template (e.g., dynamic content that will be inserted into the template).
 * 
 * @returns {Promise<Buffer>} - Returns a Promise that resolves to a Buffer containing the generated report.
 * 
 * @throws {Error} - Throws an error if fetching the template or generating the report fails.
 * 
 * @example
 * const buffer = await generateReportBuffer('https://example.com/template.docx', { title: 'Report 1' });
 */
async function generateReportBuffer(url, data) {
  try {
    const template = await fetchTemplate(url);
    const buffer = await generateReport(template, data);
    return buffer;
  } catch (error) {
    throw new Error(`Error generating report buffer: ${error.message}`);
  }
}

/**
 * Creates a zip archive from multiple file buffers and sends the zip to the client for download.
 * 
 * @async
 * @function createZipArchive
 * 
 * @param {Array<{ fileName: string, buffer: Buffer }>} files - An array of objects representing the files to be added to the zip archive. Each object should have a `fileName` (string) and a `buffer` (Buffer).
 * @param {object} res - The Express.js response object used to stream the zip file as a download to the client.
 * @param {string} [zipFileName="reports.zip"] - The name of the zip file to be sent for download (default: "reports.zip").
 * 
 * @returns {Promise<void>} - Returns a Promise that resolves when the zip archive is successfully created and sent.
 * 
 * @throws {Error} - Throws an error if creating the zip archive or streaming it to the response fails.
 */
async function createZipArchive(files, res, zipFileName = "reports.zip") {
  try {
    // Create a zip archive in memory
    const archive = archiver('zip', { zlib: { level: 9 } }); // Maximum compression

    // Prepare the response headers for sending the zip
    res.setHeader('Content-Disposition', `attachment; filename=${zipFileName}`);
    res.setHeader('Content-Type', 'application/zip');

    // Pipe the archive's output to the response
    archive.pipe(res);

    // Add each file buffer to the archive
    for (const { fileName, buffer } of files) {
      // Log the buffer to check its content
      if (!buffer || !(buffer instanceof Buffer)) {
        throw new Error(`Invalid buffer for file: ${fileName}`);
      }

      // Append the buffer to the zip archive
      archive.append(buffer, { name: fileName });
    }

    // Finalize the archive
    await archive.finalize();

    console.log('Zip archive created and sent.');
  } catch (error) {
    console.error(`Error creating zip archive: ${error.message}`);
    throw new Error(`Error creating zip archive: ${error.message}`);
  }
}



async function generateAndSendZipReport(url, data, res, fileName = "") {
  try {
    console.log('Generating report...');
    
    // Fetch the template and generate the report buffer
    const template = await fetchTemplate(url);
    const buffer = await generateReport({output: 'buffer', template, data});

    if (!Buffer.isBuffer(buffer)) {
      throw new Error('Generated document is not a valid Buffer.');
    }

    // Create a zip stream
    const zipStream = new PassThrough();
    const archive = archiver('zip', { zlib: { level: 9 } });

    // Set the appropriate headers for downloading the zip file
    res.setHeader('Content-Disposition', `attachment; filename="${getFrenchFormattedDate()} ${fileName || 'report'}.zip"`);
    res.setHeader('Content-Type', 'application/zip');

    // Pipe the archive to the response
    archive.pipe(zipStream);
    archive.append(buffer, { name: `${fileName || 'report'}.docx` });
    archive.finalize();

    // Pipe the zip stream to the response
    zipStream.pipe(res);

    console.log('Zip report generated and sent as a download.');
  } catch (error) {
    console.error('Error generating zip report:', error);
    res.status(500).json({ success: false, error: error.message });
  }
}

/**
 * Uploads the generated report to Vercel Blob Storage and returns the download URL.
 *
 * @async
 * @function uploadReportToBlobStorage
 * 
 * @param {string} fileName - The name of the file (including the extension) that will be stored in Blob Storage.
 * @param {Buffer} buffer - The generated report buffer that will be uploaded.
 * 
 * @returns {Promise<string>} - Returns the URL of the uploaded file.
 * 
 * @throws {Error} - Throws an error if uploading the file to Blob Storage fails.
 * 
 * @example
 * const downloadUrl = await uploadReportToBlobStorage('report.docx', reportBuffer);
 */
async function uploadReportToBlobStorage(fileName, buffer) {
  try {
    const { url } = await put(fileName, buffer, { access: 'public' });
    console.log(`File uploaded to Blob Storage: ${url}`);
    return url;
  } catch (error) {
    throw new Error(`Error uploading report to Blob Storage: ${error.message}`);
  }
}


// Start the server
const server = app.listen(process.env.PORT || 3000, () => {
  console.log(`Server is running on port http://localhost:${process.env.PORT || 3000}/`);
  console.log("TESTS FACTURE")
  console.log(`Impayée : http://localhost:${process.env.PORT || 3000}/facture?recordId=rechdhSdMTxoB8J1P`);
  console.log(`Rabais : http://localhost:${process.env.PORT || 3000}/facture?recordId=recFYDogCDybfujfd`);
  console.log(`Accomp : http://localhost:${process.env.PORT || 3000}/facture?recordId=recvrbZmRuUCgHrFK`);
  console.log(`Payée : http://localhost:${process.env.PORT || 3000}/facture?recordId=recLM3WRAiRYNPZ52`);
});

