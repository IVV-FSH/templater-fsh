import express from 'express';
import path from 'path';
import fs from 'fs';
import { getFrenchFormattedDate, fetchTemplate, generateReport, updateAirtableRecord, getAirtableSchema, processFieldsForDocx, getAirtableRecords, getAirtableRecord, ymd } from './utils.js';
import { put } from "@vercel/blob";
import { PassThrough } from 'stream';

// import { Stream } from 'stream';
import archiver from 'archiver';
const app = express();
app.use(express.json()); // For parsing application/json
app.use(express.urlencoded({ extended: true })); // For parsing application/x-www-form-urlencoded


app.get('/', (req, res) => {
  res.sendFile(path.join(process.cwd(), 'index.html'));
});

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
    await generateAndSendReport(
      'https://github.com/isadoravv/templater/raw/refs/heads/main/templates/catalogue.docx', 
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
    await generateAndSendReport(
      'https://github.com/isadoravv/templater/raw/refs/heads/main/templates/programme.docx', 
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
    await generateAndSendReport(
      'https://github.com/isadoravv/templater/raw/refs/heads/main/templates/devis.docx', 
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
  const table="tblxLakvfLieKsRyH"; // Inscriptions
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
        data["today"] = data["date_facture"];
        updatedInvoiceDate = true;
      }
    } else {
      console.log('Failed to retrieve data.');
      // broadcastLog('Failed to retrieve data.');
    }
    
    // Generate and send the report
    await generateAndSendReport(
    // await generateAndSendZipReport(
      'https://github.com/isadoravv/templater/raw/refs/heads/main/templates/facture.docx', 
      data, 
      res,
      `${data["id"]} ${data["nom"]} ${data["prenom"]}`
    );

    // TODO: update the record with the facture date
    if(!updatedInvoiceDate) {
      const updatedRecord = await updateAirtableRecord(table, recordId, { date_facture: new Date().toISOString() });
      if (updatedRecord) {
        console.log('Facture date updated successfully:', updatedRecord.id);
        // broadcastLog(`Facture date updated successfully: ${updatedRecord.id}`);
      } else {
        console.log('Failed to update facture date.');
        // broadcastLog('Failed to update facture date.');
      }
    }
    // res.render('index', { title: `Générer un Programme pour ${recordId}`, heading: 'Programme' });
  } catch (error) {
    console.error('Error:', error);
    // broadcastLog(`Error: ${error.message}`);
    res.status(500).json({ success: false, error: error.message });
  } 
  
});  
app.get('/realisation', async (req, res) => {
  // res.sendFile(path.join(process.cwd(), 'index.html'));
  const table="Inscriptions"; // Inscriptions
  // const table="tblxLakvfLieKsRyH"; // Inscriptions
  // const recordId="recdVx9WSFFeX5GP7";
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
    const newName = `Attestation de réalisation ${data["code_fromprog"]} ${new Date(data["au (from Session)"]).toLocaleDateString('fr-FR').replace(/\//g, '-')} - ${data["nom"]} ${data["prenom"]} ` || "err nom fact";

    // Generate and send the report
    await generateAndSendReport(
    // await generateAndSendZipReport(
      'https://github.com/isadoravv/templater/raw/refs/heads/main/templates/realisation.docx', 
      data, 
      res,
      newName
    );
    // res.render('index', { title: `Générer un Programme pour ${recordId}`, heading: 'Programme' });
  } catch (error) {
    console.error('Error:', error);
    // broadcastLog(`Error: ${error.message}`);
    res.status(500).json({ success: false, error: error.message });
  }
  
});  


app.get('/factures', async (req, res) => {
  const table = "tblxLakvfLieKsRyH";
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
      "https://github.com/isadoravv/templater/raw/refs/heads/main/templates/facture.docx",
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
});


// Reusable function to generate and send report
async function generateAndSendReport(url, data, res, fileName = "") {
  try {
    console.log('Generating report...');
    
    // Fetch the template and generate the report buffer
    const template = await fetchTemplate(url);
    const buffer = await generateReport(template, data); // This should return the correct binary buffer

    // Determine the file name
    const originalFileName = path.basename(url);
    const fileNameWithoutExt = originalFileName.replace(path.extname(originalFileName), '');
    let newTitle = fileName || fileNameWithoutExt;

    var newFileName = `${getFrenchFormattedDate()} ${newTitle}.docx`.replace(/[^a-zA-Z0-9-_.\s'éèêëàâîôùçãáÁÉÈÊËÀÂÎÔÙÇ]/g, '');
    var newFileName = newFileName.replace(/  /g, ' '); // Sanitize the filename
    // const newFileName = "file.docx"
    // Set the correct headers for file download and content type for .docx
    console.log("file name", newFileName)
    res.setHeader('Content-Disposition', `attachment; filename="${newFileName}"`);
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
});

