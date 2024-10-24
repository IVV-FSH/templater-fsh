import { createReport } from 'docx-templates';
import fs from 'fs';
import path from 'path';
import { getAirtableRecord, processFieldsForDocx, getFrenchFormattedDate, airtableMarkdownFields, getAirtableRecords, fetchTemplate, updateAirtableRecord } from './utils.js';
import archiver from 'archiver';

async function generateProg() {
  const templatePath = path.join('templates', 'cat2.docx');
  const template = fs.readFileSync(templatePath);

  const data = await getAirtableRecord("Sessions", "recAzC50Q7sCNzkcf");

  // const processedData = processFieldsForDocx(data, fieldsToProcess);

  const buffer = await createReport({
    template,
    data
  });

  fs.writeFileSync(`reports/${getFrenchFormattedDate()} report.docx`, buffer);

  // Close the reading of programme.docx
  fs.closeSync(fs.openSync(templatePath, 'r'));
}

async function facture(recordId = "recdVx9WSFFeX5GP7") {
  // const recordId = "recdVx9WSFFeX5GP7"; // payée
  const templatePath = path.join('templates', 'facture.docx');
  const template = fs.readFileSync(templatePath);
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
  
  const data = await getAirtableRecord("Inscriptions", recordId);
  // console.log(data)
  data['apaye'] = data.moyen_paiement && data.date_paiement;
  data['acquit'] = data.moyen_paiement && data.date_paiement
  ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
  : "";
  data["Montant"] = calculateCost(data)
  console.log("Montant calc", data["Montant"])
  data['montant'] = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(
    parseFloat(data["Montant"]),
  );  
  console.log("montant", data["montant"])

  // const processedData = processFieldsForDocx(data, fieldsToProcess);

  const buffer = await createReport({
    template,
    data
  });

  const factId = `${data["id"]} ${data["nom"]}` || "err nom fact";

  fs.writeFileSync(`reports/${getFrenchFormattedDate()} Facture ${factId}.docx`, buffer);

  // Close the reading of programme.docx
  fs.closeSync(fs.openSync(templatePath, 'r'));
}

facture().catch(console.error); // payée
facture("recpy3eggrFMnAXT4").catch(console.error); // impayée


async function updateARec() {
  const table = "Inscriptions";
  const recordId = "rechdhSdMTxoB8J1P"; // payée
  await updateAirtableRecord(table, recordId, { "date_facture": "2024-10-24" });
}

// updateARec().catch(console.error);

async function real() {
  const recordId = "recdVx9WSFFeX5GP7"; // payée
  const templatePath = path.join('templates', 'realisation.docx');
  // const template = fs.readFileSync(templatePath);
  const template = await fetchTemplate("https://github.com/isadoravv/templater/raw/refs/heads/main/templates/realisation.docx")

  const data = await getAirtableRecord("Inscriptions", recordId);

  // const processedData = processFieldsForDocx(data, fieldsToProcess);

  const buffer = await createReport({
    template,
    data
  });

  const newName = `Attestation de réalisation ${data["code_fromprog"]} ${new Date(data["au (from Session)"]).toLocaleDateString('fr-FR').replace(/\//g, '-')} - ${data["nom"]} ${data["prenom"]} ` || "err nom fact";

  fs.writeFileSync(`reports/${getFrenchFormattedDate()} ${newName}.docx`, buffer);

  // Close the reading of programme.docx
  // fs.closeSync(fs.openSync(templatePath, 'r'));
}

async function devis() {
  const recordId = "recGb3crKcmXQ2caL"; // payée
  const templatePath = path.join('templates', 'devis.docx');
  const template = fs.readFileSync(templatePath);

  const data = await getAirtableRecord("Devis", recordId);

  // const processedData = processFieldsForDocx(data, fieldsToProcess);

  const buffer = await createReport({
    template,
    data
  });
  let newTitle = `DEVIS FSH ${data["id"]} `

  // const factId = `${data["id"]} ${data["nom"]}` || "err nom fact";

  fs.writeFileSync(`reports/${getFrenchFormattedDate()} ${newTitle}.docx`, buffer);

  // Close the reading of programme.docx
  fs.closeSync(fs.openSync(templatePath, 'r'));
}

async function generateCatalogue() {
  const templatePath = path.join('templates', 'catalogue.docx');
  const template = fs.readFileSync(templatePath);

  const table = "Sessions";
  const view = "Catalogue";
  const data = await getAirtableRecords(table, view);
  // console.log(data["records"][0]);
  if (data) {
    console.log('Data successfully retrieved:', data.records.length, "records");
    // broadcastLog(`Data successfully retrieved: ${data.records.length} records`);
  } else {
    console.log('Failed to retrieve data.');
    // broadcastLog('Failed to retrieve data.');
  }
  
  // console.log(data);
  // let processedData = data.records.map(record => record.fields);
  // console.log(processedData);
  // processedData = processedData.map(record => processFieldsForDocx(record, airtableMarkdownFields));
  // console.log(processedData);

  // const processedData = processFieldsForDocx(data, fieldsToProcess);
  // console.log(data);

  const buffer = await createReport({
    template,
    // data: {records:data}
    data
  });

  fs.writeFileSync(`reports/${getFrenchFormattedDate()} Catalogue des Formations FSH 2025.docx`, buffer);

  // Close the reading of programme.docx
  fs.closeSync(fs.openSync(templatePath, 'r'));
}

async function tests() {
  const templatePath = path.join('templates', 'cat2.docx');
  const template = fs.readFileSync(templatePath);

  // const data = await getAirtableRecord("Sessions", null, "Catalogue");
  const data = {
    "records": [
      {
        "id": "rec1",
        "createdTime": "2024-10-03T09:29:20.000Z",
        "fields": {
          "markdownField": "# Heading\n\nThis is a **bold** text.",
          "multipleLookupField": ["Value1", "Value2", "Value3"],
          "shortStringField": "\"Short string 1\"",
          "markdownArrayField": [
            "# First Heading\n\nThis is the first **bold** text.",
            "## Second Heading\n\nThis is the second *italic* text.",
            "### Third Heading\n\nThis is the third [link](https://example.com)."
          ]
        }
      },
      {
        "id": "rec2",
        "createdTime": "2024-09-27T14:09:34.000Z",
        "fields": {
          "markdownField": "## Subheading\n\nThis is an *italic* text.",
          "multipleLookupField": ["ValueA", "ValueB"],
          "shortStringField": "\"Short string 2\"",
          "markdownArrayField": [
            "# First Heading\n\nThis is the first **bold** text.",
            "## Second Heading\n\nThis is the second *italic* text.",
            "### Third Heading\n\nThis is the third [link](https://example.com)."
          ]
        }
      },
      {
        "id": "rec3",
        "createdTime": "2024-08-15T11:45:00.000Z",
        "fields": {
          "markdownField": "### Another Heading\n\nThis is a [link](https://example.com).",
          "multipleLookupField": ["ValueX", "ValueY", "ValueZ"],
          "shortStringField": "\"Short string 3\"",
          "markdownArrayField": [
            "# First Heading\n\nThis is the first **bold** text.",
            "## Second Heading\n\nThis is the second *italic* text.",
            "### Third Heading\n\nThis is the third [link](https://example.com)."
          ]
        }
      }
    ]
  };
  
  console.log(data);
  let processedData = data.records.map(record => record.fields);
  console.log(processedData);
  processedData = processedData.map(record => processFieldsForDocx(record, airtableMarkdownFields));
  console.log(processedData);
  // console.log({data: {records:processedData}});

  const buffer = await createReport({
    template,
    data: {records:processedData}
  });

  fs.writeFileSync(`reports/${getFrenchFormattedDate()} cat2-R.docx`, buffer);

  // Close the reading of programme.docx
  fs.closeSync(fs.openSync(templatePath, 'r'));
}

/**
 * Zips several files from the file system into a single .zip file.
 */
function zipLocalFiles() {
  const output = fs.createWriteStream('reports/test-files.zip');
  const archive = archiver('zip', { zlib: { level: 9 } });

  output.on('close', function () {
    console.log(`Zipped ${archive.pointer()} total bytes`);
  });

  archive.on('error', function (err) {
    throw err;
  });

  archive.pipe(output);

  // Add multiple local files to the zip
  const files = ['test1.docx', 'test2.docx'].map(file => path.join('reports', file));
  files.forEach(file => archive.file(file, { name: path.basename(file) }));

  archive.finalize();
}

// Call this function to test local zipping of files
// zipLocalFiles();

/**
 * Zips several dynamically created files (from Buffers) into a single .zip file.
 */
async function zipFilesFromBuffers() {
  const output = fs.createWriteStream('reports/test-buffers.zip');
  const archive = archiver('zip', { zlib: { level: 9 } });

  output.on('close', function () {
    console.log(`Zipped ${archive.pointer()} total bytes`);
  });

  archive.on('error', function (err) {
    throw err;
  });

  archive.pipe(output);

  // Data for filling the template
  const files = [
    { data: { Titre: "abc" }, title: "abc.docx" },
    { data: { Titre: "xyz" }, title: "xyz.docx" }
  ];

  // Template path
  const templatePath = path.join('templates', 'test.docx');
  const template = fs.readFileSync(templatePath);

  // Generate report buffers for each file and add to zip
  for (const file of files) {
    const buffer = await createReport({
      template,
      data: file.data
    });
    archive.append(buffer, { name: file.title });
  }

  await archive.finalize();
}

// Call this function to test zipping of files from Buffers
// zipFilesFromBuffers();

function zipTwoFiles() { // this works.
  const output = fs.createWriteStream(path.join('reports', 'test-archive.zip'));
  const archive = archiver('zip', { zlib: { level: 9 } }); // Set compression level

  output.on('close', function() {
    console.log(`Archive created successfully! Total size: ${archive.pointer()} bytes.`);
  });

  archive.on('error', function(err) {
    throw new Error(`Error creating zip archive: ${err.message}`);
  });

  // Pipe the archive to the output file
  archive.pipe(output);

  // Append both files to the zip
  archive.file(path.join('templates', 'test.docx'), { name: 'test.docx' });
  archive.file(path.join('templates', 'html.docx'), { name: 'html.docx' });

  // Finalize the archive (this will actually create the zip)
  archive.finalize();
}

// Test the zipping function
// zipTwoFiles();

// devis().catch(console.error);




// Function to generate a .docx on the fly and zip it
async function generateAndZipFile() {
  const dateFormatted = getFrenchFormattedDate();
  const output = fs.createWriteStream(path.join('reports', `${dateFormatted} archive.zip`));
  const archive = archiver('zip', { zlib: { level: 9 } });

  output.on('close', function() {
    console.log(`Archive created successfully! Total size: ${archive.pointer()} bytes.`);
  });

  archive.on('error', function(err) {
    throw new Error(`Error creating zip archive: ${err.message}`);
  });

  archive.pipe(output);

  const templatePath = path.join('templates', 'test.docx');
  const template = fs.readFileSync(templatePath);

  if (!Buffer.isBuffer(template)) {
    throw new Error('Template is not a valid Buffer.');
  }
  console.log('Template read successfully.');

  const data = { 'Titre': 'Dynamic Title' };

  try {
    // Generate the .docx buffer directly
    const docxBuffer = await createReport({
      output: 'buffer',
      template,
      data
    });

    // Log the size and type of the generated buffer
    console.log(`Generated document buffer size: ${docxBuffer.length} bytes`);
    console.log(`Generated document type: ${typeof docxBuffer}`);

    // Check if docxBuffer is valid
    if (!Buffer.isBuffer(docxBuffer)) {
      throw new Error('Generated document is not a valid Buffer.');
    }

    // Append the generated file to the zip
    archive.append(docxBuffer, { name: 'generated-file.docx' });

    // Finalize the archive (this will actually create the zip)
    await archive.finalize();
  } catch (error) {
    throw new Error(`Error creating zip archive: ${error.message}`);
  }
}


// Test the function
// generateAndZipFile().catch(err => console.error(err));