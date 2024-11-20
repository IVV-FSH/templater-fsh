import express from 'express';
import path from 'path';
import fs from 'fs';
import { getFrenchFormattedDate, fetchTemplate, generateReport, updateAirtableRecord, updateAirtableRecords, getAirtableSchema, processFieldsForDocx, getAirtableRecords, getAirtableRecord, ymd, sendConfirmation, sendConfirmationToAllSession, serveConvocationPage } from './utils.js';
import { put } from "@vercel/blob";
import { PassThrough } from 'stream';
import archiver from 'archiver';
import nodemailer from 'nodemailer';
// import { Stream } from 'stream';
import { GITHUBTEMPLATES } from './constants.js';
import { downloadDocxBuffer, makeGroupFacture, makeSessionDocuments, documents, makeConvention, generateAndSendZipReport } from './documents.js';
import {createReport} from 'docx-templates';
import { processImports } from './dups.js';
import dotenv from 'dotenv';
import envoiNdf from './ndf.js';
import { getBesoins } from './besoins.js';

dotenv.config();

const app = express();

app.use(express.json()); // For parsing application/json
app.use(express.urlencoded({ extended: true })); // For parsing application/x-www-form-urlencoded

// Get the directory name using import.meta.url

app.get('/', (req, res) => {
  res.sendFile(path.join(process.cwd(), 'index.html'));
});

app.get('/ndf', async (req, res) => {
  await envoiNdf();
  res.status(200).json({ success: true, message: 'NDFs sent successfully.' });
});

app.get('/duplicates', async (req, res) => {
  try {
    await processImports();
    res.status(200).json({ success: true, message: 'Duplicates processed successfully.' });
  } catch (error) {
    console.error('Error processing duplicates:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/confirm', async (req, res) => {
  const { inscriptionId } = req.query;
  try {
    await sendConfirmation(inscriptionId);
    res.status(200).json({ success: true, message: 'Confirmation sent successfully.' });
  } catch (error) {
    console.error('Error sending confirmation:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/convoc', async (req, res) => {
  const { inscriptionId } = req.query;
  try {
    const record = await getAirtableRecord('Inscriptions', inscriptionId);
    // if (!record) {
    //   console.log(`Record with id ${inscriptionId} not found`);
    //   return;
    // }
  
    // Check if envoi_convocation is empty and necessary fields are available
    const { envoi_convocation, 
      prenom, nom, mail, 
      titre_fromprog, 
      adresses_intra, nb_adresses, 
      lieux, 
      fillout_recueil, du,
      prerequis_fromprog,
      public_fromprog,
      introcontexte_fromprog,
      contenu_fromprog,
      objectifs_fromprog,
      methodespedago_fromprog,
      modaliteseval_fromprog,
      Formateurice,
      dates, 
      sessId
    } = record;
  
    let str_lieu = '';
    console.log(`Record ${inscriptionId} has ${nb_adresses} addresses`);
    if (parseInt(nb_adresses) == 1) {
      console.log(`Record ${inscriptionId} has one address`);
      if (lieux.includes("isioconf") || lieux.join("").includes("isioconf")) {
        str_lieu = "en visioconférence (le lien de connexion vous sera envoyé prochainement)";
      } else if (lieux.includes("iège") || lieux.includes("iège")) {
        str_lieu = "au siège de la FSH, 6 rue du Chemin vert, 75011 Paris";
      } else if (lieux.includes("intra") || lieux.includes("intra")) {
        str_lieu = `à l'adresse : ${adresses_intra}`;
      }
    } else if (parseInt(nb_adresses) > 1) {
      var leslieux = lieux.map((lieu, index) => {
        if (lieu.includes("isioconf")) {
          return "en visioconférence (le lien de connexion vous sera envoyé prochainement)";
        } else if (lieu.includes("iège")) {
          return "au siège de la FSH, 6 rue du Chemin vert, 75011 Paris";
        } else if (lieu.includes("intra")) {
          return `à l'adresse : ${adresses_intra}`;
        }
      });
      str_lieu = leslieux.join(" et ");
            
    }

    const completionDate = new Date(du);
    completionDate.setDate(completionDate.getDate() - 15);
    const completionDateString = completionDate.toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' });
  
  
    await serveConvocationPage(
      res,
      prenom, 
      nom, 
      mail, 
      prerequis_fromprog,
      public_fromprog,
      titre_fromprog, 
      introcontexte_fromprog,
      contenu_fromprog,
      objectifs_fromprog,
      methodespedago_fromprog,
      modaliteseval_fromprog,
      Formateurice,
      dates, 
      str_lieu, 
      fillout_recueil, 
      completionDateString, 
      sessId,
      inscriptionId
    )
    // await sendConfirmation(inscriptionId);
    // res.status(200).json({ success: true, message: 'Confirmation sent successfully.' });
  } catch (error) {
    console.error('Error sending confirmation:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/updateDateConvoc', async (req, res) => {
  const { inscriptionId } = req.query;
  try {
    const today = new Date();
    const formattedDate = today.toISOString().split('T')[0]; // Extract the date part in YYYY-MM-DD format
    const updatedRecord = await updateAirtableRecord('Inscriptions', inscriptionId, { envoi_convocation: formattedDate });
    if (updatedRecord) {
      console.log('Convocation date updated successfully:', updatedRecord.id);
    } else {
      console.error('Failed to update convocation date.');
    }
    res.status(200).json({ success: true, message: 'Convocation date updated successfully.' });
  } catch (error) {
    console.error('Error updating convocation date:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/confirmForSession', async (req, res) => {
  const { sessionId } = req.query;
  try {
    await sendConfirmationToAllSession(sessionId);
    res.status(200).json({ success: true, message: 'Confirmations sent successfully.' });
  } catch (error) {
    console.error('Error sending confirmations:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/session', async (req, res) => {
  const { sessId, formateurId } = req.query;

  const sessionData = await getAirtableRecord("Sessions", sessId);
  const titre = sessionData["titre_fromprog"] || "";
  const datesSession = sessionData["dates"] || "";
  const inscriptionsData = await getAirtableRecords("Inscriptions", "Grid view", `sessId="${sessId}"`, "nom", "asc");
  var inscritsHtml = "";
  if(!inscriptionsData.records || inscriptionsData.records.length === 0) {
    res.send("Aucun inscrit pour cette session");
  }
  inscriptionsData.records.forEach(inscrit => {
    inscritsHtml += `<tr>`;
    inscritsHtml += `<td>${inscrit["prenom"]} ${inscrit["nom"]}</td>`;
    inscritsHtml += `<td>${inscrit["poste"]}</td>`;
    inscritsHtml += `<td>${inscrit["entite"]}</td>`;
    inscritsHtml += `</tr>`;
  });
  var inscritsHtml = `<table>
  <thead>
    <tr>
      <th>Prénom Nom</th>
      <th>Poste</th>
      <th>Entité</th>
    </tr>
  </thead>
  <tbody>${inscritsHtml}</tbody>
  </table>`;
  const besoinsData = await getAirtableRecords("Recueil des besoins", "Grid view", `sessId="${sessId}"`, "nom (from Inscrits)", "asc");

  const besoinsHtml = await getBesoins(besoinsData);

  const nbTotalInscritsEnreg = inscriptionsData.records.filter(insc => insc.Statut == "Enregistrée").length;
  const filledBesoins = besoinsData.records.filter(besoin => besoin.rempli=="🟢");
  const nbFilledBesoins = filledBesoins.length;

  var resHtml = `
      <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${titre} ${datesSession}</title>
    <style>
      body { font-family: 'Segoe UI', sans-serif; }
.fiche-besoin {
  margin-bottom: 20px;
  padding: 15px;
  border: 1px solid #ccc;
  border-radius: 8px; /* Rounded corners */
  box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); /* Subtle shadow */
  background-color: #fff; /* White background */
  transition: box-shadow 0.3s ease; /* Smooth transition for hover effect */
}
  .grid-container {
  display: grid;
  grid-template-columns: repeat(2, 1fr); /* 2 columns */
  gap: 20px; /* Space between grid items */
}

@media (max-width: 768px) {
  .grid-container {
    grid-template-columns: 1fr; /* 1 column on mobile */
  }
}
        .filled { color: green; }
      .not-filled { color: red; }
      .question { font-weight: bold; }
      .answer { color: black; }
      .font-light { font-weight: lighter; }

      table {
  border-collapse: collapse; /* Collapse borders */
}

th, td {
  border: 1px solid #333; /* Add a border to table cells */
  padding: 8px; /* Add padding to table cells */
  text-align: left; /* Align text to the left */
}

th {
  background-color: #f2f2f2; /* Add a background color to table headers */
}
    </style>
  </head>
  <h1>Formation : <span>${titre}</span>  <span class="font-light">${datesSession}</span></h1>
  `
  if(inscriptionsData.records.length > 0) {
    resHtml += `<h2>Inscrits (${nbTotalInscritsEnreg})</h2>`;
    resHtml += inscritsHtml;
  } 
  if(besoinsHtml === "") {
    resHtml += `
  <h2>Recueil des besoins <span class="font-light">(${nbFilledBesoins} remplis/${nbTotalInscritsEnreg} inscrits)</span></h2>
  <p>Aucun besoin n'a été rempli pour cette session</p>
      `
    } else {

  resHtml += `
  <h2>Recueil des besoins <span class="font-light">(${nbFilledBesoins} remplis/${nbTotalInscritsEnreg} inscrits)</span></h2>
<div class="grid-container">${besoinsHtml.html}</div>
  <div class="stats">
    <h3>Récapitulatif des besoins</h3>
    ${besoinsHtml.recap}
  </div>
    `

    res.send(resHtml+"</html>");
  }
  // const besoins = await getAirtableRecords(table, "Grid view", `sessId="${sessId}"`, "id", "asc");
  // if(!besoins.records || (besoins.records && besoins.records.length === 0)) {
  //   res.send("Aucun besoin n'a été rempli pour cette session");
  // } else {
  //   const type = besoins.records[0].Type

  // }
});

// dynamycally create routes for each document
documents.forEach(doc => {
  app.get(`/make/${doc.name}`, async (req, res) => {
    await handleReportGeneration(req, res, doc);
  });
  if(doc.documentField) {
    app.get(`/email/${doc.name}`, async (req, res) => {
      await handleReportGenerationAndSendEmail(req, res, doc);
      // TODO: in documents.js, add documentField

    });
  }
});

const handleReportGeneration = async (req, res, document) => {
  console.log('Generating report...', document.name);
  try {
    let data;
    const { recordId } = req.query;
    if(document.multipleRecords) {
      console.log(`Fetching multiple records from table: ${document.table}, view: ${document.view}, formula: ${document.formula}, sortField: ${document.sortField}, sortOrder: ${document.sortOrder}`);
      data = await getAirtableRecords(document.table, document.view, document.formula, document.sortField, document.sortOrder);
    } else {
      if (document.queriedField && !recordId) {
        console.error('Paramètre recordId manquant.');
        return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
      }
      console.log(`Fetching single record from table: ${document.table}, recordId: ${recordId}`);
      data = await getAirtableRecord(document.table, recordId);
    }

    if (data) {
      console.log('Data successfully retrieved:', `${document.multipleRecords ? data.records.length : data.length } records`);
    } else {
      console.error('Failed to retrieve data.');
    }

    if(document.dataPreprocessing) {
      console.log('Preprocessing data...');
      if(document.name === "facture_grp") {
        data = await document.dataPreprocessing(data, recordId);
      } else {
        data = await document.dataPreprocessing(data);
      }
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

const handleReportGenerationAndSendEmail = async (req, res, document) => {
  console.log('Generating report...', document.name);
  try {
    let data;
    const { recordId } = req.query;
    const email = "isadora.vuongvan@sante-habitat.org";
    // if (!email) {
      // return res.status(400).json({ success: false, error: 'Email parameter is missing.' });
    // }

    if (document.multipleRecords) {
      console.log(`Fetching multiple records from table: ${document.table}, view: ${document.view}, formula: ${document.formula}, sortField: ${document.sortField}, sortOrder: ${document.sortOrder}`);
      data = await getAirtableRecords(document.table, document.view, document.formula, document.sortField, document.sortOrder);
    } else {
      if (document.queriedField && !recordId) {
        console.error('Paramètre recordId manquant.');
        return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
      }
      console.log(`Fetching single record from table: ${document.table}, recordId: ${recordId}`);
      data = await getAirtableRecord(document.table, recordId);
    }

    if (data) {
      console.log('Data successfully retrieved:', `${document.multipleRecords ? data.records.length : data.length} records`);
    } else {
      console.error('Failed to retrieve data.');
    }

    if (document.dataPreprocessing) {
      console.log('Preprocessing data...');
      if (document.name === "facture_grp") {
        data = await document.dataPreprocessing(data, recordId);
      } else {
        data = await document.dataPreprocessing(data);
      }
    }

    // Generate the report
    console.log(`Generating report using template: ${document.template}`);
    const buffer = await generateReportBuffer(
      `${GITHUBTEMPLATES}${document.template}`,
      data
    );

    // Send the report by email
    console.log(`Sending report to email: ${email}`);
    await sendEmailWithAttachment(email, document.titleForming(data), buffer);

    // FIX:
    //   console.log('Updating Airtable record...');
    //   const updatedRecord = await updateAirtableRecord(document.table, recordId, document.airtableUpdatedData(data));
    //   if (updatedRecord) {
    //     console.log('Facture date updated successfully:', updatedRecord.id);
    //   } else {
    //     console.error('Failed to update facture date.');
    //   }
    if (document.airtableUpdatedData) {
      await updateAirtableRecord(document, data);
    }

    res.status(200).json({ success: true, message: 'Report generated and sent by email successfully.' });
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
};

const sendEmailWithAttachment = async (
  toEmail = "isadora.vuongvan@sante-habitat.org", fileName, buffer, 
  subject = 'Votre document est prêt', 
  text = 'Veuillez trouver en pièce jointe le rapport.') => {
  // Créer un objet transporteur en utilisant le transport SMTP par défaut
  const transporter = nodemailer.createTransport({
    host: 'smtp-declic-php5.alwaysdata.net',
    port: 465,
    secure: true, // true pour 465, false pour les autres ports
    auth: {
      user: 'isadora.vuongvan@sante-habitat.org', // Votre adresse email
      pass: process.env.FSH_PASSWORD // Votre mot de passe email
    }
  });

  // Configurer les données de l'email avec des symboles unicode
  const mailOptions = {
    from: '"Isadora Vuong Van - FSH" <isadora.vuongvan@sante-habitat.org>', // Adresse de l'expéditeur
    to: toEmail, // Liste des destinataires
    subject: subject, // Objet de l'email
    text: text, // Corps du texte en clair
    attachments: [
      {
        filename: `${fileName}.docx`,
        content: buffer
      }
    ]
  };

  // Envoyer l'email avec l'objet de transport défini
  await transporter.sendMail(mailOptions);
};

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

app.get("/factures_sess", async (req, res) => {
  // const template = await fetchTemplate(GITHUBTEMPLATES + 'test.docx');
  // await makeSessionDocuments(res, '2410CVS7CE');
  try {
    var { sessionId } = req.query;
    if(!sessionId) {
      return res.status(400).json({ success: false, error: 'Paramètre sessionId manquant' });
    }
    await makeSessionDocuments(res, sessionId);
    // res.status(200).json({ success: true, message: `Documents générés pour la session ${sessionId}` });
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
  // await makeSessionDocuments(res, 'recxEooSpjiO0qbvQ');
});

app.get('/test', async (req, res) => {
  // await makeGroupFacture(res, 'recvu2Muu5y0XAwWY');
  // var { factureId } = req.query;
  // await makeGroupFacture(res, factureId);
  // const factureTemplate = await fetchTemplate(GITHUBTEMPLATES + 'facture.docx');
  // const devisTemplate = await fetchTemplate(GITHUBTEMPLATES + 'devis.docx');
  console.log('Fetching test template...');
  const testTemplate = await fetchTemplate("https://github.com/IVV-FSH/templater-fsh/raw/refs/heads/main/z_testsdocx/test.docx");
  console.log('Test template fetched successfully.');

  const data = [
    { titre: 'Titre 1' },
    { titre: 'Titre 2' },
    { titre: 'Titre 3' },
    { titre: 'Titre 4' }
  ];
  let buffers = [];
  for(let i = 0; i < data.length; i++) {
    try {
      console.log(`Generating report for data index ${i}...`);
      const buffer = await createReport({
        output: 'buffer',
        template: testTemplate,
        data: data[i]
      });
      const fileName = `Test${i + 1}.docx`;
      buffers.push({ filename: fileName, content: buffer });
      console.log(`Report generated for data index ${i}: ${fileName}`);
    } catch (error) {
      console.error(`Error generating report for data index ${i}:`, error);
      const errorBuffer = Buffer.from(`Error generating report for data index ${i}: ${error.message}`);
      const errorFileName = `Error${i + 1}.txt`;
      buffers.push({ filename: errorFileName, content: errorBuffer });
    }
  }
  const zipfileName = encodeURIComponent('Test.zip');
  console.log('Generating and sending zip report...');
  await generateAndSendZipReport(
    res,
    buffers,
    zipfileName
  );
  console.log('Zip report generated and sent successfully.');
});

app.get('/catalogue', async (req, res) => {
  // res.sendFile(path.join(process.cwd(), 'index.html'));
  const table = "Sessions";
  const view = "Grid view";
  // const view = "Catalogue";
  var { annee } = req.query;

  if(!annee) { annee = new Date().getFullYear() + 1; }

  const formula = `OR(AND(OR({année}=${annee},{année}=""), OR(REGEX_MATCH(lieuxdemij_cumul&"","iège"),REGEX_MATCH(lieuxdemij_cumul&"","isio"))), AND({année}&""="",lieuxdemij_cumul&""=""))`; 
  // const formula = `
  // OR(
  //   AND(
  //       {année}="", 
  //       FIND(lieuxdemij_cumul,"intra")
  //   ),
  //   AND(
  //       {année}=${annee}, 
  //       FIND(lieuxdemij_cumul,"intra")=0
  //   )
  // )`

  try {
    var data = await getAirtableRecords(table, view, formula, "du", "asc");
    if (data) {
      console.log('Data successfully retrieved:', data.records.length, "records");
      // broadcastLog(`Data successfully retrieved: ${data.records.length} records`);
    } else {
      console.log('Failed to retrieve data.');
      // broadcastLog('Failed to retrieve data.');
    }

    // make a set of all titre_fromprog
    // var set = new Set();
    // data.records.forEach(record => {
    //   set.add(record["titre_fromprog"]);
    // });
    // var recap = Array.from(set);
    // recap = recap.map((titre, index) => {
    //   const records = data.records.filter(record => record["titre_fromprog"] === titre);
    //   return {
    //     titre_fromprog: titre,
    //     dates_lieux: records.map(record => `${record["dates"]} ${record["lieuxdemij_cumul"] ? record["lieuxdemij_cumul"].join('').toLowerCase() : ''}`),
    //     nb_dates: new Set(records.map(record => record["dates"])).size,
    //     prixadh: records.map(record => record["prixadh"]),
    //     prixnonadh: records.map(record => record["prixnonadh"]),
    //     ouvertepersaccomp_fromprog: records.map(record => record["ouvertepersaccomp_fromprog"]),
    //   }
    // });

    // console.log('Recap:', recap);
    // data.recap = recap;
    
    
    // Generate and send the report
    await generateAndDownloadReport(
      GITHUBTEMPLATES + 'catalogue.docx', 
      data, 
      res,
      'Catalogue des formations FSH ' + annee
    );
    // res.status(200).json({ success: true, message: `Catalogue généré pour l'année ${annee}` });
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

    let newTitle = `FSH ${data["Name"]} `
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

app.get('/emargement', async (req, res) => {
  const { sessionId } = req.query;
  // const sessionId = "recxEooSpjiO0qbvQ"
  if(!sessionId) {
    return res.status(400).json({ success: false, error: 'Paramètre sessionId manquant.' });
  }
  const documentData = documents.find(doc => doc.name === "emargement");
  var session = await getAirtableRecord("Sessions", sessionId);
  // console.log(`Session: ${session["recordId"]} ${session["code_fromprog"]}`);

  const processed = await documentData.dataPreprocessing(session);

  // const templatePath = path.join('templates', 'emargement.docx');
  // const template = fs.readFileSync(templatePath)
  // const docxBuffer = await createReport({
  //   output: 'buffer',
  //   template,
  //   data:processed
  // });

  // fs.writeFileSync(`reports/${getFrenchFormattedDate()} Emargement ${session["code_fromprog"]}.docx`, docxBuffer);

  // downloadDocxBuffer(res,`Emargement ${session["code_fromprog"]}.docx`, docxBuffer );
  // Generate and send the report
  await generateAndDownloadReport(
    GITHUBTEMPLATES + documentData.template,
    processed,
    res,
    documentData.titleForming(processed)
  );

});

app.get('/convention', async (req, res) => {
  const { recordId } = req.query;
  // const recordId = "recGb3crKcmXQ2caL"
  if(!recordId) {
    return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
  }
  const documentData = documents.find(doc => doc.name === "convention");
  var session = await getAirtableRecord(documentData.table, recordId);
  // console.log(`Session: ${session["recordId"]} ${session["code_fromprog"]}`);

  const processed = await documentData.dataPreprocessing(session);

  // Generate and send the report
  await generateAndDownloadReport(
    GITHUBTEMPLATES + documentData.template,
    processed,
    res,
    documentData.titleForming(processed)
  );

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
        data["today"] = new Date(data["date_facture"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'});
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

app.get('/h', async (req, res) => {
  // makeSessionDocuments(res, 'recxEooSpjiO0qbvQ');
  // const conv = await makeConvention("recGb3crKcmXQ2caL");
  // downloadDocxBuffer(res, conv.filename,conv.content );
});

app.get('/facture_grp', async (req, res) => {
  const documentData = documents.find(doc => doc.name === "facture_grp");
  try {
    var { factureId } = req.query;
    if(!factureId) {
      return res.status(400).json({ success: false, error: 'Paramètre factureId manquant.' });
    }
    var data = await getAirtableRecord("Factures-Devis-Conventions", factureId);
    // const facture = await makeGroupFacture(factureId);
    data = await documentData.dataPreprocessing(data, factureId);
    await generateAndDownloadReport(
      GITHUBTEMPLATES + documentData.template,
      data,
      res,
      documentData.titleForming(data)
    );
    // if (facture) {
    //   console.log('Document successfully created:', facture.length);
    //   // broadcastLog(`Data successfully retrieved: ${data.length} records`);
    // } else {
    //   console.log('Failed to create document.');
    //   // broadcastLog('Failed to retrieve data.');
    // }
    // downloadDocxBuffer(res, facture.filename, facture.content);
  } catch (error) {
    console.error('Error:', error);
    // broadcastLog(`Error: ${error.message}`);
    res.status(500).json({ success: false, error: error.message });
  } 
  
});  

// app.get('/attestformation', async (req, res) => {
//   const table = "Inscriptions";
//   const { recordId } = req.query;

//   if (!recordId) {
//     return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
//   }

//   try {
//     var data = await getAirtableRecord(table, recordId);
//     if (data) {
//       console.log('Data successfully retrieved:', data.length);
//       // broadcastLog(`Data successfully retrieved: ${data.length} records`);
//     } else {
//       console.log('Failed to retrieve data.');
//       // broadcastLog('Failed to retrieve data.');
//     }
    
//     // date de l'attestation au dernier jour de la formation
//     data['today'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
//     data['apaye'] = data.moyen_paiement && data.date_paiement;
//     data['acquit'] = data["paye"].includes("Payé")
//     ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
//     : "";
//     const newName = `Attestation de formation ${data["code_fromprog"]} ${new Date(data["au"]).toLocaleDateString('fr-FR').replace(/\//g, '-')} - ${data["nom"]} ${data["prenom"]}` || "err nom fact";
//     // console.log('Data retrieved in /realisation:', data);
//     // console.log('Code from prog:', data["code_fromprog"]);
//     // console.log('Date from session:', data["au (from Session)"]);
//     // console.log('New report name:', newName);
    
//     // Generate and send the report
//     await generateAndDownloadReport(
//     // await generateAndSendZipReport(
//       GITHUBTEMPLATES + 'attestation.docx', 
//       data, 
//       res,
//       newName
//     );
//     // http://localhost:3000/realisation?recordId=rec9ZMibFvLaRuTm7

//     // res.render('index', { title: `Générer un Programme pour ${recordId}`, heading: 'Programme' });
//   } catch (error) {
//     console.error('Error:', error);
//     // broadcastLog(`Error: ${error.message}`);
//     res.status(500).json({ success: false, error: error.message });
//   } 
// });


// app.get('/realisation', async (req, res) => {
//   const table = "Inscriptions";
//   const { recordId } = req.query;

//   if (!recordId) {
//     return res.status(400).json({ success: false, error: 'Paramètre recordId manquant.' });
//   }

//   try {
//     var data = await getAirtableRecord(table, recordId);
//     if (data) {
//       console.log('Data successfully retrieved:', data.length);
//       // broadcastLog(`Data successfully retrieved: ${data.length} records`);
//     } else {
//       console.log('Failed to retrieve data.');
//       // broadcastLog('Failed to retrieve data.');
//     }
//     console.log("data:NOM", data.nom)
//     // date de l'attestation au dernier jour de la formation
//     data['du'] = new Date(data["du"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
//     data['au'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
//     data['dureeh_fromprog'] = data["dureeh_fromprog"]/3600;
//     data["assiduite"] = data["assiduite"] * 100;
//     data['nom'] = Array.isArray(data["nom"]) ? data["nom"][0].toUpperCase() : data["nom"].toUpperCase();
//     data['today'] = new Date(data["au"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
//     data['apaye'] = data.moyen_paiement && data.date_paiement;
//     data['acquit'] = data["paye"].includes("Payé")
//     ? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
//     : "";
//     const newName = `Certificat de réalisation ${data["code_fromprog"]} ${new Date(data["au"]).toLocaleDateString('fr-FR').replace(/\//g, '-')} - ${data["nom"]} ${data["prenom"]}` || "err nom fact";
//     // console.log('Data retrieved in /realisation:', data);
//     // console.log('Code from prog:', data["code_fromprog"]);
//     // console.log('Date from session:', data["au (from Session)"]);
//     // console.log('New report name:', newName);
    
//     // Generate and send the report
//     await generateAndDownloadReport(
//     // await generateAndSendZipReport(
//       GITHUBTEMPLATES + 'certif_realisation.docx', 
//       data, 
//       res,
//       newName
//     );
//     // http://localhost:3000/realisation?recordId=rec9ZMibFvLaRuTm7

//     // res.render('index', { title: `Générer un Programme pour ${recordId}`, heading: 'Programme' });
//   } catch (error) {
//     console.error('Error:', error);
//     // broadcastLog(`Error: ${error.message}`);
//     res.status(500).json({ success: false, error: error.message });
//   } 
// });

// app.get('/factures', async (req, res) => {
//   const table = "Sessions";
//   const sessionId = "recxEooSpjiO0qbvQ";
//   // const { sessionId } = req.query

//   const session = await getAirtableRecord("Sessions", sessionId);
//   const inscrits = session["Inscrits"];

//   var files = [];

//   await Promise.all(inscrits.map(async id => {
//     const data = await getAirtableRecord("tblxLakvfLieKsRyH", id);
//     const fileName = `Facture ${data["id"]} ${data["nom"]} ${data["prenom"]}.docx`;
//     console.log(data);
//     const buffer = await generateReportBuffer(
//       GITHUBTEMPLATES+"facture.docx",
//       data
//     );

//     if (buffer) {
//       files.push({ fileName, buffer });
//     } else {
//       console.error(`Invalid buffer for file: ${fileName}`);
//     }
//   }));

//   if (files.length > 0) {
//     await createZipArchive(files, res, 'all_reports.zip');
//   } else {
//     res.status(500).json({ success: false, error: 'No valid files to archive.' });
//   }

//   // TODO: update all records with the facture date
//   // const updateAirtableRecords = async (table, recordIds, data) => {
  
// });


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

    var newFileName = `${getFrenchFormattedDate(false)} ${newTitle}`.replace(/[^a-zA-Z0-9-_.\s'éèêëàâîôùçãáÁÉÈÊËÀÂÎÔÙÇ]/g, '');
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



async function generateAndSendZipReport2(url, data, res, fileName = "") {
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
  for (const doc of documents) {
    console.log(doc.name.toUpperCase())
    if(doc.examples) {
      doc.examples.forEach(example => {
        console.log(`http://localhost:${process.env.PORT || 3000}/make/${doc.name}?recordId=${example.recordId}`, example.desc);
      });
    } else {
    console.log(`http://localhost:${process.env.PORT || 3000}/make/${doc.name}`);
    }
  }
  console.log("Documents session ------")
  for (const doc of 
    [
      {recordId:"recXcDJ6nYYJYsVmV", desc:"SM14JV"}
    ]
  ) {
    console.log(`http://localhost:${process.env.PORT || 3000}/factures_sess?sessionId=${doc.recordId}`, doc.desc);
  }
  console.log("TESTS SESSION")
  console.log(`http://localhost:${process.env.PORT || 3000}/test`);
});

