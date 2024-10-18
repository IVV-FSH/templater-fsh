import { createReport } from 'docx-templates';
import fs from 'fs';
import path from 'path';
import { getAirtableData, processMarkdownFields, getFrenchFormattedDate } from './utils.js';

async function generateReport() {
  const templatePath = path.join('templates', 'programme.docx');
  const template = fs.readFileSync(templatePath);

  const data = await getAirtableData("Sessions", "recAzC50Q7sCNzkcf");

  // Process specified fields with marked
  const fieldsToProcess = [
    'objectifs_fromprog', 
    'contenu_fromprog', 
    'methodespedago_fromprog', 
    'modaliteseval_fromprog', 
  ]; // Example fields
  const processedData = processMarkdownFields(data, fieldsToProcess);

  const buffer = await createReport({
    template,
    data: processedData
  });

  fs.writeFileSync(`reports/${getFrenchFormattedDate()} report.docx`, buffer);

  // Close the reading of programme.docx
  fs.closeSync(fs.openSync(templatePath, 'r'));
}

generateReport().catch(console.error);