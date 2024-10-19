import { createReport } from 'docx-templates';
import fs from 'fs';
import path from 'path';
import { getAirtableData, processFieldsForDocx, getFrenchFormattedDate, airtableMarkdownFields, getAirtableRecords } from './utils.js';

async function generateProg() {
  const templatePath = path.join('templates', 'cat2.docx');
  const template = fs.readFileSync(templatePath);

  const data = await getAirtableData("Sessions", "recAzC50Q7sCNzkcf");

  // const processedData = processFieldsForDocx(data, fieldsToProcess);

  const buffer = await createReport({
    template,
    data
  });

  fs.writeFileSync(`reports/${getFrenchFormattedDate()} report.docx`, buffer);

  // Close the reading of programme.docx
  fs.closeSync(fs.openSync(templatePath, 'r'));
}

async function generateCatalogue() {
  const templatePath = path.join('templates', 'catalogue.docx');
  const template = fs.readFileSync(templatePath);

  // const data = await getAirtableData("Sessions", null, "viwcgKhUjiPDkE0bs");
  const data = await getAirtableRecords("Sessions", "viwcgKhUjiPDkE0bs");
  // console.log(data["records"][0]);
  console.log(data["records"].length);

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

  fs.writeFileSync(`reports/${getFrenchFormattedDate()} catalogue-R.docx`, buffer);

  // Close the reading of programme.docx
  fs.closeSync(fs.openSync(templatePath, 'r'));
}

async function tests() {
  const templatePath = path.join('templates', 'cat2.docx');
  const template = fs.readFileSync(templatePath);

  // const data = await getAirtableData("Sessions", null, "Catalogue");
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

generateCatalogue().catch(console.error);