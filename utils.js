import fs from 'fs';
import { GITHUBTEMPLATES } from './constants.js';
import path from 'path';
import { createReport } from 'docx-templates';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import axios from 'axios';
import { marked } from 'marked';
import moment from 'moment';
import nodemailer from 'nodemailer';
import mjml from 'mjml';
import puppeteer from 'puppeteer';
import { calculTotalPrixInscription } from './documents.js';
import mjml2html from 'mjml';
import { documents } from './documents.js';
// import { broadcastLog } from './server.js';

dotenv.config();
const AIRTABLE_API_KEY = process.env.AIRTABLE_TOKEN; // Get API key from environment variable
const AIRTABLE_BASE_ID = 'appK5MDuerTOMig1H'; // Replace with your Airtable Base ID
const AUTH_HEADERS = {
	Authorization: `Bearer ${AIRTABLE_API_KEY}`,
};

const mjmlApiApplicationId = process.env.MJML_API_APPLICATION_ID;
const mjmlApiPublicKey = process.env.MJML_API_PUBLIC_KEY;
const mjmlApiSecretKey = process.env.MJML_API_SECRET_KEY;

/**
* Encodes an Airtable formula for use in URLs.
*
* This function takes an Airtable formula string and encodes it using `encodeURIComponent`
* to ensure it can be safely included in a URL. It also handles field names with special characters
* or spaces by wrapping them in curly braces.
*
* @param {string} formula - The Airtable formula to encode.
* @returns {string} The encoded formula.
*
* @example
* // Example Airtable formulas:
* // Formula to filter records where the "Name" field is "John Doe"
* const formula1 = "{Name}='John Doe'";
* console.log(formulaFilter(formula1)); // Output: "%7BName%7D%3D'John%20Doe'"
*
* // Formula to filter records where the "Age" field is greater than 30
* const formula2 = "{Age}>30";
* console.log(formulaFilter(formula2)); // Output: "%7BAge%7D%3E30"
*
* // Formula to filter records where the "Name" field is "John Doe" and "Age" field is greater than 30
* const formula3 = "AND({Name}='John Doe', {Age}>30)";
* console.log(formulaFilter(formula3)); // Output: "AND(%7BName%7D%3D'John%20Doe'%2C%20%7BAge%7D%3E30)"
*
* // Formula to filter records where the "Status" field is "Active" or "Age" field is less than 25
* const formula4 = "OR({Status}='Active', {Age}<25)";
* console.log(formulaFilter(formula4)); // Output: "OR(%7BStatus%7D%3D'Active'%2C%20%7BAge%7D%3C25)"
*/
function formulaFilter(formula) {
	return encodeURIComponent(formula);
}

export const airtableMarkdownFields = [
	"modalitesacces_fromprog",
	"modaliteseval_fromprog",
	"methodespedago_fromprog",
	"contenu_fromprog",
	"objectifs_fromprog",
	"introcontexte_fromprog",
	"Intro/Contexte",
	"Contenu",
	"Objectifs",
	"Méthodes pédagogiques",
	"Modalités d’évaluation",
	"Modalités d'accès",
	"Modalités de certification",
	"markdownField",
	"markdownArrayField",
	"objectifs",
	"contenu",
	"methodespedago",
	"modaliteseval",
	"modalitesacces",
	"modalitescertif",
	"introcontexte",
]

export const getAirtableSchema = async (table) => {
	let url = `https://api.airtable.com/v0/meta/bases/${AIRTABLE_BASE_ID}/tables`;
	
	const response = await axios.get(url, {
		headers: AUTH_HEADERS,
	});

	const tables = response.data.tables;
	// console.log("Retrieved tables from Airtable:", tables.map(t => t.name));

	return tables;
}

/**
 * Adds missing fields to an object.
 *
 * This function takes an array of field names and an object containing fields with missing values.
 * It returns a new object where each field from the array is present. If a field is missing in the
 * input object, it will be set to empty string in the result.
 *
 * @param {string[]} allFieldsFromTable - An array of field names to be included in the result.
 * @param {Object} fieldsWithMissing - An object containing fields with their values. Missing fields will be set to empty string.
 * @returns {Promise<Object>} A promise that resolves to an object with all fields from the array, with missing fields set to empty string.
 */
export const addMissingFields = (allFieldsFromTable, fieldsWithMissing) => {
    let res = {};
	try {

		allFieldsFromTable.forEach((fn) => {
			if (fieldsWithMissing[fn]) {
				res[fn] = fieldsWithMissing[fn];
				// TODO: if an array, if empty, set to empty array
			} else {
				res[fn] = "";
			}
		});
		return res;
	} catch (error) {
		console.error(error);
		return fieldsWithMissing;
		throw new Error("Error adding missing fields to object.");
	}
};

/**
 * Updates a record in an Airtable table.
 *
 * This function sends a PATCH request to the Airtable API to update a specific record in the specified table.
 *
 * @param {string} table - The name of the Airtable table.
 * @param {string} recordId - The ID of the record to update.
 * @param {Object} data - The data to update in the record.
 * @returns {Promise<Object>} A promise that resolves to the updated record data.
 */
export const updateAirtableRecord = async (table, recordId, data) => {
	try {
		let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${table}/${recordId}`;
		console.log(`Updating record at URL: ${url}`);
		console.log(`Data being sent: ${JSON.stringify(data)}`);

		const response = await axios.patch(url, { fields: data }, {
			headers: {
				...AUTH_HEADERS,
				'Content-Type': 'application/json'
			}
		});

		// console.log(`Response from Airtable: ${JSON.stringify(response.data)}`);
		return response.data;
	} catch (error) {
		console.error(`Error updating record in Airtable: ${error.message}`);
		throw new Error("Error updating record in Airtable.", error);
	}
};

/**
 * Updates multiple records in an Airtable table.
 *
 * This function sends a PATCH request to the Airtable API to update multiple records in the specified table.
 *
 * @param {string} table - The name of the Airtable table.
 * @param {Array<Object>} records - An array of objects, each containing the record ID and the data to update.
 * @returns {Promise<Object>} A promise that resolves to the updated records data.
 */
export const updateAirtableRecords = async (table, records) => {
	try {
		let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${table}`;
		console.log(`Updating records at URL: ${url}`);
		console.log(`Data being sent: ${JSON.stringify(records)}`);

		const response = await axios.patch(url, { records }, {
			headers: {
				...AUTH_HEADERS,
				'Content-Type': 'application/json'
			}
		});

		// console.log(`Response from Airtable: ${JSON.stringify(response.data)}`);
		return response.data;
	} catch (error) {
		console.error(`Error updating records in Airtable: ${error.message}`);
		throw new Error("Error updating records in Airtable.", error);
	}
};

export const getAirtableRecord = async (table, recordId) => {
    try {
        let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${table}/${recordId}`;

        const response = await axios.get(url, {
            headers: AUTH_HEADERS,
        });

        const sch = await getAirtableSchema(table);
        const allFields = sch.filter(t => t.name == table)[0].fields.map(t => t.name);

        // console.log(response.data.fields);
        const withAllFields = await addMissingFields(allFields, response.data.fields);
        // console.log(withAllFields);

        // let processedData = recordId
        //     ? response.data.fields
        //     : response.data.records;

        let processedData = processFieldsForDocx(
            withAllFields,
            airtableMarkdownFields,
        );
		processedData['today'] = new Date().toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})
        return processedData;
    } catch (error) {
        console.error(error);
        throw new Error("Error retrieving data from Airtable.");
    }
};
export const getAirtableRecords = async (table, view = null, formula = null, sortField = null, sortDir = null) => {
	try {
	  let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${table}`;
	  const params = [];
	  if (formula) {
		params.push(`filterByFormula=${formulaFilter(formula)}`);
	  }
	  if (view) {
		params.push(`view=${encodeURIComponent(view)}`);
	  }
	  if (sortField) {
		let sortParam = `sort%5B0%5D%5Bfield%5D=${sortField}`;
		if (sortDir && ['asc', 'desc'].includes(sortDir)) {
			sortParam += `&sort%5B0%5D%5Bdirection%5D=${sortDir}`;
		}
		params.push(sortParam);
	  }
	  if (params.length > 0) {
		url += `?${params.join('&')}`;
	  }
	  console.log(`Fetching records from URL: ${url}`);
	  // broadcastLog(`Fetching records from URL: ${url}`); 
	
	  const response = await axios.get(url, {
		headers: AUTH_HEADERS,
	  });
  
	//   if lenth of records is 0, return empty array
	  if(response.data.records.length == 0) {
		return [];
	  }
	  let processedData = response.data.records.map(record => record.fields);
	//   console.log(`Fetched records: ${JSON.stringify(processedData)}`);
	//   processedData = processedData.map(record => {
	// 	if (Array.isArray(record)) {
	// 	  return record.join(', ');
	// 	}
	// 	return record;
	//   });
  
	processedData = processedData.map(record => processFieldsForDocx(
		record,
		airtableMarkdownFields
	));
	processedData['today'] = new Date().toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})

	return { records: processedData };
	} catch (error) {
	  console.error(error);
	  throw new Error('Error retrieving data from Airtable.');
	}
  };

/**
 * Processes fields for a DOCX document by modifying the provided data object.
 * 
 * - If a field is a string and is enclosed in double quotes, the quotes are removed.
 * - If a field is specified in the markdownFields array, it is converted to HTML using the `marked` library.
 *   - If the field is an array, its elements are joined with newline characters before conversion.
 * 
 * @param {Object} data - The data object containing fields to be processed.
 * @param {Array<string>} markdownFields - An array of field names that should be converted from Markdown to HTML.
 * @returns {Object} The processed data object.
 */
export function processFieldsForDocx(data, markdownFields) {
	// data['html'] = {};
	Object.entries(data).forEach(([key, record]) => {
		// if is an array, join the elements
		// if (Array.isArray(record) && record.length == 1) {
		// 	data[key] = record.join(', ');
		// }
		// if is a string and delimited by '"' then remove the quotes
		if (typeof data[key] === 'string' && data[key].startsWith('"') && data[key].endsWith('"')) {
			data[key] = data[key].slice(1, -1);
		}
	});
	markdownFields.forEach(field => {
		if (data[field]) {
			if (Array.isArray(data[field])) {
				data[field] = marked(data[field].join('\n'));
			} else if(typeof data[field] === 'string') {
				data[field] = marked(data[field]);
			}
		}
	});
	// console.log(data);
	return data;
}

/**
 * Converts a date to a string in the format YYMMDD.
 *
 * @param {Date} date - The date to be formatted.
 * @returns {string} The formatted date string in YYMMDD format, or an empty string if the date is invalid.
 */
export function ymd(date) {
	// can wa add timezone europe/paris
	if(moment(date, "DD/MM/YYYY", true).isValid()) {
		const year = String(date.getFullYear()).slice(-2);  // Get last 2 digits of the year
		const month = String(date.getMonth() + 1).padStart(2, '0');  // Add 1 to month (0-based) and pad to 2 digits
		const day = String(date.getDate()).padStart(2, '0');  // Pad day to 2 digits
		const formattedDate = `${year}${month}${day}`;
		
		return formattedDate;
	} else {
		// throw new Error(`${date} is not a valid date`)
		return "";
	}
}

export function getFrenchFormattedDate(withTime = true) {
	const options = {
		year: '2-digit',
		month: '2-digit',
		day: '2-digit',
		hour: '2-digit',
		minute: '2-digit',
		timeZone: 'Europe/Paris',
		hour12: false,
	};
	
	const formatter = new Intl.DateTimeFormat('fr-FR', options);
	const parts = formatter.formatToParts(new Date());
	
	const dateParts = {};
	parts.forEach(({ type, value }) => {
		dateParts[type] = value;
	});
	if(withTime) {
		return `${dateParts.year}${dateParts.month}${dateParts.day}-${dateParts.hour}${dateParts.minute}`;
		// return `${dateParts.day}/${dateParts.month}/${dateParts.year} ${dateParts.hour}${dateParts.minute}`;
	} else {
		return `${dateParts.year}${dateParts.month}${dateParts.day}`;
		// return `${dateParts.day}/${dateParts.month}/${dateParts.year}`;
	}
	// const now = new Date();
	// return now.toLocaleDateString('fr-FR').replace(/\//g, '-'); // Format as DD-MM-YYYY
}

export async function fetchTemplate(url) {
	console.log(`Fetching template from URL: ${url}`);
	const response = await fetch(url);
	const arrayBuffer = await response.arrayBuffer();
	const buffer = Buffer.from(arrayBuffer);
	console.log(`Fetched template of size: ${buffer.length} bytes`);
	return buffer;
}

export async function generateReport(template, data) {
	return await createReport({
		output: 'buffer',
		template,
		// cmdDelimiter: ['{{', '}}'],
		data,
		// failFast: false,
	});
}

export function ensureDirectoryExists(dir) {
	if (!fs.existsSync(dir)) {
		fs.mkdirSync(dir);
	}
}

const generateAndAttachReportFile = async (req, res, document) => {
	console.log('Generating report file for attachment...', document.name);
	try {
	  let data;
	  const { recordId } = req.query;
  
	  // Fetch the records from Airtable
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
  
	  // Data preprocessing if needed
	  if (data) {
		console.log('Data successfully retrieved:', `${document.multipleRecords ? data.records.length : data.length } records`);
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
  
	  // Generate the report buffer using template
	  console.log(`Generating report using template: ${document.template}`);
	  const reportBuffer = await generateAndDownloadReport(
		`${GITHUBTEMPLATES}${document.template}`,
		data,
		document.titleForming(data)
	  );
  
	  // Set response headers for file download
	  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
	  res.setHeader('Content-Disposition', `attachment; filename=${document.titleForming(data)}.docx`);
	  
	  // Send the generated buffer as the attachment
	  res.send(reportBuffer); // Send the buffer as response
  
	} catch (error) {
	  console.error('Error generating report file:', error);
	  res.status(500).json({ success: false, error: error.message });
	}
  };
  

// Function to create a new record in "Recueil des besoins" and link it to an Inscription record by ID
export const createRecueil = async (inscriptionId) => {
    // First, check if the Inscription record exists
    // const record = await getAirtableRecord('Inscriptions', inscriptionId);
    // if (!record) {
    //     console.log(`Inscription with id ${inscriptionId} does not exist.`);
    //     return;
    // }

    // Define the URL for creating a record in "Recueil des besoins"
    const url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/Recueil%20des%20besoins`;
    const data = {
        fields: {
            Inscrits: [inscriptionId] // Linking to the Inscription ID
        }
    };
    console.log(`Creating record in "Recueil des besoins" at URL: ${url}`);
    console.log(`Data being sent: ${JSON.stringify(data)}`);

    try {
        const response = await axios.post(url, data, {
            headers: {
                ...AUTH_HEADERS,
                'Content-Type': 'application/json'
            }
        });

        console.log(`Response from Airtable: ${JSON.stringify(response.data)}`);
        return response.data;
    } catch (error) {
        console.error(`Failed to create record in "Recueil des besoins":`, error);
        throw new Error('Error creating record in Recueil des besoins');
    }
};

export const serveConvocationPage = async (res, 
    prenom, 
    nom, 
    email, 
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
) => {
    // Prepare halfdaysHtml like in sendConvocation
	var halfdaysHtml = null;
	try {
		const data = await getAirtableRecords("Demi-journées", null, `sessId='${sessId}'`);
		if(data) {
			// console.log('Half days:', data);
			const halfdays = data.records
			.sort((a, b) => new Date(a.debut) - new Date(b.debut))
			.map(record => {
				const params = {date: new Intl.DateTimeFormat('fr-FR', { dateStyle: 'full' }).format(new Date(record.debut)),
				horaires: `${new Intl.DateTimeFormat('fr-FR', { timeStyle: 'short' }).format(new Date(record.debut))} - ${new Intl.DateTimeFormat('fr-FR', { timeStyle: 'short' }).format(new Date(record.fin))}`,
				lieu: record.adresse}
				return `<tr style="border-bottom: 0.4px solid lightgray;">
						<td style="padding: 4px 16px;font-family:'Open Sans';font-size:11px">${params.date}</td>
						<td style="padding: 4px 16px;font-family:'Open Sans';font-size:11px">${params.horaires}</td>
						<td style="padding: 4px 16px;font-family:'Open Sans';font-size:11px">${params.lieu}</td>
					</tr>`;
			});
			halfdaysHtml = `<table>
    <thead>
        <tr style="border-bottom: 0.4px solid lightgray;">
            <th style="padding: 4px 16px;font-size:11px">Date</th>
            <th style="padding: 4px 16px;font-size:11px">Horaires</th>
            <th style="padding: 4px 16px;font-size:11px">Lieu</th>
        </tr>
    </thead>
    <tbody>`+halfdays.join('')+`    </tbody>
</table>`.replace(/font-family:'Open Sans';/g, '');
// halfdaysMjml = '<mj-table font-family="Open Sans" >' + halfdaysHtml + "</mj-table>";

	// console.log(halfdaysMjml);
		}
	} catch(error) {
		console.error('Failed to get half days', error);
		// res.send({ success: false, error: 'Failed to get half days' });
	}


    // let halfdaysHtml = '<table>...</table>'; // This should be generated based on actual records as in the sendConvocation function
    try {
    // MJML template for email body
    const mjmlContent = `
        <!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Confirmation d'inscription à la formation et recueil des besoins</title>
  <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Open Sans', 'Segoe UI', sans-serif; color: #000; }
    .orange { color: #f5a157; }
    .button {
      font-family: 'Open Sans';
      background-color: #f4913a;
      color: white;
      font-size: 16px;
      border-radius: 14px;
      text-decoration: none;
      font-weight: bold;
      padding: 10px 25px;
      display: inline-block;
	  margin-left: 60px;
    }
	.flex-container {
      display: flex;
      justify-content: space-between;
      align-items: center;
      background-color: #f0f0f0; /* Light gray background */
      padding: 10px;
    }
  </style>
</head>
<body>

  <div class="flex-container">
    <p>Envoyer à : <a href="mailto:${email}">${email}</a></p>
    <a href="/updateDateConvoc?inscriptionId=${inscriptionId}">Mettre à jour la date de convocation</a>
  </div>
  
  <p>Bonjour ${prenom} ${nom},</p>
  <p>Nous avons le plaisir de confirmer votre inscription à la formation <em>${titre_fromprog}</em> qui se déroulera ${dates}, ${str_lieu}.</p>

  ${halfdaysHtml}

  <p>Je vous prie de bien vouloir <strong>compléter la fiche de recueil des besoins avant le ${completionDateString}</strong> en cliquant sur le bouton ci-après :</p>

  <a href="${fillout_recueil}" class="button">Complétez votre fiche maintenant</a>

  <p>Vous trouverez sous ce courriel le programme de la formation.</p>
  <p>Pour information, vous trouverez ici le <a style="color: rgb(0, 113, 187)" href="https://www.sante-habitat.org/images/formations/FSH-Livret-accueil-stagiaire.pdf">Livret d’accueil du stagiaire</a></p>
  <p>En attendant de recevoir votre fiche complétée, je me tiens à votre disposition pour toute précision.</p>

  <p style="margin: 0; padding: 0;">Isadora Vuong Van</p>
  <p style="margin: 0; padding: 0;">Pôle Formations</p>
  <p style="margin: 0; padding: 0; font-weight: bold;">Fédération Santé Habitat</p>
  <p style="margin: 0; padding: 0;">6 rue du Chemin Vert - Paris 11ème</p>
  <p style="margin: 0; padding: 0;">Tél. 01 48 05 55 54 / 06 33 82 17 52</p>
  <p style="margin: 0; padding: 0;"><a href="http://www.sante-habitat.org" style="color: #000;">www.sante-habitat.org</a></p>

  <hr style="border: 0.4px solid lightgrey; margin: 20px 0;">

  <h1 class="orange">${titre_fromprog}</h1>

  ${Formateurice ? `<p><strong>Formateur</strong> : ${Formateurice}</p>` : ''}
  ${prerequis_fromprog ? `<p><strong>Prérequis</strong> : ${prerequis_fromprog}</p>` : ''}
  ${public_fromprog ? `<p><strong>Public</strong> : ${public_fromprog}</p>` : ''}
  ${introcontexte_fromprog ? `${introcontexte_fromprog}` : ''}
  ${objectifs_fromprog ? `<h2 style="background-color: #a0d3f6; color: #005f9e; text-align: center;">Objectifs</h2>${objectifs_fromprog}` : ''}
  ${contenu_fromprog ? `<h2 style="background-color: #a0d3f6; color: #005f9e; text-align: center;">Contenu</h2>${contenu_fromprog}` : ''}
  ${modaliteseval_fromprog ? `<h2 style="background-color: #a0d3f6; color: #005f9e; text-align: center;">Modalités d'évaluation</h2>${modaliteseval_fromprog}` : ''}
  ${methodespedago_fromprog ? `<h2 style="background-color: #a0d3f6; color: #005f9e; text-align: center;">Méthodes pédagogiques</h2>${methodespedago_fromprog}` : ''}

</body>
</html>

    `;

    // Convert MJML to HTML
    // const { html } = mjml2html(mjmlContent);

    // Serve HTML
    res.send(mjmlContent);
    // res.send(html);

	} catch(error) {
		res.send({ success: false, error: 'Failed to get half days' });
	}
};
// Helper function to send the HTML email
export const sendConvocation = async (
	inscriptionData,
	str_lieu, 
	fillout_recueil, 
	completionDateString, 
) => {
	const {
		prenom, 
		nom, 
		mail:email, 
		prerequis_fromprog, public_fromprog, titre_fromprog, introcontexte_fromprog, contenu_fromprog, 
		objectifs_fromprog, methodespedago_fromprog, modaliteseval_fromprog,
		Formateurice,
		dates, 
		sessId,
		inscriptionId
	} = inscriptionData;
	const transporter = nodemailer.createTransport({
		host: 'smtp-declic-php5.alwaysdata.net',
		port: 465,
		secure: true,
		auth: {
			user: 'formation@sante-habitat.org',
			pass: process.env.FSH_PASSWORD
		}
	});

	var halfdaysMjml = null;
	var halfdaysHtml = null;

	try {
		const data = await getAirtableRecords("Demi-journées", null, `sessId='${sessId}'`);
		if(data) {
			// console.log('Half days:', data);
			const halfdays = data.records
			.sort((a, b) => new Date(a.debut) - new Date(b.debut))
			.map(record => {
				const params = {date: new Intl.DateTimeFormat('fr-FR', { dateStyle: 'full' }).format(new Date(record.debut)),
				horaires: `${new Intl.DateTimeFormat('fr-FR', { timeStyle: 'short' }).format(new Date(record.debut))} - ${new Intl.DateTimeFormat('fr-FR', { timeStyle: 'short' }).format(new Date(record.fin))}`,
				lieu: record.adresse}
				return `<tr>
						<td style="padding: 4px 16px;font-family:'Open Sans';font-size:11px">${params.date}</td>
						<td style="padding: 4px 16px;font-family:'Open Sans';font-size:11px">${params.horaires}</td>
						<td style="padding: 4px 16px;font-family:'Open Sans';font-size:11px">${params.lieu}</td>
					</tr>`;
			});
			halfdaysHtml = `<table>
    <thead>
        <tr>
            <th style="padding: 4px 16px;font-size:11px">Date</th>
            <th style="padding: 4px 16px;font-size:11px">Horaires</th>
            <th style="padding: 4px 16px;font-size:11px">Lieu</th>
        </tr>
    </thead>
    <tbody>`+halfdays.join('')+`    </tbody>
</table>`.replace(/font-family:'Open Sans';/g, '');
halfdaysMjml = '<mj-table font-family="Open Sans" >' + halfdaysHtml + "</mj-table>";

	// console.log(halfdaysMjml);
		}
	} catch(error) {
		console.error('Failed to get half days', error);
	}

	const enIntra = inscriptionData.lieux.includes('intra');

	// attach a file
	const docDefinition = documents.find(doc => doc.name === 'facture');
	let attachmentBuffer = null;
	const sessionId = Array.isArray(sessId) ? sessId[0] : sessId;
	let filename;
	let attachments = [];
	if(!enIntra) {
		try {
			// const factureData = await getAirtableRecord(docDefinition.table, inscriptionId);
				// const template = await fetchTemplate(`${GITHUBTEMPLATES}${docDefinition.template}`);
				filename = docDefinition.titleForming(inscriptionData)+ ".pdf";
				// attachmentBuffer = await generateReport(template, data);
				attachmentBuffer = await createFacturePdf(inscriptionData);
		} catch(error) {
			console.error('Failed to generate report for attachment', error);
		}
		if(attachmentBuffer) {
			attachments = [
				{
					filename: filename || 'facture.pdf',
					content: attachmentBuffer,
					contentType: 'application/pdf'
				}
			];
		}
	}
	

	// MJML template for the email body
	const mjmlContent = `
	<mjml>
		<mj-head>
			<mj-attributes>
				<mj-class name="orange" color="#f5a157" />
				<mj-all font-family="Open Sans, sans-serif" color="#000" />
			</mj-attributes>
			<mj-preview>Vous êtes inscrit(e) ! Veuillez compléter votre fiche de recueil des besoins.</mj-preview>
			<mj-title>Confirmation d'inscription à la formation et recueil des besoins</mj-title>
		</mj-head>
		<mj-body>
			<mj-text>
				<p>Bonjour ${prenom} ${nom},</p>
				<p>Nous avons le plaisir de confirmer votre inscription à la formation <em>${titre_fromprog}</em> qui se déroulera ${dates}, ${str_lieu}.</p>
			</mj-text>
			${halfdaysMjml ? `${halfdaysMjml}` : ''}
			<mj-text>
				<p>Je vous prie de bien vouloir <strong>compléter la fiche de recueil des besoins avant le ${completionDateString}</strong> en cliquant sur le bouton ci-après :</p>
			</mj-text>

			<!-- Centered Button in Section and Column -->
			<mj-section padding="0">
				<mj-column width="100%">
					<mj-button font-family="Open Sans" background-color="#f4913a" color="white" font-size="16px" border-radius="14px" href="${fillout_recueil}">
						<a href="${fillout_recueil}" style="text-decoration:none;color:white;font-weight:bold;">Complétez votre fiche maintenant</a>
					</mj-button>
				</mj-column>
			</mj-section>

			<mj-text>
				<p>Vous trouverez sous ce courriel le programme de la formation${!enIntra && attachments.length > 0 ? `, ainsi que votre facture en pièce jointe` : ''}.</p>
				<p>Pour information, vous trouverez ici le <a style="color: rgb(0, 113, 187)" href="https://www.sante-habitat.org/images/formations/FSH-Livret-accueil-stagiaire.pdf">Livret d’accueil du stagiaire</a></p>
				<p>En attendant de recevoir votre fiche complétée, je me tiens à votre disposition pour toute précision.</p>
			</mj-text>

			<!-- Signature -->
			<mj-text font-family="Open Sans, sans-serif">
				<p style="margin: 0; padding: 0;">Isadora Vuong Van</p>
				<p style="margin: 0; padding: 0;">Pôle Formations</p>
				<p style="margin: 0; padding: 0; font-weight:bold;">Fédération Santé Habitat</p>
				<p style="margin: 0; padding: 0;">6 rue du Chemin Vert - Paris 11ème</p>
				<p style="margin: 0; padding: 0;">Tél. 01 48 05 55 54 / 06 33 82 17 52</p>
				<p style="margin: 0; padding: 0;"><a href="http://www.sante-habitat.org" style="color: #000;">www.sante-habitat.org</a></p>
			</mj-text>
		
			

			<!-- Image below signature -->
			<!-- <mj-image src="https://www.sante-habitat.org/images/2019/LOL.png" alt="FSH Logo" width="3.02cm" height="1.15cm" /> -->
			
			<mj-divider border-width="1px" border-color="lightgrey" />

			<mj-text mj-class="orange">
				<h1>${titre_fromprog}</h1>
			</mj-text>
			<mj-text>
				${Formateurice ? `<p><strong>Formateur</strong> : ${Formateurice}</p>` : ''}
				${prerequis_fromprog ? `<p><strong>Prérequis</strong> : ${prerequis_fromprog}</p>` : ''}
				${public_fromprog ? `<p><strong>Public</strong> : ${public_fromprog}</p>` : ''}
				${introcontexte_fromprog ? `${introcontexte_fromprog}` : ''}
				${objectifs_fromprog ? `<h2 style="background-color:#a0d3f6; color: #005f9e;text-align:center">Objectifs</h2>${objectifs_fromprog}` : ''}
				${contenu_fromprog ? `<h2 style="background-color:#a0d3f6; color: #005f9e;text-align:center">Contenu</h2>${contenu_fromprog}` : ''}
				${modaliteseval_fromprog ? `<h2 style="background-color:#a0d3f6; color: #005f9e;text-align:center">Modalités d'évaluation</h2>${modaliteseval_fromprog}` : ''}
				${methodespedago_fromprog ? `<h2 style="background-color:#a0d3f6; color: #005f9e;text-align:center">Méthodes pédagogiques</h2>${methodespedago_fromprog}` : ''}
			</mj-text>

		</mj-body>
	</mjml>

    `;

	// Convert MJML to HTML
	const { html } = mjml(mjmlContent);

	// Define the email options
	const mailOptions = {
		from: '"Formations FSH" <formation@sante-habitat.org>',
		// to: email,
		bcc: 'formation@sante-habitat.org',
		replyTo: 'formation@sante-habitat.org',
		subject: `Convocation à la formation ${titre_fromprog}`,
		html: html,  // MJML rendered to HTML
		attachments,
		alternatives: [
			{
			  contentType: 'text/plain',
			  content: `
		Bonjour ${prenom} ${nom},
		
		Nous avons le plaisir de confirmer votre inscription à la formation ${titre_fromprog} qui se déroulera le ${dates}, ${str_lieu}.
		
		Je vous prie de bien vouloir compléter la fiche de recueil des besoins avant le ${completionDateString} en cliquant sur le lien suivant :
		${fillout_recueil}
		
		Vous trouverez ci-dessous le programme de la formation.
		
		Pour information, vous trouverez ici le Livret d’accueil du stagiaire : https://www.sante-habitat.org/images/formations/FSH-Livret-accueil-stagiaire.pdf
		
		Dans cette attente, et restant à votre disposition pour tout renseignement complémentaire.
		
		Bien cordialement,
		Isadora Vuong Van
		Pôle Formations
		Fédération Santé Habitat
		6 rue du Chemin Vert - Paris 11ème
		Tél. 01 48 05 55 54 / 06 33 82 17 52
		www.sante-habitat.org

		--------------------------------

		${titre_fromprog}

		${Formateurice ? `Formateur : ${Formateurice}\n` : ''}
		${prerequis_fromprog ? `Prérequis : ${prerequis_fromprog}\n` : ''}
		${public_fromprog ? `Public : ${public_fromprog}\n` : ''}
		${introcontexte_fromprog ? `${introcontexte_fromprog}\n` : ''}
		${objectifs_fromprog ? `\nObjectifs\n${objectifs_fromprog}\n` : ''}
		${contenu_fromprog ? `\nContenu\n${contenu_fromprog}\n` : ''}
		${modaliteseval_fromprog ? `\nModalités d'évaluation\n${modaliteseval_fromprog}\n` : ''}
		${methodespedago_fromprog ? `\nMéthodes pédagogiques\n${methodespedago_fromprog}\n` : ''}
			  `
			},
			{
				contentType: 'text/html',
				content: `
				  <html>
					<body>
					  <p>Bonjour ${prenom} ${nom},</p>
					  <p>Nous avons le plaisir de confirmer votre inscription à la formation <strong>${titre_fromprog}</strong> qui se déroulera le <strong>${dates}, ${str_lieu}</strong>.</p>
					  ${halfdaysHtml ? halfdaysHtml : ''}
					  <p><strong>Merci de compléter la fiche de recueil des besoins avant le <u>${completionDateString}</u> en cliquant sur ce bouton :</strong></p>
						<div style="text-align: center;">
						<a href="${fillout_recueil}" style="color: white; background-color: #f4913a; font-weight: bold; padding: 10px 20px; border-radius: 5px; text-decoration: none;">Remplir la fiche des besoins</a>
						</div>
						<p>Vous trouverez ci-dessous le programme de la formation.</p>
					  <p>Pour plus d'informations, vous pouvez consulter le <a href="https://www.sante-habitat.org/images/formations/FSH-Livret-accueil-stagiaire.pdf">Livret d’accueil du stagiaire</a>.</p>
					  <p>Bien cordialement,</p>
					  <p>Isadora Vuong Van<br/>
						 Pôle Formations<br/>
						 <strong>Fédération Santé Habitat</strong><br/>
						 6 rue du Chemin Vert - Paris 11ème<br/>
						 Tél. 01 48 05 55 54 / 06 33 82 17 52<br/>
						 <a href="http://www.sante-habitat.org">www.sante-habitat.org</a></p>

				<h1>${titre_fromprog}</h1>
				${Formateurice ? `<p><strong>Formateur</strong> : ${Formateurice}</p>` : ''}
				${prerequis_fromprog ? `<p><strong>Prérequis</strong> : ${prerequis_fromprog}</p>` : ''}
				${public_fromprog ? `<p><strong>Public</strong> : ${public_fromprog}</p>` : ''}
				${introcontexte_fromprog ? `${introcontexte_fromprog}` : ''}
				${objectifs_fromprog ? `<h2>Objectifs</h2>${objectifs_fromprog}` : ''}
				${contenu_fromprog ? `<h2>Contenu</h2>${contenu_fromprog}` : ''}
				${modaliteseval_fromprog ? `<h2>Modalités d'évaluation</h2>${modaliteseval_fromprog}` : ''}
				${methodespedago_fromprog ? `<h2>Méthodes pédagogiques</h2>${methodespedago_fromprog}` : ''}

					</body>
				  </html>
				`
				// .replace(/é/g, '&eacute;')
				// .replace(/è/g, '&egrave;')
				// .replace(/ê/g, '&ecirc;')
				// .replace(/ë/g, '&euml;')
				// .replace(/à/g, '&agrave;')
				// .replace(/â/g, '&acirc;')
				// .replace(/ä/g, '&auml;')
				// .replace(/ç/g, '&ccedil;')
				// .replace(/î/g, '&icirc;')
				// .replace(/ï/g, '&iuml;')
				// .replace(/ô/g, '&ocirc;')
				// .replace(/ö/g, '&ouml;')
				// .replace(/ù/g, '&ugrave;')
				// .replace(/û/g, '&ucirc;')
				// .replace(/ü/g, '&uuml;')
				// .replace(/œ/g, '&oelig;')
				// .replace(/’/g, '&rsquo;')
				// .replace(/'/g, '&apos;')
				// .replace(/"/g, '&quot;')
				// .replace(/&/g, '&amp;')
				// .replace(/€/g, '&euro;')
				// .replace(/«/g, '&laquo;')
				// .replace(/»/g, '&raquo;')
				// .replace(/–/g, '&ndash;')
				// .replace(/—/g, '&mdash;')
				// .replace(/…/g, '&hellip;')
			  },
			// {
			//   contentType: 'application/json',
			//   content: JSON.stringify({
			// 	subject: `Convocation à la formation ${titre_fromprog}`,
			// 	body: {
			// 	  text: `Bonjour ${prenom} ${nom},\n\nVotre inscription à la formation ${titre_fromprog} est confirmée pour le ${dates}, à ${str_lieu}.`,
			// 	  links: [
			// 		{
			// 		  label: 'Complétez votre fiche maintenant',
			// 		  url: fillout_recueil
			// 		},
			// 		{
			// 		  label: 'Livret d’accueil du stagiaire',
			// 		  url: 'https://www.sante-habitat.org/images/formations/FSH-Livret-accueil-stagiaire.pdf'
			// 		}
			// 	  ]
			// 	}
			//   })
			// },
			{
			  contentType: 'application/xml',
			  content: `
				<email>
				  <subject>Convocation à la formation ${titre_fromprog}</subject>
				  <body>
					<text>Bonjour ${prenom} ${nom},</text>
					<text>Nous avons le plaisir de confirmer votre inscription à la formation ${titre_fromprog} qui se déroulera le ${dates}, à ${str_lieu}.</text>
					<text>Veuillez remplir la fiche de recueil des besoins avant le ${completionDateString} en cliquant sur ce <url>${fillout_recueil}</url>.</text>
					<text>Consultez le Livret d’accueil du stagiaire à cette adresse : <url>https://www.sante-habitat.org/images/formations/FSH-Livret-accueil-stagiaire.pdf</url></text>
				  </body>
				</email>
			  `
			}
		  ]	
		};

	// Send the email
	try {
		await transporter.sendMail(mailOptions);
		console.log(`Email sent to ${email}`);
	  } catch (error) {
		console.error('Error sending email:', error);
	  }
};

// Function to send email to all eligible records for a specific session
export const sendConfirmationToAllSession = async (sessId) => {
	// Fetch records from Airtable with conditions (only records without envoi_convocation)
	const { records } = await getAirtableRecords('Inscriptions', null, `{sessId} = '${sessId}' AND {Statut} = 'Enregistrée' AND {envoi_convocation} = ""`);
	if (records.length === 0) {
		console.log('No records found');
		return;
	}

	const rec1 = records[0];
	let str_lieu = '';
	if (rec1.nb_adresses === 1) {
		if (rec1.lieux === "Visioconférence") {
			str_lieu = "en visioconférence (le lien de connexion vous sera envoyé prochainement)";
		} else if (rec1.lieux.includes("Siège")) {
			str_lieu = "au siège de la FSH, 6 rue du Chemin vert, 75011 Paris";
		} else if (rec1.lieux.includes("intra")) {
			str_lieu = `à l'adresse : ${rec1.adresses_intra}`;
		}
	} else if (rec1.nb_adresses > 1) {
		console.log(`Record ${rec1.inscriptionId} has multiple addresses and cannot determine location`);
		return;
	}

	// Calculate the completion date string
	const completionDate = new Date(rec1.du);
	completionDate.setDate(completionDate.getDate() - 15);
	
	const today = new Date();
	while (completionDate < today) {
		completionDate.setDate(completionDate.getDate() + 1);
	}
	
	const completionDateString = completionDate.toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' });


	for (const record of records) {
		// const { inscriptionId } = record;
		// await sendConfirmation(inscriptionId);
		const { envoi_convocation, 
			prenom, nom, mail:email, 
			titre_fromprog, 
			fillout_recueil, du,
			dates, 
			sessId
		
		} = record;
	
		if (envoi_convocation) {
			console.log(`Record ${prenom} ${nom} already has envoi_convocation`);
			return;
		}
		if (!(prenom && nom && mail && titre_fromprog && dates && lieux && fillout_recueil && du)) {
			console.log(`Record ${inscriptionId} is missing necessary fields`);
			return;
		}
	
		let recueilLink = fillout_recueil;
		if (!recueilLink) {
			try {
				const recueil = await createRecueil(inscriptionId);
				if (recueil.fields && recueil.fields.fillout) {
					recueilLink = recueil.fields.fillout; // Use the 'fillout' field from the created record
					console.log(`Created recueil with fillout link: ${recueilLink}`);
				} else {
					recueilLink = "https://forms.fillout.com/t/1wNMoFGDTYus?id=" + recueil.id; // Use the recueilId for the link
					console.log(`Created recueil, but 'fillout' field is missing. Using default link: ${recueilLink}`);
				}
			} catch (error) {
				console.error(`Failed to create recueil for inscriptionId ${inscriptionId}:`, error);
				return;
			}
		}

		try {
			// Send the convocation email
			console.log(`Will send email with this data:`, {
				record,
				str_lieu, 
				fillout_recueil: fillout_recueil || recueilLink,
				completionDateString, 
				sessId
			});
			
			await sendConvocation(
				record,
				str_lieu,
				recueilLink,
				completionDateString,
			)
				
	
			// Update the record with the current datetime
			const currentDateTime = new Date().toISOString().split('T')[0]; // Extract the date part in YYYY-MM-DD format
			await updateAirtableRecord('Inscriptions', inscriptionId, { envoi_convocation: currentDateTime });
			console.log(`Updated record ${inscriptionId} with envoi_convocation: ${currentDateTime}`);
	
		} catch (error) {
			console.error(`Failed to send email to ${mail} or update Airtable record:`, error);
		}
	
		
	}
};

// Function to send email for a single Inscription record
export const sendConfirmation = async (inscriptionId) => {
	// Retrieve a single record
	const record = await getAirtableRecord('Inscriptions', inscriptionId);
	if (!record) {
		console.log(`Record with id ${inscriptionId} not found`);
		return;
	}

	// Check if envoi_convocation is empty and necessary fields are available
	const { envoi_convocation, 
		prenom, nom, mail:email, 
		titre_fromprog, 
		adresses_intra, nb_adresses, 
		lieux, 
		fillout_recueil, du,
		dates, 
	} = record;
	if (envoi_convocation) {
		console.log(`Record ${prenom} ${nom} already has envoi_convocation`);
		return;
	}
	if (!(prenom && nom && mail && titre_fromprog && dates && lieux && fillout_recueil && du)) {
		console.log(`Record ${inscriptionId} is missing necessary fields`);
		return;
	}

	let recueilLink = fillout_recueil;
	if (!recueilLink) {
		try {
			const recueil = await createRecueil(inscriptionId);
			if (recueil.fields && recueil.fields.fillout) {
				recueilLink = recueil.fields.fillout; // Use the 'fillout' field from the created record
				console.log(`Created recueil with fillout link: ${recueilLink}`);
			} else {
				recueilLink = "https://forms.fillout.com/t/1wNMoFGDTYus?id=" + recueil.id; // Use the recueilId for the link
				console.log(`Created recueil, but 'fillout' field is missing. Using default link: ${recueilLink}`);
			}
		} catch (error) {
			console.error(`Failed to create recueil for inscriptionId ${inscriptionId}:`, error);
			return;
		}
	}
	
	// Determine the location description based on the available data
	let str_lieu = '';
	if (nb_adresses === 1) {
		if (lieux === "Visioconférence") {
			str_lieu = "en visioconférence (le lien de connexion vous sera envoyé prochainement)";
		} else if (lieux.includes("Siège")) {
			str_lieu = "au siège de la FSH, 6 rue du Chemin vert, 75011 Paris";
		} else if (lieux.includes("intra")) {
			str_lieu = `à l'adresse : ${adresses_intra}`;
		}
	} else if (nb_adresses > 1) {
		console.log(`Record ${inscriptionId} has multiple addresses and cannot determine location`);
		return;
	}

	// Calculate the completion date string
	const completionDate = new Date(du);
	completionDate.setDate(completionDate.getDate() - 15);

	const today = new Date();
	while (completionDate < today) {
		completionDate.setDate(completionDate.getDate() + 1);
	}

	const completionDateString = completionDate.toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' });
	
	try {
		// Send the convocation email
		// console.log(`Will send email with this data:`, {
		// 	prenom, 
		// 	nom, 
		// 	email, 
		// 	Formateurice,
		// 	dates, 
		// 	str_lieu, 
		// 	fillout_recueil: fillout_recueil || recueilLink,
		// 	completionDateString, 
		// 	sessId
		// });

		console.log(`Will send email to:`, record.nom)
		
		await sendConvocation(
			record,
			str_lieu,
			recueilLink,
			completionDateString,
		)

		// Update the record with the current datetime
		const currentDateTime = new Date().toISOString().split('T')[0]; // Extract the date part in YYYY-MM-DD format
		await updateAirtableRecord('Inscriptions', inscriptionId, { envoi_convocation: currentDateTime });
		console.log(`Updated record ${inscriptionId} with envoi_convocation: ${currentDateTime}`);

	} catch (error) {
		console.error(`Failed to send email to ${mail} or update Airtable record:`, error);
	}
};

export async function createFacturePdf(factureData) {
	var data = factureData;
	if(data["date_facture"]) {
		data["today"] = new Date(data["date_facture"]).toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'});
	}
	data['acquit'] = data["paye"].includes("Payé")
	? `Acquittée par ${data.moyen_paiement.toLowerCase()} le ${(new Date(data.date_paiement)).toLocaleDateString('fr-FR')}`
	: "";
	const possibleArrays = ["entite", "nom", "prenom", "cp", "ville", "rue", "etab_payeur", "ville_payeur", "cp_payeur", "rue_payeur", "entite_payeur"];
	possibleArrays.forEach(field => {
		if (Array.isArray(data[field])) {
			data[field] = data[field][0]; // Extract the first element if it's an array
		}
	});

	data.entite = data.entite_payeur || data.entite || "";
	data.cp = data.cp_payeur || data.cp || "";
	data.rue = data.rue_payeur || data.rue || "";
	data.ville = data.ville_payeur || data.ville || "";
	console.log("data", data)
	data.factId = `${ymd(new Date(data["du"]))}${data.sessCode ? data.sessCode.replace("[","").replace("]","") : ''}${data.entite ? data.entite.substring(0,2) : ''}${data["nom"] ? data["nom"].substring(0,2).toUpperCase() : ''}${data["prenom"] ? data["prenom"].substring(0,1) : ''}`;

	data["Montant"] = calculTotalPrixInscription(data)
	// logIfNotVercel("Montant calc", data["Montant"])
	data['montant'] = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(
		parseFloat(data["Montant"]),
	);  

	const template = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Facture</title>
    <link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;700&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: 'Open Sans', 'Segoe UI', sans-serif;
        }
        p {
            display: block;
            margin-block-start: 0;
            margin-block-end: 0;
            margin-inline-start: 0px;
            margin-inline-end: 0px;
            unicode-bidi: isolate;
        }
        .header, .footer {
            /* text-align: center; */
        }
        .bluetext {
            color: #063c64;
        }
        .redtext {
            color: #d51038;
        }
        .smtext {
            font-size: small;
        }
        .textxxs {
            font-size: xx-small;
        }
        .bold {
            font-weight: bold;
        }
        .details, .footer {
            margin-top: 20px;
            font-size: 12px;
        }
        .section-title {
            font-weight: bold;
            margin-top: 20px;
        }
        .table {
            width: 100%;
            border-collapse: collapse;
        }
        /* .table, .table th, .table td { */
            /* border: 0.4px solid #333; */
        /* } */
        .table th, .table td {
            padding: 8px;
            text-align: left;
        }
        .totals {
            margin-top: 10px;
        }
        .totals .right {
            text-align: right;
        }

        .smallcaps {
            font-variant: small-caps;
            font-size: larger;
            font-weight: 600;
        }

        .tablenoborder {
            border: none !important;
        }
        .text-left {
            text-align: left;
        }
        .text-center {
            text-align: center;
        }
        .text-right {
            text-align: right !important;
        }
        .table-paiements {
            width: 100%;
            border-collapse: collapse;
            background-color: lightgray;
            border-bottom: 0.4px solid black; 
            border-top: 0.4px solid #333;
        }
        .altop {
            vertical-align: top;
        }

        .totaldiv div {
            float: right;
            width: 158px;
        }

        .blnone {
            border: 0.4px solid #063c64;
            border-left: none;
        }

    </style>
</head>
<body>
<div style="width: 100%; text-align: center; font-size: 12px; padding-top: 10px;">
            <table style="width: 100%; border: 0;">
              <tr>
                <td style="width: 161px; text-align: left;">
                    <img src="https://github.com/IVV-FSH/templater-fsh/blob/main/assets/Logo%20FSH%20-%20transparent.png?raw=true" style="width: 160px; height: auto;">
                </td>
                <td style="width: auto; text-align: left; font-size: large;"><p class="bold">FORMATIONS</p><p>Facture n°${data.factId}</p></td>
                <td style="width: 56px; text-align: right;">Page 1/1</td>
              </tr>
            </table>
          </div>
    <div style="display: grid; grid-template-columns: 1fr 1fr; width: 100%;">
        <div>
            FÉDÉRATION SANTÉ HABITAT</br>
            6 rue du Chemin Vert</br>
            75011 PARIS 11</br>
            Tel : +33 1 48 05 55 54</br>
            Port. : +33 6 33 82 17 52</br>
            formation@sante-habitat.org</br>
            N° SIRET : 43776264400049</br>
            Code NAF : 8790B</br>
        </div>
        <div class="bluetext">
            <div class="smallcaps bluetext">destinataire</div>
            <p><span class="bold">${data.entite}</span><br>
            ${data.rue}<br>
            ${data.cp} ${data.ville}</p>
            <p style="margin-top: 8px;"><span class="smallcaps">stagiaire</span> : <span class="bold">${data.prenom} ${data.nom}</span><br>
            ${data.poste}<br>
            ${data.mail}</p>
        </div>
    </div>
    <table class="tablenoborder text-left table">
        <tr>
            <th>
                <h1 class="bluetext">Facture n°${data.factId}</h1>
                <p class="bluetext">du ${data.today}</p>
            </th>
            <th style="max-width:86px;" class="text-center">
                ${data.paye.includes("Payé") ? `<p class="redtext smtext text-center">Facture acquittée le ${data.date_paiement} par ${data.moyen_paiement}</p>` : ""}
            </th>
        </tr>
    </table>
    

    <table class="table tablenoborder" style="border: 0.4px solid #063c64; border-collapse: collapse;">
        <tr style="background-color: #063c64; color:white;">
            <th>Formation</th>
            <th class="text-right altop" style="border: 0.4px solid #063c64;">Quantité</th>
            <th class="text-right altop" style="width: 108px; border: 0.4px solid #063c64;">Montant TTC</th>
        </tr>
        <tr>
            <td class="altop" style="border: 0.4px solid #063c64;">
                <p class="bold">${data.titre_fromprog}</p>
                <p>Dates : ${data.dates}</p>
                <p>Lieu : ${data.lieux}</p>
            </td>
            <td class="altop text-right" style="width: 80px;border: 0.4px solid #063c64;">1,00</td>
            <td class="altop text-right" style="width: 108px;border: 0.4px solid #063c64;">${data.montant}</td>
        </tr>
        <tr>
            <td></td>
            <td colspan="2" class="textxxs text-center">TVA non applicable (article 239B du CGI)</td>
        </tr>
        <tr class="text-right bold">
            <td></td>
            <td colspan="2" class="text-right bold" style="background-color: #063c64; color:white;">${!data.paye.includes("Payé") ? "<p>Net à payer</p>" : "<p>Réglé</p>"}</td>
        </tr>
        <tr>
            <td></td>
            <td class="text-left bold" style="background-color: #063c64; color:white;">Total TTC</td>
            <td class="text-right bold" style="background-color: #063c64; color:white;">${data.montant}</td>
        </tr>
    </table>


    ${!data.paye.includes("Payé") ? `
		
    <table class="table-paiements" style="margin-bottom: 8px; margin-top:8px;">
        <thead>
          <td style="width: 260px;">
            <p class="bold">Règlement par virement</p>
            <p>Indiquer la référence :</p> 
            <p class="text-center">${data.factId}</p>
          </td>
          <td>
              <p>CCM STRASBOURG KRUTENAU<br>
              IBAN FR76 1027 8010 8800 0277 6084 557</p>
          </td>    
        </thead>  
    </table>

    <table class="table-paiements">
        <thead>
          <td style="width: 260px;">
            <p class="bold">Règlement par chèque</p>
            <p>Indiquer la référence :</p> 
            <p class="text-center">${data.factId}</p>
          </td>
          <td>
            <p>A l’ordre de la FÉDÉRATION SANTÉ HABITAT<br>
                6 rue du Chemin Vert 75011 PARIS</p>
          </td>
        </thead>  
    </table>    
   

    <p class="textxxs">En cas de retard de paiement, des indemnités de retard seront appliquées au taux de 10% du montant total dû par jour de retard, conformément à l'article L. 441-10 du Code de commerce.</p>
    <p class="textxxs">Toute somme non réglée à l'échéance entraînera l'application d'intérêts de retard au taux de 10% par an, calculés à partir de la date d'échéance jusqu'au paiement complet.</p>
    ` : ""}
</body>
</html>
`;

    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    // Set HTML content and wait until it is fully loaded
    await page.setContent(template, { waitUntil: 'networkidle0' });

    // Generate PDF with A4 size
    const pdfBuffer = await page.pdf({
        format: 'A4',
        printBackground: true,
        displayHeaderFooter: true,
        // headerTemplate: `
        // <div style="width: 100%; text-align: center; font-size: 12px; padding-top: 10px;">
        //     <table style="width: 100%; border: 0;">
        //       <tr>
        //         <td style="width: 161px; text-align: left;">
        //             <img src="https://github.com/IVV-FSH/templater-fsh/blob/main/assets/Logo%20FSH%20-%20transparent.png?raw=true" style="width: 160px; height: auto;">
        //         </td>
        //         <td style="width: auto; text-align: left; font-size: large;"><p class="bold">FORMATIONS</p><p>Facture n°${'${factId}'}</p></td>
        //         <td style="width: 56px; text-align: right;">Page ${'${pageNumber}'} / ${'${totalPages}'}</td>
        //       </tr>
        //     </table>
        //   </div>
        // `,
        footerTemplate: `
        <div style="font-family: 'Segoe UI', sans-serif;font-size: 10px; text-align: center; width: 100%; ">
            <p style="padding-bottom:0";margin-block-start: 0;
            margin-block-end: 0;
            margin-inline-start: 0px;
            margin-inline-end: 0px;>FÉDÉRATION SANTÉ HABITAT │ 6 rue du Chemin Vert, 75011 Paris │ 01 48 05 55 54 │ www.sante-habitat.org</p>
            <p style="padding-bottom:0";margin-block-start: 0;
            margin-block-end: 0;
            margin-inline-start: 0px;
            margin-inline-end: 0px;>SIRET 437 762 644 000 49 │ Code APE/NAF 8790B │ Organisme de formation n°11 75 49764 75 - certifié Qualiopi</p>
        </div>
        `,
        margin: {
            top: '0',
            bottom: '120px',
            left: '20px',
            right: '20px'
        }
    });

    await browser.close();
    return pdfBuffer;

}