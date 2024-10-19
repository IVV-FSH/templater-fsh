import fs from 'fs';
import path from 'path';
import { createReport } from 'docx-templates';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import axios from 'axios';
import { marked } from 'marked';
// import { broadcastLog } from './server.js';

dotenv.config();
const AIRTABLE_API_KEY = process.env.AIRTABLE_TOKEN; // Get API key from environment variable
const AIRTABLE_BASE_ID = 'appK5MDuerTOMig1H'; // Replace with your Airtable Base ID
const AUTH_HEADERS = {
	Authorization: `Bearer ${AIRTABLE_API_KEY}`,
};

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
]

export const getAirtableSchema = async (table) => {
	let url = `https://api.airtable.com/v0/meta/bases/${AIRTABLE_BASE_ID}/tables`;
	
	const response = await axios.get(url, {
		headers: AUTH_HEADERS,
	});

	const tables = response.data.tables;

	return tables;

}

// export const getAirtableData = async (table, recordId = null, view = null, formula = null) => {
// 	try {
// 		let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${table}`;
// 		if (recordId && formula) {
// 			throw new Error('Cannot specify both recordId and formula.');
// 		}
// 		const response = await axios.get(url, {
// 			headers: AUTH_HEADERS,
// 		});

// 		let processedData = recordId ? response.data.fields : response.data.records;

// 		if (recordId) {
// 			url += `/${recordId}`;

// 			processedData = processFieldsForDocx(
// 				processedData,
// 				airtableMarkdownFields
// 			);
// 			return processedData;  
// 		} else {
// 			const params = [];
// 			if (formula) {
// 				params.push(`filterByFormula=${formulaFilter(formula)}`);
// 			}
// 			if (view) {
// 				params.push(`view=${formulaFilter(view)}`);
// 			}
// 			if (params.length > 0) {
// 				url += `?${params.join('&')}`;
// 			}
// 			console.log(url);

// 			processedData = response.data.records.map(record => record.fields);

// 			// processedData = processedData.map(record => {
// 			// if (Array.isArray(record)) {
// 			// 	return record.join(', ');
// 			// }
// 			// return record;
// 			// });
			
// 			processedData = processedData.map(record => processFieldsForDocx(
// 				record,
// 				airtableMarkdownFields
// 			));
			
// 			return { records: processedData };
			

// 		}

// 	} catch (error) {
// 		console.error(error);
// 		throw new Error('Error retrieving data from Airtable.');
// 	}
// };

/**
 * Adds missing fields to an object.
 *
 * This function takes an array of field names and an object containing fields with missing values.
 * It returns a new object where each field from the array is present. If a field is missing in the
 * input object, it will be set to `null` in the result.
 *
 * @param {string[]} allFieldsFromTable - An array of field names to be included in the result.
 * @param {Object} fieldsWithMissing - An object containing fields with their values. Missing fields will be set to `null`.
 * @returns {Promise<Object>} A promise that resolves to an object with all fields from the array, with missing fields set to `null`.
 */
const addMissingFields = async (allFieldsFromTable, fieldsWithMissing) => {
    let res = {};
    allFieldsFromTable.forEach((fn) => {
        if (fieldsWithMissing[fn]) {
            res[fn] = fieldsWithMissing[fn];
        } else {
            res[fn] = null;
        }
    });
    return res;
};

export const getAirtableRecord = async (table, recordId) => {
    try {
        let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${table}/${recordId}`;

        const response = await axios.get(url, {
            headers: AUTH_HEADERS,
        });

        const sch = await getAirtableSchema(table);
        const allFields = sch.fieldNames;

        // console.log(response.data.fields);
        const withAllFields = await addMissingFields(allFields, response.data.fields);
        console.log(withAllFields);

        // let processedData = recordId
        //     ? response.data.fields
        //     : response.data.records;

        let processedData = processFieldsForDocx(
            withAllFields,
            airtableMarkdownFields,
        );
        return processedData;
    } catch (error) {
        console.error(error);
        throw new Error("Error retrieving data from Airtable.");
    }
};
export const getAirtableRecords = async (table, view = null, formula = null) => {
	try {
	  let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${table}`;
	  const params = [];
	  if (formula) {
		params.push(`filterByFormula=${formulaFilter(formula)}`);
	  }
	  if (view) {
		params.push(`view=${encodeURIComponent(view)}`);
	  }
	  if (params.length > 0) {
		url += `?${params.join('&')}`;
	  }
	  console.log(`Fetching records from URL: ${url}`);
	//   broadcastLog(`Fetching records from URL: ${url}`); // FIXME:
	
	  const response = await axios.get(url, {
		headers: AUTH_HEADERS,
	  });
  
	  let processedData = response.data.records.map(record => record.fields);
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


// +++IF $prog.lieuxdemij_cumul.includes("Siège")+++
// La salle de formation répond aux exigences de l’accessibilité aux personnes à mobilité réduite.
// +++END-IF+++

// +++IF $prog.lieuxdemij_cumul.includes("intra")+++
// En intra : tarif sur devis
// +++END-IF+++
// +++IF $prog.ouvertepersaccomp_fromprog != null+++
// Personne accompagnée : moitié prix
// +++END-IF+++


export function getFrenchFormattedDate() {
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
	
	return `${dateParts.year}${dateParts.month}${dateParts.day}-${dateParts.hour}${dateParts.minute}`;
}

export async function fetchTemplate(url) {
	const response = await fetch(url);
	const arrayBuffer = await response.arrayBuffer();
	return Buffer.from(arrayBuffer);
}

export async function generateReport(template, data) {
	return await createReport({
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