import fs from 'fs';
import path from 'path';
import { createReport } from 'docx-templates';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import axios from 'axios';
import { marked } from 'marked';

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

export const getAirtableData = async (table, recordId = null, formula = null) => {
	try {
		let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${table}`;
		if (recordId) {
			url += `/${recordId}`;
		} else if (formula) {
			url += `?filterByFormula=${formulaFilter(formula)}`;
		}
		
		const response = await axios.get(url, {
			headers: AUTH_HEADERS,
		});
		
		return recordId ? response.data.fields : response.data.records; // Return fields if specific record, otherwise return records
	} catch (error) {
		console.error(error);
		throw new Error('Error retrieving data from Airtable.');
	}
};

export function processMarkdownFields(data, fieldsToProcess) {
	// data['html'] = {};
	Object.entries(data).forEach(([key, record]) => {
		// if is a string and delimited by '"' then remove the quotes
		if (typeof data[key] === 'string' && data[key].startsWith('"') && data[key].endsWith('"')) {
			data[key] = data[key].slice(1, -1);
		}
	});
	fieldsToProcess.forEach(field => {
		if (data[field]) {
			if (Array.isArray(data[field])) {
				data[field] = marked(data[field].join('\n'));
			} else {
				data[field] = marked(data[field]);
			}
		}
	});
	console.log(data);
	return data;
}


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