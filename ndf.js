import dotenv from 'dotenv';

import axios from 'axios';

import nodemailer from 'nodemailer';

dotenv.config();
const AIRTABLE_API_KEY = process.env.AIRTABLE_TOKEN; // Get API key from environment variable
const AIRTABLE_BASE_ID = 'appaMzGMfnXGbGBd2'; // Replace with your Airtable Base ID
const AUTH_HEADERS = {
	Authorization: `Bearer ${AIRTABLE_API_KEY}`,
};
const transporter = nodemailer.createTransport({
	host: 'smtp-declic-php5.alwaysdata.net',
	port: 465,
	secure: true,
	auth: {
		user: 'isadora.vuongvan@sante-habitat.org',
		pass: process.env.FSH_PASSWORD
	}
});

export const getAirtableRecords = async (table, view = null, formula = null, sortField = null, sortDir = null) => {
	try {
	  let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${encodeURIComponent(table)}`;
	  const params = [];
	  if (formula) {
		params.push(`filterByFormula=${encodeURIComponent(formula)}`);
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
	  let processedData = response.data.records.map(record => ({id: record.id, ...record.fields}));
	//   console.log(`Fetched records: ${JSON.stringify(processedData)}`);
	//   processedData = processedData.map(record => {
	// 	if (Array.isArray(record)) {
	// 	  return record.join(', ');
	// 	}
	// 	return record;
	//   });
  
	// processedData = processedData.map(record => processFieldsForDocx(
	// 	record,
	// 	airtableMarkdownFields
	// ));
	// processedData['today'] = new Date().toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Paris'})

	return { records: processedData };
	} catch (error) {
	  console.error(error);
	  throw new Error('Error retrieving data from Airtable.');
	}
  };



export default async function envoiNdf() {
    console.log('Envoi NDF');
    const ndfs = await getAirtableRecords("Dépenses", "Main View", `AND({remboursé_compta} = 0, {Envoyé à compta}=0)`);
    if (ndfs.records.length > 0) {
        const firstRecord = ndfs.records[0];
		const nbRecords = ndfs.records.length;
		if(nbRecords ==0) return;
        var tableRows = '';
		var total = 0;
		var allDataAvailable = true;
		var updatedIds = [];
		for (const ndf of ndfs.records) {
			console.log(`Processing record ID: ${ndf.id}`);
			if (!ndf['Qui a payé?'] || // Updated field name
				!ndf['Date de la dépense'] ||
				!ndf['Description courte'] ||
				!ndf['Montant (€)'] ||
				!ndf['Remboursement par'] ||
				!ndf['RIB (from Qui a payé?)']
			) {
				console.log(ndf);
				console.log(`Missing data in record ID: ${ndf.id}`);
				allDataAvailable = false;
				break;
			} else {
				updatedIds.push(ndf.id);
				console.log(`Record ID: ${ndf.id} has all required data`);
			}

			const justificatifs = ndf['Justificatifs (photo/scan)'];
			var attachments = [];
			if (justificatifs && justificatifs.length > 0) {
				// const firstJustificatif = justificatifs[0];
				// console.log(firstJustificatif.thumbnails.small);
				justificatifs.forEach(justificatif => {
					attachments.push({
						filename: justificatif.filename,
						size: justificatif.size,
						type: justificatif.type,
						content: justificatif.url,
					});
				});
			}
			total += parseFloat(ndf['Montant (€)']);
			tableRows += `<tr>
				<td>${Array.isArray(ndf.quiapaye) ? ndf.quiapaye.join(', ') : (ndf.quiapaye || '')}</td>
				<td>${ndf['Date de la dépense'] || ''}</td>
				<td>${ndf['Description courte'] || ''}</td>
				<td>${ndf['Montant (€)'] ? ndf['Montant (€)'].toString().replace(".", ",") + " €" : ""}</td>
				<td>${justificatifs ? justificatifs.map(j => j.filename).join(', ') : ''}</td>
				<td>${ndf['Remboursement par'] || ''}</td>
				<td>${Array.isArray(ndf['RIB (from Qui a payé?)']) ? ndf['RIB (from Qui a payé?)'][0] : (ndf['RIB (from Qui a payé?)'] || '')}</td>
				</tr>`;
		}
			const mailOptions = {
				from: "isadora.vuongvan@sante-habitat.org",
				bcc: "isadora.vuongvan@sante-habitat.org",
				// to: "gregory.caumes@sante-habitat.org",
				// cc: ["gladys.biabiany@assoags.org"],
				subject: "Nouvelle note de frais à valider",
				html: `<style>
            table {
				border-collapse: collapse;
				table-layout: auto; /* Allow the table to adjust based on content */
			}

			th, td {
				border: 1px solid #333;
				padding: 8px;
				text-align: left;
			}

			th {
				background-color: #f2f2f2;
			}

			tfoot {
				background-color: #f2f2f2;
				border: none;
			}

			p {
				font-family: 'Segoe UI', sans-serif;
			}
        </style>
				<p>Hello Grégory,</p>
				<p>Peux-tu s'il te plaît valider ${nbRecords >1 ? 'ces notes' : 'cette note'} de frais pour que Gladys, en copie, puisse me rembourser ?</p>
				<table>
					<thead>
						<tr>
							<th>Dépense effectuée par</th>
							<th>Date de la dépense</th>
							<th>Description courte</th>
							<th>Montant</th>
							<th>Justificatif(s)</th>
							<th>Rembourser par</th>
							<th>RIB</th>
						</tr>
					</thead>
					<tbody>
					${tableRows}
					</tbody>
					<tfoot>
						<tr>
							<td colspan="3" style="text-align: right; font-weight: bold;">Total</td>
							<td style="text-align: right; font-weight: bold;">${new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' }).format(total)}</td>
							<td colspan="3"></td>
						</tr>
					</tfoot>
				</table>
				<p>Justificatifs en pièce jointe</p>

				<p>Merci par avance !</p>
				<p>Isadora</p>`,
				attachments: attachments
			};
			try {
				await transporter.sendMail(mailOptions);
				console.log('Email sent');

				try {
					// Update records in Airtable
					const updateUrl = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/Dépenses`;
					const today = new Date().toISOString().split('T')[0]; // Ensure the date is in YYYY-MM-DD format
					const updateData = {
						records: updatedIds.map(id => ({
							id,
							fields: {
								'date_envoicompta': today,
							},
						})),
					};
					await axios.patch(updateUrl, updateData, {
						headers: AUTH_HEADERS,
					});
					console.log('Records updated');					
				} catch (error) {
					console.error(error);
				}

			} catch (error) {
				console.error(error);
			}


    // Envoyer un mail à responsable, copie la compta
    // Joindre tous les justificatifs ("Justificatifs (photo/scan)")
    // Demander de valider les dépenses (par écrit puis TODO: en cliquant sur un lien qui valide et envoie à la compta)
    } else {
        console.log('No records found.');
    }

}

// run envoiNdf
await envoiNdf();
