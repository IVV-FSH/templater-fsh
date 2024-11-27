import axios from 'axios';
import Fuse from 'fuse.js';
import dotenv from 'dotenv';
import stringSimilarity from 'string-similarity';

dotenv.config();
const AIRTABLE_API_KEY = process.env.AIRTABLE_TOKEN; // Get API key from environment variable
const AIRTABLE_BASE_ID = 'appPYTaiSofTygjwm'; // Replace with your Airtable Base ID
const AUTH_HEADERS = {
	Authorization: `Bearer ${AIRTABLE_API_KEY}`,
};


// Helper function to fetch all records from a table
async function fetchAllRecords(tableName, formula = null, fields = []) {
    let allRecords = [];
    let url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${tableName}`;
    const params = [];

    if (formula) {
        params.push(`filterByFormula=${encodeURIComponent(formula)}`);
    }

    if (fields.length > 0) {
        const fieldsParams = fields.map(field => `fields[]=${encodeURIComponent(field)}`).join('&');
        params.push(fieldsParams);
    }

    if (params.length > 0) {
        url += `?${params.join('&')}`;
    }

    try {
        let hasMore = true;
        let offset = null;

        while (hasMore) {
            const response = await axios.get(url + (offset ? `&offset=${offset}` : ''), {
                headers: AUTH_HEADERS,
            });

            allRecords = allRecords.concat(response.data.records.map(record => ({ id: record.id, ...record.fields })));

            offset = response.data.offset;
            hasMore = !!offset; // If there's an offset, there are more records to fetch
        }
        // console.log(allRecords)
        return allRecords;
    } catch (error) {
        handleFetchError(error);
    }
}



// Fetch all records from Personnes, Entités, and Imports (where Checked is false)
async function getData() {
    const personnesRecords = await fetchAllRecords('Personnes', null, ['Nom', 'Prénom', 'Mail']);
    const personnes = personnesRecords.map(record => ({
        fullName: `${record.Nom} ${record.Prénom}`,
        email: record.Mail,
        id: record.id,
    }));
    console.log(personnes.length)
    // console.log(personnes)
    const entites = await fetchAllRecords('Entités',null,["Dénomination"]);
    const imports = await fetchAllRecords('Imports', "NOT(Checked)");

    return {
        personnes,
        entites,
        imports,
    };
}

// Fuzzy matching function for finding the best match using Fuse.js
function findBestMatch(target, list) {
    const options = {
        includeScore: true,
        keys: ['Dénomination'], 
    };

    const fuse = new Fuse(list, options);
    const result = fuse.search(target);

    return result.length > 0 && result[0].score < 0.5 ? result[0].item : null; // Adjust score threshold as needed
}


/**
 * Processes import records by matching them with existing records and updating them in Airtable.
 * 
 * This function performs the following steps:
 * 1. Retrieves data including personnes, entites, and imports.
 * 2. Iterates over each import record and checks for the presence of 'nom' and 'prénom'.
 * 3. Constructs the full name from 'nom' and 'prénom' and finds similar records in 'personnes' based on the full name.
 * 4. Finds the best match for 'entité' in 'entites' if 'entité' exists in the import record.
 * 5. Updates the import record with potential duplicates and best match for 'entité' in Airtable.
 * 
 * @async
 * @function processImports
 * @returns {Promise<void>} A promise that resolves when all import records have been processed.
 */
export async function processImports() {
    const { personnes, entites, imports } = await getData();


    // const fuse = new Fuse(personnes, options);
    imports.forEach(async importRecord => {
        // Ensure Nom and Prénom exist in importRecord
        if (!importRecord.nom || !importRecord.prénom) {
            console.log("Import record missing name information:", importRecord);
            return; // Skip this record
        }

        const importFullName = `${importRecord.nom} ${importRecord.prénom}`;
        
        // Find similar records in personnes based on full name
        const matchedIds = personnes
            .filter(person => {
                // Ensure fullName exists in person
                if (!person.fullName) {
                    console.log("Person record missing fullName:", person);
                    return false;
                }

                const similarityScore = stringSimilarity.compareTwoStrings(importFullName, person.fullName);
                if (similarityScore > 0.7) {
                    // Log match found with similarity score
                    // console.log(`Match found: ${importFullName} is similar to ${person.fullName} with a score of ${similarityScore}`);
                    return true;
                }
                return false;
            })
            .map(person => person.id); // Only get the id of each matched person

        let entiteMatch = null;
        if(importRecord.entité) {
            entiteMatch = findBestMatch(importRecord.entité, entites);
            // console.log("entiteMatch", entiteMatch)
        }
        const entite_bestmatch = importRecord.entite_bestmatch || [];
        // const entiteMatch = findBestMatch(importRecord.entité, entites);
        const updateData = {
            "Checked": true,
            "potential duplicates": matchedIds,
            // Merge existing entite_bestmatch with new match if it exists
            "entite_bestmatch": entiteMatch ? [...new Set([...entite_bestmatch, entiteMatch.id])] : entite_bestmatch,
        };

        // Update the import record in Airtable
        await axios.patch(
            `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/Imports/${importRecord.id}`,
            { fields: updateData },
            { headers: AUTH_HEADERS }
        );

        // console.log(`Processed import record ID: ${importRecord.id}`);


    });
};
// Run the script
// await processImports().catch(console.error);

// Function to identify potential duplicates
async function findPotentialDuplicates() {
    // Retrieve data
    const { personnes, imports } = await getData();

    const potentialDuplicates = [];

    // Iterate over each import record
    imports.forEach(importRecord => {
        // Ensure Nom and Prénom exist in importRecord
        if (!importRecord.nom || !importRecord.prénom) {
            console.log("Import record missing name information:", importRecord);
            return; // Skip this record
        }

        const importFullName = `${importRecord.nom} ${importRecord.prénom}`;
        
        // Find similar records in personnes based on full name
        const matchedIds = personnes
            .filter(person => {
                // Ensure fullName exists in person
                if (!person.fullName) {
                    console.log("Person record missing fullName:", person);
                    return false;
                }

                const similarityScore = stringSimilarity.compareTwoStrings(importFullName, person.fullName);
                if (similarityScore > 0.8) {
                    // Log match found with similarity score
                    console.log(`Match found: ${importFullName} is similar to ${person.fullName} with a score of ${similarityScore}`);
                    return true;
                }
                return false;
            })
            .map(person => person.id); // Only get the id of each matched person

        // If there are matches, add them to potential duplicates
        if (matchedIds.length > 0) {
            potentialDuplicates.push({
                importRecordId: importRecord.id,
                matchedPersonIds: matchedIds,
            });
        }
    });

    console.log("Potential duplicates:", potentialDuplicates);
    return potentialDuplicates;
}

// Main function to execute the search
// (async () => {
//     await findPotentialDuplicates();
// })();

