// let table = base.getTable('Sessions');
// let selectedRecord = await input.recordAsync('Select the Session record', table);
// if (!selectedRecord) {
//     output.text('No session record selected. Exiting script.');
// } else {
//     // Prompt for start date and end date (in YYYY-MM-DD format)
//     let startDate = await input.textAsync('Enter the start date (YYYY-MM-DD)');
//     let endDate = await input.textAsync('Enter the end date (YYYY-MM-DD)');

//     // Offer time options
//     let timeOptions = ['0930-1300 1400-1730', '0900-1230 1330-1700'];
//     let selectedTimes = await input.buttonsAsync('Select time slots', timeOptions);

//     // Set time slots based on user selection
//     let startMorning, endMorning, startAfternoon, endAfternoon;
//     if (selectedTimes === '0930-1300 1400-1730') {
//         startMorning = '09:30';
//         endMorning = '13:00';
//         startAfternoon = '14:00';
//         endAfternoon = '17:30';
//     } else {
//         startMorning = '09:00';
//         endMorning = '12:30';
//         startAfternoon = '13:30';
//         endAfternoon = '17:00';
//     }
    
//     // Ask the user for the time type
//     let timeTypeOptions = ['Matins seulement', 'AM seulement', 'Toute la journée'];
//     let selectedTimeType = await input.buttonsAsync('Select the time type', timeTypeOptions);

//     // List of timezones for France and its DOM-TOM
//     const timezones = {
//         'France': 'Europe/Paris',
//         'Guadeloupe': 'America/Guadeloupe',
//         'Martinique': 'America/Martinique',
//         'Réunion': 'Indian/Reunion',
//         'Mayotte': 'Indian/Mayotte',
//         'French Guiana': 'America/Cayenne',
//         'Saint Pierre and Miquelon': 'America/Miquelon',
//         'Saint Martin': 'America/Marigot',
//         'Saint Barthélemy': 'America/Guadeloupe'
//     };

//     // Prompt for timezone selection using buttonsAsync
//     let timezoneOptions = Object.keys(timezones);
//     let selectedTimezone = await input.buttonsAsync('Select your timezone', timezoneOptions);

//     // Set default to France if the user does not select another option
//     if (!selectedTimezone) {
//         selectedTimezone = 'France'; // Default to France
//     }
//     let selectedTimezoneValue = timezones[selectedTimezone];


//     // Get available options from the 'Lieu' field in the 'Demi-journées' table
//     let lieuField = base.getTable('Demi-journées').getField('Lieu');
//     let lieuOptions = lieuField.options.choices.map(choice => choice.name);
//     let lieu = await input.buttonsAsync('Select the location', lieuOptions);

//     let lieu_intra;

//     if (lieu === 'En intra') {
//         let tableEtab = base.getTable('Etablissements');

//         // Initialize a variable to store the selected record
//         let selectedEtab = null;

//         while (!selectedEtab) {
//             // Prompt the user to select an Etablissement record
//             selectedEtab = await input.recordAsync('Select an Etablissement', tableEtab);

//             // Check if a record was selected
//             if (selectedEtab) {      
//                 // Output the selected Etablissement's name
//                 output.text(`You selected: ${selectedEtab.getCellValue('id')}`);
//             } else {
//                 // If no record is selected, show the message and link to create a new Etablissement
//                 output.markdown(`Click [here](https://airtable.com/appPYTaiSofTygjwm/pag1x39Vr1uGQi0zn/form) to create a new Etablissement.`);
                
//                 // Show a button for acknowledging the addition
//                 await input.buttonsAsync('Etablissement ajouté.', [
//                     'Etablissement ajouté'
//                 ]);
                
//                 // Exit the loop after the user clicks the button
//                 // break;
//             }
//         }

//         // let etablissementsTable = base.getTable('Etablissements');
        
//         // // Select a record from the Etablissements table
//         // let selectedEtablissementRecord = await input.recordAsync('Select an Etablissement', etablissementsTable);
        
//         // // Check if a record was selected
//         // if (selectedEtablissementRecord) {
//         //     lieu = selectedEtablissementRecord.getCellValue('Nom'); // Get the name of the selected Etablissement
            
//         //     // Optional: Validate if the selected lieu is acceptable
//         //     if (!acceptedLieux.includes(lieu)) {
//         //         output.text(`The selected Etablissement "${lieu}" is not valid.`);
//         //         return; // Exit or handle the error accordingly
//         //     }
//         // } else {
//         //     output.text('No Etablissement selected. Operation canceled.');

//         // }
//     }



//     // Confirm details before creating records
//     let confirm = await input.buttonsAsync('Confirm details:', [
//         `Dates: ${startDate} to ${endDate}`,
//         `Morning: ${startMorning} - ${endMorning}`,
//         `Afternoon: ${startAfternoon} - ${endAfternoon}`,
//         `Lieu: ${lieu}`,
//         'Confirm', 'Cancel'
//     ]);

//     if (confirm === 'Cancel') {
//         output.text('Operation canceled.');
//     } else {

//         // Create records in the "Demi-journées" table for each half-day in the date range
//         let demiJourneesTable = base.getTable('Demi-journées');

//         // Adjust for the Europe/Paris timezone using Intl.DateTimeFormat
//         let timeZone = 'Europe/Paris';

//         let currentDate = new Date(startDate);
//         let endDateObj = new Date(endDate);

//         let options = { timeZone, hour12: false };

//         while (currentDate <= endDateObj) {
//             let currentDateString = currentDate.toISOString().split('T')[0];

//             // Convert start and end times to the selected timezone
//             function toTimezone(date, time, timezone) {
//                 let dateTimeStr = `${date}T${time}:00`;
//                 return new Date(new Date(dateTimeStr).toLocaleString('en-US', { timeZone: timezone }));
//             }

//             let morningStart = toTimezone(currentDateString, startMorning, selectedTimezoneValue);
//             let morningEnd = toTimezone(currentDateString, endMorning, selectedTimezoneValue);
//             let afternoonStart = toTimezone(currentDateString, startAfternoon, selectedTimezoneValue);
//             let afternoonEnd = toTimezone(currentDateString, endAfternoon, selectedTimezoneValue);

//             if (selectedTimeType === 'Matins seulement' || selectedTimeType === 'Toute la journée') {
//                 // Create morning session
//                 await demiJourneesTable.createRecordAsync({
//                     'Sessions': [{ id: selectedRecord.id }],
//                     'Début': morningStart,  // Adjusted for the selected timezone
//                     'Fin': morningEnd,      // Adjusted for the selected timezone
//                     'Lieu': { name: lieu }
//                 });
//             }

//             if (selectedTimeType === 'AM seulement' || selectedTimeType === 'Toute la journée') {
//                 // Create afternoon session
//                 await demiJourneesTable.createRecordAsync({
//                     'Sessions': [{ id: selectedRecord.id }],
//                     'Début': afternoonStart,  // Adjusted for the selected timezone
//                     'Fin': afternoonEnd,      // Adjusted for the selected timezone
//                     'Lieu': { name: lieu }
//                 });
//             }

//             currentDate.setDate(currentDate.getDate() + 1);
//         }

//     }


//     output.text('Records created successfully!');
// }


let table = base.getTable('Sessions');
let selectedRecord = await input.recordAsync('Select the Session record', table);
if (!selectedRecord) {
    output.text('No session record selected. Exiting script.');
} else {
    // Prompt for start date and end date, allowing flexible formats
    let startDateInput = await input.textAsync('Enter the start date (e.g., 1/10 or 1/10/24)');
    let endDateInput = await input.textAsync('Enter the end date (e.g., 2/10 or 2/10/24)');

    // Function to parse date inputs
    function parseDate(dateInput) {
        let dateParts = dateInput.split('/');
        let currentYear = new Date().getFullYear();
        let day = dateParts[0].padStart(2, '0'); // Ensure day is two digits
        let month = dateParts[1].padStart(2, '0'); // Ensure month is two digits
        let year = dateParts[2] ? `20${dateParts[2]}` : currentYear; // Default to current year if not provided
        return `${year}-${month}-${day}`; // Format as YYYY-MM-DD
    }

    // Convert start and end dates to proper format
    let startDate = parseDate(startDateInput);
    let endDate = parseDate(endDateInput);

    // Offer time options
    let timeOptions = ['0930-1300 1400-1730', '0900-1230 1330-1700'];
    let selectedTimes = await input.buttonsAsync('Select time slots', timeOptions);

    // Set time slots based on user selection
    let startMorning, endMorning, startAfternoon, endAfternoon;
    if (selectedTimes === '0930-1300 1400-1730') {
        startMorning = '09:30';
        endMorning = '13:00';
        startAfternoon = '14:00';
        endAfternoon = '17:30';
    } else {
        startMorning = '09:00';
        endMorning = '12:30';
        startAfternoon = '13:30';
        endAfternoon = '17:00';
    }

    // Ask the user for the time type
    let timeTypeOptions = ['Matins seulement', 'AM seulement', 'Toute la journée'];
    let selectedTimeType = await input.buttonsAsync('Select the time type', timeTypeOptions);

    // List of timezones for France and its DOM-TOM
    const timezones = {
        'France': 'Europe/Paris',
        'Guadeloupe': 'America/Guadeloupe',
        'Martinique': 'America/Martinique',
        'Réunion': 'Indian/Reunion',
        'Mayotte': 'Indian/Mayotte',
        'French Guiana': 'America/Cayenne',
        'Saint Pierre and Miquelon': 'America/Miquelon',
        'Saint Martin': 'America/Marigot',
        'Saint Barthélemy': 'America/Guadeloupe'
    };

    // Prompt for timezone selection using buttonsAsync
    let timezoneOptions = Object.keys(timezones);
    let selectedTimezone = await input.buttonsAsync('Select your timezone', timezoneOptions);

    // Set default to France if the user does not select another option
    if (!selectedTimezone) {
        selectedTimezone = 'France'; // Default to France
    }
    let selectedTimezoneValue = timezones[selectedTimezone];

    // Get available options from the 'Lieu' field in the 'Demi-journées' table
    let lieuField = base.getTable('Demi-journées').getField('Lieu');
    let lieuOptions = lieuField.options.choices.map(choice => choice.name);
    let lieu = await input.buttonsAsync('Select the location', lieuOptions);

    let lieu_intra = null; // Initialize variable for intra location

    if (lieu === 'En intra') {
        let tableEtab = base.getTable('Etablissements');

        // Loop until a valid Etablissement is selected
        while (!lieu_intra) {
            // Prompt the user to select an Etablissement record
            let selectedEtab = await input.recordAsync('Select an Etablissement', tableEtab);

            if (selectedEtab) {
                // Store the selected Etablissement's name in lieu_intra
                lieu_intra = selectedEtab;
                output.text(`You selected: ${selectedEtab.getCellValue('Nom')}`);
            } else {
                // If no record is selected, show the message and link to create a new Etablissement
                output.markdown(`Click [here](https://airtable.com/appPYTaiSofTygjwm/pag1x39Vr1uGQi0zn/form) to create a new Etablissement.`);
                
                // Wait for user confirmation
                await input.buttonsAsync('Etablissement ajouté.', [
                    'Etablissement ajouté'
                ]);
            }
        }
    }

    // Display summary before final confirmation
    output.markdown(`
    ### Review your selections:
    - **Dates**: ${startDate} to ${endDate}
    - **Morning**: ${startMorning} - ${endMorning}
    - **Afternoon**: ${startAfternoon} - ${endAfternoon}
    - **Lieu**: ${lieu_intra ? lieu_intra : lieu}
    `);

    let confirm = await input.buttonsAsync('Please confirm:', [
        'Confirm', 
        'Cancel'
    ]);

    if (confirm === 'Cancel') {
        output.text('Operation canceled.');
    } else {

        // Create records in the "Demi-journées" table for each half-day in the date range
        let demiJourneesTable = base.getTable('Demi-journées');

        // Adjust for the selected timezone using Intl.DateTimeFormat
        let currentDate = new Date(startDate);
        let endDateObj = new Date(endDate);

        while (currentDate <= endDateObj) {
            let currentDateString = currentDate.toISOString().split('T')[0];

            // Convert start and end times to the selected timezone
            function toTimezone(date, time, timezone) {
                let dateTimeStr = `${date}T${time}:00`;
                return new Date(new Date(dateTimeStr).toLocaleString('en-US', { timeZone: timezone }));
            }

            let morningStart = toTimezone(currentDateString, startMorning, selectedTimezoneValue);
            let morningEnd = toTimezone(currentDateString, endMorning, selectedTimezoneValue);
            let afternoonStart = toTimezone(currentDateString, startAfternoon, selectedTimezoneValue);
            let afternoonEnd = toTimezone(currentDateString, endAfternoon, selectedTimezoneValue);

            if (selectedTimeType === 'Matins seulement' || selectedTimeType === 'Toute la journée') {
                
                // Create the initial options object for the morning session
                let morningOptions = {
                    'Sessions': [{ id: selectedRecord.id }],
                    'debut': morningStart,  // Adjusted for the selected timezone
                    'fin': morningEnd,      // Adjusted for the selected timezone
                    'Lieu': { name: lieu }  // Use the selected Lieu (Visioconférence, En intra, etc.)
                };

                // If "En intra" is selected, populate the linked "lieu_intra" field
                if (lieu === 'En intra' && lieu_intra) {
                    // Add the linked record from the Etablissements table
                    morningOptions['lieu intra'] = [{ id: lieu_intra.id }];
                }

                // Create the morning session record
                await demiJourneesTable.createRecordAsync(morningOptions);
            }

            if (selectedTimeType === 'AM seulement' || selectedTimeType === 'Toute la journée') {
                
                let afternoonOptions = {
                    'Sessions': [{ id: selectedRecord.id }],
                    'debut': afternoonStart,  // Adjusted for the selected timezone
                    'fin': afternoonEnd,      // Adjusted for the selected timezone
                    'Lieu': { name: lieu }  // Use the selected Lieu
                };

                if (lieu === 'En intra' && lieu_intra) {
                    afternoonOptions['lieu intra'] = [{ id: lieu_intra.id }];
                }

                await demiJourneesTable.createRecordAsync(afternoonOptions);
            }


            currentDate.setDate(currentDate.getDate() + 1);
        }
    }

    output.text('Records created successfully!');
}
