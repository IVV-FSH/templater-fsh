let table = base.getTable("Demi-journées");

// Prompt the user to select a record
let record = await input.recordAsync("Sélectionner une demi-journée", table);

if (record) {
    let sessId = record.getCellValueAsString("sessId");
    let ampm = record.getCellValueAsString("ampm");
    let lieuass = record.getCellValueAsString("Lieu");
    let lieuIntraass = record.getCellValueAsString("lieu intra");
    let lieu = record.getCellValue("Lieu");
    let lieuintra = record.getCellValue("lieu intra");
    let debut = record.getCellValue("debut");
    let fin = record.getCellValue("fin");
    let sessions = record.getCellValue("Sessions"); // Assuming "Sessions" is the field to copy
    let recordDate = debut ? new Date(debut) : null;

        output.markdown(`### Détails de la demi-journée sélectionnée:
    - **Lieu**: ${lieuass || "Non spécifié"}
    - **Lieu intra**: ${lieuIntraass || "Non spécifié"}
    - **AM/PM**: ${ampm || "Non spécifié"}
    - ${debut}`
    );

    let query = await table.selectRecordsAsync({
        fields: ["sessId", "ampm"], // We only need the "sessId" field
    });
    let other = ampm == "am" ? "pm" : "am";
    let otherhalf = query.records.filter(
        (r) => r.id !== record.id && r.getCellValueAsString("sessId") === sessId && r.getCellValueAsString("ampm") == other
    );
    let hasOtherHalf = otherhalf.length > 0;

    if(debut && fin && !hasOtherHalf) {
        let newDebut;
        let newFin;
        let d = new Date(debut);
        let f = new Date(fin);
        const hoursToMs = 3.5 * 60 * 60 * 1000; // 3.5 hours in milliseconds
        if (ampm === "am") {
            newDebut = new Date(debut);
            newDebut.setHours(newDebut.getHours() + 1); // Adds 1 hour to the "debut" time

            newFin = new Date(newDebut.getTime() + hoursToMs); // Adds 3.5 hours to newDebut
        } else {
            newFin = new Date(fin);
            newFin.setHours(newFin.getHours() - 1); // Subtracts 1 hour from the "fin" time

            newDebut = new Date(newFin.getTime() - hoursToMs); // Subtracts 3.5 hours from newFin
        }
        await table.createRecordAsync({
            "debut": newDebut,
            "fin": newFin,
            "Sessions": sessions,
        });
    }
    
}
//     if (!recordDate) {
//         continue; // Skip if there's no debut date
//     }

//     // Get the date part of the debut field (ignoring the time)
//     let dateStr = recordDate.toISOString().split('T')[0]; // YYYY-MM-DD format

//     if (ampm === "am") {
//         // Check if there's a "PM" record for the same day
//         let pmRecords = query.records.filter(r => {
//             let rDate = r.getCellValue("debut") ? new Date(r.getCellValue("debut")) : null;
//             return rDate && rDate.toISOString().split('T')[0] === dateStr && r.getCellValue("ampm") === "pm";
//         });

//         if (pmRecords.length === 0) {
//             // No PM record exists, so create one
//             let newFin = new Date(fin);
//             newFin.setHours(newFin.getHours() + 1); // Add one hour to the "fin" time

//             // Create a new record for "PM"
//             await table.createRecordAsync({
//                 "debut": newFin.toISOString(),
//                 "fin": new Date(newFin).setHours(newFin.getHours() + 7).toISOString(), // Setting a reasonable "fin" time for PM, e.g., 7 hours later
//                 "Sessions": sessions // Copy "Sessions" from the original record
//             });
//             output.text(`Created new PM record for ${dateStr}`);
//         }
//     } else if (ampm === "pm") {
//         // Check if there's an "AM" record for the same day
//         let amRecords = query.records.filter(r => {
//             let rDate = r.getCellValue("debut") ? new Date(r.getCellValue("debut")) : null;
//             return rDate && rDate.toISOString().split('T')[0] === dateStr && r.getCellValue("ampm") === "am";
//         });

//         if (amRecords.length === 0) {
//             // No AM record exists, so create one
//             let newDebut = new Date(debut);
//             newDebut.setHours(newDebut.getHours() - 1); // Subtract one hour from the "debut" time

//             // Create a new record for "AM"
//             await table.createRecordAsync({
//                 "debut": newDebut.toISOString(),
//                 "fin": new Date(newDebut).setHours(newDebut.getHours() + 7).toISOString(), // Setting a reasonable "fin" time for AM, e.g., 7 hours later
//                 "Sessions": sessions // Copy "Sessions" from the original record
//             });
//             output.text(`Created new AM record for ${dateStr}`);
//         }
//     }
// }
