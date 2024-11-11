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

    // Display the Lieu, lieu intra, and ampm in markdown
    output.markdown(`### Détails de la demi-journée sélectionnée:
    - **Lieu**: ${lieuass || "Non spécifié"}
    - **Lieu intra**: ${lieuIntraass || "Non spécifié"}
    - **AM/PM**: ${ampm || "Non spécifié"}`);
    
    if (sessId) {
        let debut = record.getCellValue("debut");
        let fin = record.getCellValue("fin");

        // Use selectRecordsAsync with a filter and sorting
        let query = await table.selectRecordsAsync({
            fields: ["sid", "sessId", "ampm", "Lieu", "lieu intra", "adresse"], // We only need the "sessId" field
            sorts: [{ field: "sid" }], // Sorting by "sessId" for better structure
        });
        // output.text(query.records[0].getCellValue("lieu intra"))


        let matchingRecords = query.records.filter(
            (r) => r.id !== record.id && r.getCellValueAsString("sessId") === sessId
        );
        output.markdown(`Il y a **${matchingRecords.length}** autres demi-journées avec le même **sessId**.`);
        // output.text(JSON.stringify(matchingRecords[0].getCellValue("lieu intra")));
        // output.text(JSON.stringify(matchingRecords[0].getCellValue("Lieu")));

        let matchingRecordsHours = query.records.filter(
            (r) => r.id !== record.id && r.getCellValueAsString("sessId") === sessId && r.getCellValueAsString("ampm") === ampm
        );
        output.markdown(`Il y a **${matchingRecordsHours.length}** autres demi-journées avec le même **sessId** et **ampm** (${ampm}).`);

        // Now, create buttons for the options
        let horaires = '';
        if (debut && fin) {
            // Explicitly convert debut and fin to Date objects
            let d = new Date(debut);
            let f = new Date(fin);

            // Format the debut and fin times to show only the hour and minute (HH:mm)
            horaires = `${d.getHours().toString().padStart(2, '0')}h${d.getMinutes()>0 ? d.getMinutes().toString().padStart(2, '0'):""} - ${f.getHours().toString().padStart(2, '0')}h${f.getMinutes()>0?f.getMinutes().toString().padStart(2, '0'):''}`;
        }

        // Get the Lieu (address) from the current record
        let adresse = record.getCellValue("adresse") || "Non spécifié";

        // Prompt the user with checkboxes for updating horaires and lieux
        let options = await input.buttonsAsync("Que souhaitez-vous répercuter sur les autres créneaux de la session ?", [
            { label: `Horaires: ${horaires} (${matchingRecordsHours.length} pareil)`, value: "horaires" },
            { label: `Lieux: ${adresse} (${matchingRecords.length} pareil)`, value: "lieux" },
            { label: `Horaires et Lieux`, value: "both" }
        ]);

        // Handle the selected option (horaires, lieux, or both)
        if (options === "horaires") {
            output.text("Vous avez choisi de mettre à jour les horaires.");
            // Update matching records with the horaires (debut and fin)
            for (let recordToUpdate of matchingRecordsHours) {
                let debut = record.getCellValue("debut");
                let fin = record.getCellValue("fin");

                await table.updateRecordAsync(recordToUpdate.id, {
                    "debut": debut,
                    "fin": fin,
                });
            }
        } else if (options === "lieux") {
            output.text("Vous avez choisi de mettre à jour les lieux.");
            // Update matching records with the lieu and lieu intra
            for (let recordToUpdate of matchingRecords) {
                await table.updateRecordAsync(recordToUpdate.id, {
                    "Lieu": lieu, // Correct format: array of objects with 'id' and 'name'
                    "lieu intra": lieuintra // Correct format: array of objects with 'id' and 'name'
                });
            }
} else if (options === "both") {
    output.text("Vous avez choisi de mettre à jour les horaires et les lieux.");
    // Update matching records with both horaires (debut and fin) and lieux (Lieu, lieu intra)
    for (let recordToUpdate of matchingRecordsHours) {
            await table.updateRecordAsync(recordToUpdate.id, {
                "Lieu": lieu, // Ensure it's an array with the correct object structure
                "lieu intra": lieuintra, // Ensure it's an array with the correct object structure
                "debut": debut,
                "fin": fin,
            });
    }
}
    } else {
        output.text("Aucun sessId trouvé pour cette demi-journée.");
    }
} else {
    output.text("Aucun enregistrement sélectionné.");
}
