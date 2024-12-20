import puppeteer from 'puppeteer';
import fs from 'fs';

// Sample Data
const invoiceData = {
    factId: '12345',
    entite: 'ABC Corporation',
    rue: '123 Business Street',
    cp: '75011',
    ville: 'Paris',
    prenom: 'John',
    nom: 'Doe',
    poste: 'Project Manager',
    mail: 'john.doe@example.com',
    today: '11/11/2024',
    acquit: 'Acquitté', // Only if "Payé" status applies
    titre_fromprog: 'Leadership Training Program',
    dates: '20/11/2024 - 22/11/2024',
    lieux: 'Paris, France',
    montant: '1500.00 €',
    paye: 'Payé' // or "retard" if not paid
};


// Compile the template with data
// const compiledHTML = Handlebars.compile(template)(data);

const {
    factId,
    entite,
    rue,
    cp,
    ville,
    prenom,
    nom,
    poste,
    mail,
    today,
    titre_fromprog,
    dates,
    lieux,
    montant,
    paye,
} = invoiceData;

const date_paiement = paye.includes("Payé") ? '11/11/2024' : '';
const moyen_paiement = paye.includes("Payé") ? 'Virement' : '';
// Load HTML Template and Compile with Handlebars
const template = `
<!DOCTYPE html>
<html lang="fr">
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
            /* border: 1px solid #333; */
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
            border-bottom: 1px solid black; 
            border-top: 1px solid #333;
        }
        .altop {
            vertical-align: top;
        }

        .totaldiv div {
            float: right;
            width: 158px;
        }

        .blnone {
            border: 1px solid #063c64;
            border-left: none;
        }

    </style>
</head>
<body>
<div style="width: 100%; text-align: center; font-size: 12px; padding-top: 10px;">
            <table style="width: 100%; border: 0;">
              <tr>
                <td style="width: 164px; text-align: left;">
                    <img src="https://github.com/IVV-FSH/templater-fsh/blob/main/assets/Logo%20FSH%20-%20transparent.png?raw=true" style="width: 160px; height: auto;">
                </td>
                <td style="width: auto; text-align: left; font-size: large;"><p class="bold">FORMATIONS</p><p>Facture n°${factId}</p></td>
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
            <p><span class="bold">${entite}</span><br>
            ${rue}<br>
            ${cp} ${ville}</p>
            <p style="margin-top: 8px;"><span class="smallcaps">stagiaire</span> : <span class="bold">${prenom} ${nom}</span><br>
            ${poste}<br>
            ${mail}</p>
        </div>
    </div>
    <table class="tablenoborder text-left table">
        <tr>
            <th>
                <h1 class="bluetext">Facture n°${factId}</h1>
                <p class="bluetext">du ${today}</p>
            </th>
            <th style="max-width:86px;" class="text-center">
                ${paye.includes("Payé") ? `<p class="redtext smtext text-center">Facture acquittée le ${date_paiement} par ${moyen_paiement}</p>` : ""}
            </th>
        </tr>
    </table>
    

    <table class="table tablenoborder" style="border: 1px solid #063c64; border-collapse: collapse;">
        <tr style="background-color: #063c64; color:white;">
            <th>Formation</th>
            <th class="text-right altop" style="border: 1px solid #063c64;">Quantité</th>
            <th class="text-right altop" style="width: 108px; border: 1px solid #063c64;">Montant TTC</th>
        </tr>
        <tr>
            <td class="altop" style="border: 1px solid #063c64;">
                <p class="bold">${titre_fromprog}</p>
                <p>Dates : ${dates}</p>
                <p>Lieu : ${lieux}</p>
            </td>
            <td class="altop text-right" style="width: 80px;border: 1px solid #063c64;">1,00</td>
            <td class="altop text-right" style="width: 108px;border: 1px solid #063c64;">${montant}</td>
        </tr>
        <tr>
            <td></td>
            <td colspan="2" class="textxxs text-center">TVA non applicable (article 239B du CGI)</td>
        </tr>
        <tr class="text-right bold">
            <td></td>
            <td colspan="2" class="text-right bold" style="background-color: #063c64; color:white;">${!paye.includes("Payé") ? "<p>Net à payer</p>" : "<p>Réglé</p>"}</td>
        </tr>
        <tr>
            <td></td>
            <td class="text-left bold" style="background-color: #063c64; color:white;">Total TTC</td>
            <td class="text-right bold" style="background-color: #063c64; color:white;">${montant}</td>
        </tr>
    </table>


    ${!paye.includes("Payé") ? `
    <table class="table-paiements" style="margin-bottom: 8px; ">
        <thead>
          <td style="width: 260px;">
            <p class="bold">Règlement par virement</p>
            <p>Indiquer la référence :</p> 
            <p class="text-center">${factId}</p>
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
            <p class="text-center">${factId}</p>
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

async function createPdfFromHtml(htmlContent) {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    // Set HTML content and wait until it is fully loaded
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

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

// Generate and save the PDF
(async () => {
    try {
        const pdfBuffer = await createPdfFromHtml(template);
        fs.writeFileSync('output.pdf', pdfBuffer);
        console.log('PDF created successfully!');
    } catch (error) {
        console.error('Error generating PDF:', error);
    }
})();
