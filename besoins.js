import { getAirtableRecords } from "./utils.js";
// import dotenv from 'dotenv';

// dotenv.config();
// const AIRTABLE_API_KEY = process.env.AIRTABLE_TOKEN; // Get API key from environment variable
// const AIRTABLE_BASE_ID = 'appK5MDuerTOMig1H'; // Replace with your Airtable Base ID
// const AUTH_HEADERS = {
// 	Authorization: `Bearer ${AIRTABLE_API_KEY}`,
// };

export async function getBesoins(besoinsData) {
    const questionsCg = [
        {
            intitule: "Pour quelle(s)s raison(s) souhaitez-vous suivre cette formation ?",
            fieldType: "array", // Answers include lists
            fieldName: "Pour quelle(s)s raison(s) souhaitez-vous suivre cette formation ?"
        },
        {
            intitule: "Avez-vous d'autres raisons ?",
            fieldType: "text",
            fieldName: "raisons_autres"
        },
        {
            intitule: "Quelles sont vos attentes prioritaires en participant à cette formation ?",
            fieldType: "array", // Answers include lists and quoted elements
            fieldName: "Quelles sont vos attentes prioritaires en participant à cette formation ?"
        },
        {
            intitule: "Avez-vous d'autres attentes ?",
            fieldType: "text",
            fieldName: "attentes_autres"
        },
        {
            intitule: "Quels sont les 3 critères les plus importants pour vous en assistant à cette formation ?",
            fieldType: "array", // Answers include quoted elements
            fieldName: "Quels sont les 3 critères les plus importants pour vous en assistant à cette formation ?"
        },
        {
            intitule: "A l’issue de cette formation, avez-vous un projet à court moyen ou long terme ?",
            fieldType: "array",
            fieldName: "A l’issue de cette formation, avez-vous un projet à court moyen ou long terme ? "
        },
        {
            intitule: "Veuillez expliquer votre projet",
            fieldType: "text",
            fieldName: "projet_plus"
        },
    ];


    const questionsFsh = [
        {
            intitule: "Qu’attendez-vous de cette formation ?",
            fieldType: "array",
            fieldName: "Qu’attendez-vous de cette formation ?"
        },
        {
            intitule: "Qu’en attendez-vous en priorité ? ex : objectifs, méthodes, outils, contenu des apports...",
            fieldType: "text",
            fieldName: "Qu’en attendez-vous en priorité ? ex : objectifs, méthodes, outils, contenu des apports..."
        },
        {
            intitule: "Avez-vous déjà suivi une formation sur ce thème ou un thème en rapport ? Si oui laquelle ?",
            fieldType: "text",
            fieldName: "Avez-vous déjà suivi une formation sur ce thème ou un thème en rapport ? Si oui laquelle ?"
        },
        {
            intitule: "Veuillez évaluer vos connaissances sur la thématique",
            fieldType: "rating",
            fieldName: "Veuillez évaluer vos connaissances sur la thématique"
        },
        {
            intitule: "Veuillez évaluer vos compétences sur la thématique",
            fieldType: "rating",
            fieldName: "Veuillez évaluer vos compétences sur la thématique"
        },
        {
            intitule: "Quelles difficultés rencontrez-vous sur le terrain ?",
            fieldType: "text",
            fieldName: "Quelles difficultés rencontrez-vous sur le terrain ?"
        },
        {
            intitule: "Avez-vous un cas concret pour lequel vous souhaiteriez des éclaircissements ?",
            fieldType: "text",
            fieldName: "Avez-vous un cas concret pour lequel vous souhaiteriez des éclaircissements ?"
        },
    ];

    const questionsFormassad = [
        {
            intitule: "Quel poste occupez-vous au sein de la structure, quelles sont vos missions principales ?",
            fieldType: "text",
            fieldName: "Quel poste occupez-vous au sein de la structure, quelles sont vos missions principales ?"
        },
        {
            intitule: "Avez-vous déjà suivi une formation sur ce thème ou un thème en rapport ? Si oui laquelle",
            fieldType: "text",
            fieldName: "Avez-vous déjà suivi une formation sur ce thème ou un thème en rapport ? Si oui laquelle"
        },
        {
            intitule: "Qu’en attendez-vous en priorité ? ex : objectifs, méthodes, outils, contenu des apports...",
            fieldType: "text",
            fieldName: "Qu’en attendez-vous en priorité ? ex : objectifs, méthodes, outils, contenu des apports..."
        },
        {
            intitule: "Quelles difficultés rencontrez-vous sur le terrain ?",
            fieldType: "text",
            fieldName: "Quelles difficultés rencontrez-vous sur le terrain ?"
        },
        {
            intitule: "Avez-vous un cas concret pour lequel vous souhaiteriez des éclaircissements ?",
            fieldType: "text",
            fieldName: "Avez-vous un cas concret pour lequel vous souhaiteriez des éclaircissements ?"
        },
    ];

    const questionsAll = [
        {
            intitule: "Avez-vous besoin d’un aménagement spécifique au regard de votre situation personnelle (situation de handicap, régime alimentaire…)",
            fieldType: "text",
            fieldName: "Avez-vous besoin d’un aménagement spécifique au regard de votre situation personnelle (situation de handicap, régime alimentaire…)"
        },
        {
            intitule: "Note personnelle à l’attention de l’intervenant",
            fieldType: "text",
            fieldName: "Note personnelle à l’attention de l’intervenant"
        }
    ];
    //   const { sessId, formateurId } = req.query;
    // console.log("Fetching besoins for session:", sessId);
    // console.log("Fetched besoins:", besoinsData);
    if(!besoinsData.records || besoinsData.records.length === 0) {
        console.log("No besoins found.");
        return {html:""};
    }
    const type = besoinsData.records[0].Type;
    console.log("Determined type:", type);

    var questions = [];
    switch (type) {
        case "CG":
            questions = questionsCg;
            break;
        case "FSH":
            questions = questionsFsh;
            break;
        case "Formassad":
            questions = questionsFormassad;
            break;
        default:
            questions = questionsFsh;
            break;
    }
    questions = [...questions, questionsAll[1]];
    var answersHtml = "";

    var recapHtml = "";
    var answerCounters = {};

    besoinsData.records.filter(b=> b.rempli=="🟢").forEach(besoin => {
        var answers = "";

        questions.forEach(question => {
            // console.log("Question:", question.intitule);
            if(question.fieldType !== "array") {
                // console.log(`Q: ${question.intitule} -- A: ${besoin[question.intitule]}\n`);
                if(besoin[question.fieldName]) {
                    answers += `<p style="font-weight:bold">${question.intitule}</p><p>${besoin[question.fieldName]}${question.fieldType == "rating" ? "/10" : ""}</p>`;
                } else {
                    // console.log(`Q: ${question.intitule}\nA: -\n\n`);
                    // answers += `<p style="font-weight:bold">${question.intitule}</p><p>A: ${besoin[question.intitule]}</p><br>`;
                }
                
            } else {
                if (!answerCounters[question.intitule]) {
                    answerCounters[question.intitule] = {};
                }
                if (!Array.isArray(besoin[question.fieldName])) {
                    besoin[question.fieldName] = [besoin[question.fieldName]];
                    // console.warn(`Expected an array for question "${question.intitule}" but got:`, recapAnswers);
                }
                var recapAnswers = besoin[question.fieldName];
                
                // Ensure the value is an array
                recapAnswers.forEach(answer => {
                    if (!answerCounters[question.intitule][answer]) {
                        answerCounters[question.intitule][answer] = 0;
                    }
                    answerCounters[question.intitule][answer]++;
                });
            
                if (besoin[question.fieldName]) {
                    answers += `<p style="font-weight:bold">${question.intitule}</p><p>${besoin[question.fieldName].map(b => `<span>${b}</span>`).join("<br>")}</p>`;
                }

            }
            // if(besoin[question.intitule]) {
            //     // console.log(`Q: ${question.intitule}\nA: ${besoin[question.intitule]}\n\n`);
            //     answers += `<p style="font-weight:bold">${question.intitule}</p><p>${besoin[question.intitule]}</p>`;
            // } else {
            //     // console.log(`Q: ${question.intitule}\nA: -\n\n`);
            //     answers += `<p style="font-weight:bold">${question.intitule}</p><p>A: ${besoin[question.intitule]}</p><br>`;
            // }
        });

        answersHtml += `<div class="fiche-besoin">
        <h3>Apprenant.e: ${besoin["prenom (from Inscrits)"][0]} ${besoin["nom (from Inscrits)"][0]}${besoin["poste (from Inscrits)"][0] ? `, <span class="font-light">${besoin["poste (from Inscrits)"][0]}` : ""}</span></h3>
        ${answers}
        </div>`;
    });

// Generate recapHtml
Object.entries(answerCounters).forEach(([question, answers]) => {
    // Sort answers by count in descending order
    const sortedAnswers = Object.entries(answers).sort((a, b) => b[1] - a[1]);

    // Add the question in bold
    recapHtml += `<p style="font-weight:bold">${question}</p><ul>`;

    // Add each answer and its count
    sortedAnswers.forEach(([answer, count]) => {
        recapHtml += `<li>${answer}: ${count}</li>`;
    });

    recapHtml += `</ul>`;
});

// // Log the result for each question of type "array"
// Object.entries(answerCounters).forEach(([question, answers]) => {
//     console.log(`Question: ${question}`);
//     Object.entries(answers).forEach(([answer, count]) => {
//         console.log(`  ${answer}: ${count}`);
//     });
// });

    return {
        html: answersHtml,
        type: type,
        titre: besoinsData.records[0]["titre_fromprog (from Inscrits)"],
        recap: recapHtml,
    };
}

// await getBesoins("recDxQgO7JThS9toQ");