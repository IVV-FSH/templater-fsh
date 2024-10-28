// import dotenv from 'dotenv';
// import brevo from '@getbrevo/brevo';
// import axios from 'axios';
import dotenv from 'dotenv';
import nodemailer from 'nodemailer';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';
import {createReport} from 'docx-templates';

// Load environment variables from .env file
dotenv.config();

const sender = "isadora.vuongvan@sante-habitat.org";
const recipient = "isadoravuongvan@gmail.com";

// Get __dirname equivalent
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Create a transporter
const transporter = nodemailer.createTransport({
    host: 'smtp-declic-php5.alwaysdata.net',
    port: 465,
    secure: true,
    auth: {
        user: sender,
        pass: 'Renée_Fédér@tion_75'
    }
});

// Function to create a simple buffer file
const createBufferFile = async (filename, content) => {
    const buffer = await createReport({
        output: 'buffer',
        template: fs.readFileSync(path.join('templates', 'test.docx')),
        data: { Titre: 'Hello' }
    });
    return {
        filename: filename,
        content: buffer
    };
    // return {
    //     filename: filename,
    //     content: Buffer.from(content, 'utf-8') // Create a buffer from the string content
    // };
};

// Function to send email
export const sendEmail = async (to, subject, text) => {
    const filePath = path.join('templates', 'test.docx');

    // Check if the .docx file exists
    if (!fs.existsSync(filePath)) {
        console.error('Attachment file not found:', filePath);
        return;
    }

    // Create a buffer file for hello.txt
    const helloBufferFile = await createBufferFile('hello.docx', 'Hello'); // Create buffer with the content "Hello"
    // const helloBufferFile = await createBufferFile('hello.txt', 'Hello'); // Create buffer with the content "Hello"

    const mailOptions = {
        from: sender,
        to: to,
        subject: subject,
        text: text,
        attachments: [
            {
                filename: 'test.docx',
                path: filePath // path to the .docx attachment
            },
            helloBufferFile // Attach the hello.txt buffer file
        ]
    };

    try {
        const info = await transporter.sendMail(mailOptions);
        console.log('Email sent: ', info.response);
    } catch (error) {
        console.error('Error sending email:', error);
    }
};

// Main function to send the email
const main = async () => {
    await sendEmail(recipient, 'Test Email with Attachments', 'Hello, this is a test email with attachments.');
};

main().catch(console.error);


// // Email details
// const apiKey = process.env.BREVO_API_KEY;
// const emailData = {
//     sender: {
//         name: "Isadora Vuong Van", // You can customize the sender's name here
//         email: sender
//     },
//     to: [
//         {
//             email: recipient,
//             name: "Isadora Vuong Van" // You can customize the recipient's name here
//         }
//     ],
//     subject: "Hello world",
//     htmlContent: "<html><head></head><body><p>Hello,</p><p>This is my first transactional email sent from Brevo.</p></body></html>"
// };

// // Send the email using Brevo API
// axios.post('https://api.brevo.com/v3/smtp/email', emailData, {
//     headers: {
//         'accept': 'application/json',
//         'api-key': apiKey,
//         'content-type': 'application/json'
//     }
// })
// .then(response => {
//     console.log('Email sent successfully:', response.data);
// })
// .catch(error => {
//     console.error('Error sending email:', error.response ? error.response.data : error.message);
// });

