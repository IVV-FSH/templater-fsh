import express from 'express';
import { createReport } from 'docx-templates';
import archiver from 'archiver';
import fs from 'fs';
import path from 'path';
import { Readable } from 'stream';
import { fetchTemplate } from './utils';
import { GITHUBTEMPLATES } from './constants';

const bufferToStream = (buffer) => {
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null); // Signal the end of the stream
  return stream;
};
const app = express();


const createBufferFile = async (filename, content, titre) => {
    const buffer = await createReport({
        output: 'buffer',
        template: fs.readFileSync(path.join('templates', 'test.docx')),
        data: { Titre: content }
    });
    return {
        filename: filename,
        content: buffer // Create a buffer from the string content
    };
    // return {
    //     filename: filename,
    //     content: Buffer.from(content, 'utf-8') // Create a buffer from the string content
    // };
};


app.get('/download', async (req, res) => {
    // const templatePath = 'templates/test.docx'; // Path to your template file
    // const template = await fs.promises.readFile(templatePath);
    const template = await fetchTemplate(GITHUBTEMPLATES+'test.docx');

    const titles = ['Title for 1.docx', 'Title for 2.docx'];
    const zipFileName = 'documents.zip';

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename=${zipFileName}`);

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(res);
    try {
        for (let i = 0; i < titles.length; i++) {
            const { filename, content } = await createBufferFile(`document${i + 1}.docx`, titles[i]);
            const stream = bufferToStream(content);
            // archive.append(content, { name: filename });
            archive.append(stream, { name: filename });
        }
        await archive.finalize();
    } catch (error) {
        console.error('Error creating zip archive:', error);
        res.status(500).send('Error creating zip archive');
    }
    // try {
    //     for (let i = 0; i < titles.length; i++) {
    //         // Function to create a simple buffer file

    //         // const buffer = await createReport({
    //         //     output: 'buffer',
    //         //     template,
    //         //     data: { Titre: titles[i] }, // Insert title into the document
    //         //     // cmdDelimiter: '+++',
    //         // });

    //         const buffer = await createBufferFile(`${i + 1}.docx`, 'Hello', titles[i]); // Create buffer with the content "Hello

    //         // // Debugging: log the type of buffer
    //         // console.log(`Buffer type for ${titles[i]}:`, typeof buffer);
    //         // console.log(`Buffer instance check for ${titles[i]}:`, Buffer.isBuffer(buffer));

    //         // if (!Buffer.isBuffer(buffer)) {
    //         //     throw new Error(`Generated report for ${titles[i]} is not a valid Buffer`);
    //         // }

    //         // const docxFileName = `${i + 1}.docx`;
    //         // archive.append(buffer, { name: docxFileName });
    //         archive.append(buffer.content, { name: buffer.filename });
    //     }


    //     await archive.finalize();
    // } catch (error) {
    //     console.error('Error generating documents:', error);
    //     res.status(500).send('Error generating documents');
    // }
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
