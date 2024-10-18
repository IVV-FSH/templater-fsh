import {createReport} from 'docx-templates';
import fs from 'fs';

async function generateReport() {
  const template = fs.readFileSync('test.docx');

  const buffer = await createReport({
    template,
    cmdDelimiter: ['{{', '}}'],
    data: {
      Titre: 'John',
      Objectifs: 'Appleseed',
    },
  });

  fs.writeFileSync('report.docx', buffer);
}

generateReport().catch(console.error);