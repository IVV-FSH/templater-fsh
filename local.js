import {createReport} from 'docx-templates';
import fs from 'fs';
import { marked } from 'marked'; // Add this import

async function generateReport() {
  const templatePath = path.join('templates', 'html.docx');
  const template = fs.readFileSync(templatePath);
  
  const buffer = await createReport({
    template,
    // cmdDelimiter: ['{{', '}}'],
    data: {
      Titre: 'John',
      film: {
        title: 'Inception',
        releaseDate: '2010-07-16',
        feature1: 'Mind-bending plot',
        feature2: 'Stunning visuals',
        feature3: 'Great soundtrack',
        description: marked('# A thief who steals corporate secrets\n\n* Mind-bending plot\n* Stunning visuals\n* Great soundtrack') // Example Markdown
      }
    },
  });

  fs.writeFileSync('report.docx', buffer);
}

generateReport().catch(console.error);