import express from 'express';
import path from 'path';
import fs from 'fs';
import { getFrenchFormattedDate, fetchTemplate, generateReport, ensureDirectoryExists } from './utils.js';

const app = express();
app.use(express.json()); // For parsing application/json

app.post('/doc', async (req, res) => {
  let { url, data } = req.body;

  // Set default values if not provided
  if (!url) {
    url = 'https://ivv-fsh.github.io/templates/test.docx';
  }
  if (!data) {
    data = {
      Titre: 'John',
      Objectifs: 'Appleseed',
    };
  }

  try {
    const template = await fetchTemplate(url);
    const buffer = await generateReport(template, data);

    const originalFileName = path.basename(url);
    const fileNameWithoutExt = originalFileName.replace(path.extname(originalFileName), '');
    const newFileName = `${getFrenchFormattedDate()}-${fileNameWithoutExt}-R.docx`;

    const reportsDir = path.join(process.cwd(), 'reports');
    ensureDirectoryExists(reportsDir);

    const filePath = path.join(reportsDir, newFileName);
    fs.writeFileSync(filePath, buffer);

    console.log(`Report generated: ${filePath}`);

    // Send the file as a download
    res.download(filePath, newFileName, (err) => {
      if (err) {
        console.error('Error sending file:', err);
        res.status(500).send('Could not download the file.');
      }
    });

  } catch (error) {
    console.error('Error generating report:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});