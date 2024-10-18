import fs from 'fs';
import path from 'path';
import { createReport } from 'docx-templates';
import fetch from 'node-fetch';

export function getFrenchFormattedDate() {
    const options = {
        year: '2-digit',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        timeZone: 'Europe/Paris',
        hour12: false,
    };
    
    const formatter = new Intl.DateTimeFormat('fr-FR', options);
    const parts = formatter.formatToParts(new Date());
    
    const dateParts = {};
    parts.forEach(({ type, value }) => {
        dateParts[type] = value;
    });
    
    return `${dateParts.year}${dateParts.month}${dateParts.day}-${dateParts.hour}${dateParts.minute}`;
}

export async function fetchTemplate(url) {
    const response = await fetch(url);
    const arrayBuffer = await response.arrayBuffer();
    return Buffer.from(arrayBuffer);
  }
  
  export async function generateReport(template, data) {
    return await createReport({
      template,
    //   cmdDelimiter: ['{{', '}}'],
      data,
    });
  }
  
  export function ensureDirectoryExists(dir) {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir);
    }
  }