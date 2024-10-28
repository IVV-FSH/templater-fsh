import fs from 'fs';
import { put } from '@vercel/blob';

async function uploadDocxToVercelBlob() {
  try {
    // Path to your .docx file
    const filePath = './templates/test.docx';
    
    // Read the file buffer
    const fileBuffer = fs.readFileSync(filePath);
    console.log('File read successfully:', filePath);

    // Use the put function to upload the file
    const { url } = await put('documents/test.docx', fileBuffer, {
      access: 'public',  // Set access as needed: 'public' or 'private'
    });

    console.log('Upload successful:', url);
  } catch (error) {
    console.error('Error uploading .docx to Vercel Blob:', error);
  }
}

// Run the function to test
uploadDocxToVercelBlob();
