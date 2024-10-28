import express from 'express';
import path from 'path';
import fs from 'fs';
import { getFrenchFormattedDate, fetchTemplate, generateReport, updateAirtableRecord, updateAirtableRecords, getAirtableSchema, processFieldsForDocx, getAirtableRecords, getAirtableRecord, ymd } from './utils.js';
import { Readable } from 'stream';
import archiver from 'archiver';
// import { Stream } from 'stream';
import { GITHUBTEMPLATES } from './constants.js';


const bufferToStream = (buffer) => {
    const stream = new Readable();
    stream.push(buffer);
    stream.push(null); // Signal the end of the stream
    return stream;
};


  