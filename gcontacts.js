import axios from 'axios';
// const { google } = require('googleapis');
import { google } from 'googleapis';
const { OAuth2 } = google.auth;

// Configurer l'authentification Google OAuth2
const oauth2Client = new OAuth2(
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET,
  process.env.REDIRECT_URI
);

oauth2Client.setCredentials({ refresh_token: process.env.REFRESH_TOKEN });

export default async function handler(req, res) {
  try {
    // Récupérer les contacts depuis Airtable
    const airtableResponse = await axios.get(`https://api.airtable.com/v0/appPYTaiSofTygjwm/Personnes`, {
      headers: { Authorization: `Bearer ${process.env.AIRTABLE_API_KEY}` }
    });
    const airtableContacts = airtableResponse.data.records;

    const people = google.people({ version: 'v1', auth: oauth2Client });

    for (const contact of airtableContacts) {
      const { fields } = contact;
      
      // Construire l'objet contactData avec vérifications de la présence des champs
      const contactData = {
        names: fields.FirstName || fields.LastName ? [{ givenName: fields.FirstName, familyName: fields.LastName }] : [],
        emailAddresses: [],
        phoneNumbers: [],
        organizations: []
      };
      
      // Ajouter l'email principal, si présent
      if (fields.Email) {
        contactData.emailAddresses.push({ value: fields.Email, type: 'work' });
      }
      
      // Ajouter un autre email, si présent
      if (fields.Mail) {
        contactData.emailAddresses.push({ value: fields.Mail, type: 'other' });
      }
      
      // Ajouter le numéro de téléphone principal, si présent
      if (fields.Phone) {
        contactData.phoneNumbers.push({ value: fields.Phone, type: 'mobile' });
      }
      
      // Ajouter un autre numéro de téléphone, si présent
      if (fields.Téléphone) {
        contactData.phoneNumbers.push({ value: fields.Téléphone, type: 'work' });
      }
      
      // Ajouter les informations d'emploi si l'un des champs est présent
      if (fields.Entité || fields.Poste || fields.Établissement) {
        contactData.organizations.push({
          name: fields.Entité || '',        // Nom de l'entité (ex. société ou institution)
          title: fields.Poste || '',        // Poste de la personne
          department: fields.Établissement || '' // Nom de l'établissement
        });
      }
      
      // Créer ou mettre à jour le contact dans Google Contacts
      await people.people.createContact({
        requestBody: contactData
      });
    }

    res.status(200).json({ message: 'Synchronisation terminée avec succès !' });
  } catch (error) {
    console.error('Erreur de synchronisation :', error);
    res.status(500).json({ error: 'Erreur de synchronisation' });
  }
}
