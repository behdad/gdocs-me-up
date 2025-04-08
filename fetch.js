/**
 * print-doc-json.js
 *
 * Usage:
 *   node print-doc-json.js <DOC_ID>
 */

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json'; // or your credentials path

async function main() {
  const docId = process.argv[2];
  if (!docId) {
    console.error('Usage: node print-doc-json.js <DOC_ID>');
    process.exit(1);
  }

  // Auth using a service account key
  const auth = new google.auth.GoogleAuth({
    keyFile: SERVICE_ACCOUNT_KEY_FILE,
    scopes: [
      'https://www.googleapis.com/auth/documents.readonly',
      'https://www.googleapis.com/auth/drive.readonly'
    ]
  });
  const authClient = await auth.getClient();

  // Create the Docs client
  const docs = google.docs({ version: 'v1', auth: authClient });

  // Fetch the doc
  const { data } = await docs.documents.get({ documentId: docId });

  // 'data' is a JS object representing the doc. Let's print it as JSON.
  console.log(JSON.stringify(data, null, 2));
}

main().catch(err => {
  console.error('Error fetching doc JSON:', err);
  process.exit(1);
});

