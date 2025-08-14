const { Client } = require("@microsoft/microsoft-graph-client");
const { CertificateCredential } = require("@azure/identity");
const multer = require("multer");
const upload = multer();

module.exports = async function (context, req) {
  try {
    const form = await new Promise((resolve, reject) => {
      upload.single("attachment")(req, {}, err => {
        if (err) reject(err);
        else resolve(req);
      });
    });

    const { name, email, message } = form.body;
    const file = form.file;

    const credential = new CertificateCredential(
      "phinia.onmicrosoft.com",
      "a560568c-73e3-499c-abfe-3f1b1523775c",
      {
        certificatePath: "C:\Users\mrone\Downloads\Corporate Web Apps Automation.pfx",
    
      }
    );

    const graphClient = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          const token = await credential.getToken("https://graph.microsoft.com/.default");
          return token.token;
        }
      }
    });

    if (file) {
      await graphClient
        .api(`/sites/YOUR_SITE_ID/drives/YOUR_DRIVE_ID/root:/Attachments/${file.originalname}:/content`)
        .put(file.buffer);
    }

    await graphClient
      .api(`/sites/YOUR_SITE_ID/lists/YOUR_LIST_ID/items`)
      .post({
        fields: {
          Title: name,
          Email: email,
          Message: message
        }
      });

    context.res = {
      status: 200,
      body: "Submission successful"
    };
  } catch (error) {
    context.res = {
      status: 500,
      body: "Error: " + error.message
    };
  }
};
