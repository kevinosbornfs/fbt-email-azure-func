import * as msal from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { config } from "../config";

const cca = new msal.ConfidentialClientApplication({
    auth: {
        clientId: config.clientId,
        authority: `${config.authority}/${config.tenantId}`,
        clientSecret: config.clientSecret
    }
});

async function getAccessToken() {
    const result = await cca.acquireTokenByClientCredential({
        scopes: config.scope,
    });
    return result.accessToken;
}

function getGraphClient(accessToken: string) {
    return Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}

const sendEmail = async function (context: any, req: any): Promise<void> {  // Explicit signature isn't needed for now
    const { to, subject, body } = req.body;

    if (!to || !subject || !body) {
        context.res = {
            status: 400,
            body: "Please provide 'to', 'subject', and 'body' in the request body."
        };
        return;
    }

    try {
        const accessToken = await getAccessToken();
        const graphClient = getGraphClient(accessToken);

        const message = {
            subject: subject,
            body: {
                contentType: "Text",
                content: body
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: to
                    }
                }
            ]
        };

        await graphClient.api("/me/sendMail").post({ message });

        context.res = {
            status: 200,
            body: `Email sent successfully to ${to}!`
        };
    } catch (error) {
        context.res = {
            status: 500,
            body: `Error sending email: ${error.message}`
        };
    }
};

export default sendEmail;
