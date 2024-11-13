export const config = {
    clientId: process.env.CLIENT_ID || "",  
    tenantId: process.env.TENANT_ID || "",  
    clientSecret: process.env.CLIENT_SECRET || "",  
    authority: "https://login.microsoftonline.com",
    scope: ["Mail.Send"]
};
