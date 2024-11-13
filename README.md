# Azure Function - Send email via Microsoft Graph

This Azure function can be called to send an email using the Microsoft Graph API. It can be invoked via an HTTP request.

Example request body:

```
{
  "to": "recipient@example.com",
  "subject": "Test Email",
  "body": "This is a test email sent using Microsoft Graph API."
}
```