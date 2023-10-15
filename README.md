# MSGraphEmailReader

MSGraphEmailReader is a .NET class library that provides functionality for reading emails from Microsoft Graph. It allows you to fetch emails, including their attachments, from a specified mailbox folder in Microsoft 365 using Microsoft Graph API.

## Features

- Authenticate with Microsoft Graph using client credentials.
- Retrieve emails based on various filters, such as received date and time.
- Retrieve email attachments.
- Simple and easy-to-use API for integrating with Microsoft Graph.
- Also Retrieve Un-opened emails based on filter. 

## Prerequisites

Before using MSGraphEmailReader, make sure you have the following prerequisites in place:

- **Client Application Registration**: You need to register your application in the Azure Portal and obtain a `ClientId` and `ClientSecret` for authentication with Microsoft Graph.

- **Azure AD Tenant**: You must have access to an Azure AD tenant where you can register your application.

- **Access to Microsoft 365**: Ensure that your application has the necessary permissions to access Microsoft 365 mailbox data. The required permissions may vary depending on your use case.

- **User Mailbox and Shared Mailbox Information**: You should have the email addresses and folder IDs for the user mailbox and the shared mailbox folder from which you want to read emails.

- **While Using ReadUnopenedEmailsAsync**: Need to give some specific permissions `Mail.ReadWrite` or `Mail.Send` depending on your use case.


## Getting Started

To get started with MSGraphEmailReader, follow these steps:

1. Clone or download the repository.

2. Configure your application settings by providing the necessary values in the appsettings.json file, including `ClientId`, `ClientSecret`, `TenantId`, `UserMailAddress`, and `SharedMailBoxFolderId`.

3. Initialize the `GraphEmailRequest` object with your configuration.

4. Use the `ReadEmailAsync` method to retrieve emails from Microsoft Graph based on the specific Dateframe.

5. Use the `ReadUnopenedEmailsAsync` method to read all un-opened emails and their attachments without applying a date filter.


```csharp
// Example usage
GraphEmailRequest graphEmailRequest = new GraphEmailRequest
{
    ClientId = "YourClientId",
    ClientSecret = "YourClientSecret",
    TenantId = "YourTenantId",
    UserMailAddress = "UserEmailAddress",
    SharedMailBoxFolderId = "SharedMailboxFolderId",
};

ReadEmails reader = new ReadEmails();
List<GraphMail> emails = await reader.ReadEmailAsync(graphEmailRequest);
```
## Contributing
Contributions to MSGraphEmailReader are welcome! Please feel free to submit pull requests or raise issues if you have any feedback, suggestions, or bug reports.

## Acknowledgments
This project is inspired by the need to interact with Microsoft 365 mailbox data using the Microsoft Graph API. We would like to acknowledge the developers and contributors to the Microsoft Graph SDK for making this functionality accessible.