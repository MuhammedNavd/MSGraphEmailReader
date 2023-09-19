﻿using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace MSGraphEmailReader
{
    public class ReadEmails
    {
        public async Task<List<GraphMail>> ReadEmailAsync(GraphEmailRequest graphEmailRequest)
        {
            GraphServiceClient graphServiceClient = await FetchGraphServiceClient(graphEmailRequest);
            List<QueryOption> queryOptions = new()
            {
                 new QueryOption("$filter", $"ReceivedDateTime ge {DateTimeOffset.Now.AddHours(-1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")}")
            };

#if DEBUG
            // Find the shared mailboxfolderId for every email address
            foreach (MailFolder mailboxFolder in await graphServiceClient.Users[graphEmailRequest.UserMailAddress].MailFolders.Request().GetAsync())
            {
                Console.WriteLine($"Folder name: {mailboxFolder.DisplayName}, Folder ID: {mailboxFolder.Id}");
            }
            MailFolder sharedMailboxFolder = await graphServiceClient.Users[graphEmailRequest.UserMailAddress]
                                           .MailFolders[graphEmailRequest.SharedMailBoxFolderId]
                                           .Request()
                                           .GetAsync();
#endif

            IMailFolderMessagesCollectionPage messages = await FetchMessage(graphEmailRequest, graphServiceClient, queryOptions);
            List<GraphMail> graphMails = await SetGraphMails(graphEmailRequest, graphServiceClient, messages);
            return graphMails;
        }

        private async Task<GraphServiceClient> FetchGraphServiceClient(GraphEmailRequest graphEmailRequest)
        {
            ClientCredential clientCredential = new(graphEmailRequest.ClientId, graphEmailRequest.ClientSecret);
            AuthenticationContext authContext = new($"https://login.microsoftonline.com/{graphEmailRequest.TenantId}");
            AuthenticationResult result = await authContext.AcquireTokenAsync("https://graph.microsoft.com", clientCredential);
            GraphServiceClient graphServiceClient = new(new DelegateAuthenticationProvider((requestMessage) =>
            {
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", result.AccessToken);
                return Task.FromResult(0);
            }));
            return graphServiceClient;
        }

        private async Task<IMailFolderMessagesCollectionPage> FetchMessage(GraphEmailRequest graphEmailRequest, GraphServiceClient graphServiceClient, List<QueryOption> queryOptions)
        {
            // Get message from IMailFolderMessagesCollectionPage based on applied filters
            return await graphServiceClient.Users[graphEmailRequest.UserMailAddress]
                                                               .MailFolders[graphEmailRequest.SharedMailBoxFolderId].Messages
                                                               .Request(queryOptions)
                                                               .Expand("attachments")
                                                               .GetAsync();
        }

        private async Task<List<GraphMail>> SetGraphMails(GraphEmailRequest graphEmailRequest, GraphServiceClient graphServiceClient, IMailFolderMessagesCollectionPage messages)
        {
            List<GraphMail> graphMails = new();
            foreach (Message message in messages.CurrentPage)
            {
                GraphMail graphMail = new();
                graphMail.From = message.From.EmailAddress.Address;
                graphMail.Subject = message.Subject;
                graphMail.Body = message.Body.Content;
                await AddAttachment(graphEmailRequest, graphServiceClient, graphMails, message, graphMail);
            }

            return graphMails;
        }

        private async Task AddAttachment(GraphEmailRequest graphEmailRequest, GraphServiceClient graphServiceClient, List<GraphMail> graphMails, Message message, GraphMail graphMail)
        {
            if (message.Attachments != null && message.Attachments.Count > 0)
            {
                foreach (Attachment attachment in message.Attachments)
                {
                    if (attachment is FileAttachment)
                    {
                        await SetAttachment(graphEmailRequest, graphServiceClient, message, graphMail, attachment);
                    }
                }
                graphMails.Add(graphMail);
            }
        }

        private async Task SetAttachment(GraphEmailRequest graphEmailRequest, GraphServiceClient graphServiceClient, Message message, GraphMail graphMail, Attachment attachment)
        {
            Attachment fileStream = await graphServiceClient.Users[graphEmailRequest.UserMailAddress]
                                                                .Messages[message.Id]
                                                                .Attachments[attachment.Id]
                                                                .Request().GetAsync();
            FileAttachment fileAttachment = fileStream as FileAttachment;
            graphMail.Attachments.Add(new GraphMail.Attachment
            {
                Content = fileAttachment.ContentBytes,
                ContentType = fileAttachment.ContentType,
                FileName = fileStream.Name,
            });
        }
    }
}