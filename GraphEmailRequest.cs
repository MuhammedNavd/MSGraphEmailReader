namespace MSGraphEmailReader
{
    public class GraphEmailRequest
    {
        public string? ClientId { get; set; }                   // YOUR_AZURE_CLIENT_ID
        public string? ClientSecret { get; set; }               // YOUR_REGISTERED_APP_SECRET
        public string? TenantId { get; set; }                   // YOUR_AZURE_TENANT_ID
        public string? UserMailAddress { get; set; }            // YOUR_EMAIL_ADDRESS
        public string? SharedMailBoxFolderId { get; set; }      // YOUR_SHARE_BOX_FOLDER_ID
        public DateTimeOffset RequestedDateTime { get; set; }   // RequestedDateTime
    }
}
