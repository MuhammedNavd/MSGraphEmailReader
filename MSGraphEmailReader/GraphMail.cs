namespace MSGraphEmailReader
{
    [Serializable]
    public class GraphMail
    {
        public GraphMail() => Attachments = new List<Attachment>();
        public int MessageNumber { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime DateSent { get; set; }
        public List<Attachment> Attachments { get; set; }

        [Serializable]
        public class Attachment
        {
            public string FileName { get; set; }
            public string ContentType { get; set; }
            public byte[] Content { get; set; }
        }
    }
}
