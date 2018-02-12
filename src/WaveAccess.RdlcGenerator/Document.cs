namespace WaveAccess.RdlcGenerator {
    public class Document {
        public Document(byte[] content, string mimeType, string encoding, string extension, string[] streamids, int pageCount) { 
            this.Content = content;
            this.MimeType = mimeType;
            this.Encoding = encoding;
            this.Extension = extension;
            this.StreamIds = streamids;
            this.PageCount = pageCount;
        }
        public byte[] Content { get; private set; }
        public string MimeType { get; private set; }
        public string Encoding { get; private set; }
        public string Extension { get; private set; }
        public string[] StreamIds { get; private set; }
        public int PageCount { get; private set; }
    }
}
