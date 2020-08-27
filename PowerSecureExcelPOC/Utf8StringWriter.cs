using System.IO;
using System.Text;

public sealed class Utf8StringWriter : StringWriter
    {
        private readonly Encoding stringWriterEncoding;
        public Utf8StringWriter(Encoding desiredEncoding)
            : base()
        {
            this.stringWriterEncoding = desiredEncoding;
        }

        public override Encoding Encoding
        {
            get
            {
                return this.stringWriterEncoding;
            }
        }
    }