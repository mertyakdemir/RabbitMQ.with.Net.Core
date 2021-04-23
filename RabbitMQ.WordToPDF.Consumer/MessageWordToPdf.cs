using System;
using System.Collections.Generic;
using System.Text;

namespace RabbitMQ.WordToPDF.Consumer
{
    internal class MessageWordToPdf
    {
        public byte[] WordByte { get; set; }
        public string Email { get; set; }
        public string FileName { get; set; }
    }
}
