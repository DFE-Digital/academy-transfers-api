using System;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentGeneration.Interfaces;

namespace DocumentGeneration.Builders
{
    public class HeaderBuilder : IHeaderBuilder
    {
        private readonly Header _header;

        public HeaderBuilder(Header header)
        {
            _header = header;
        }

        public void AddParagraph(Action<IParagraphBuilder> action)
        {
            var paragraph = new Paragraph();
            var builder = new ParagraphBuilder(paragraph);
            action(builder);
            _header.AppendChild(paragraph);
        }
    }
}