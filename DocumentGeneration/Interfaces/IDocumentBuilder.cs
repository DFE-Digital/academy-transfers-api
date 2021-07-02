using System;
using DocumentGeneration.Builders;

namespace DocumentGeneration.Interfaces
{
    public interface IDocumentBuilder : IDocumentBodyBuilder
    {
        public void ReplacePlaceholderWithContent(string placeholderText, Action<DocumentBodyBuilder> action);
        public void AddHeader(Action<IHeaderBuilder> action);
        public void AddFooter(Action<IFooterBuilder> action);
        public byte[] Build();
    }
}