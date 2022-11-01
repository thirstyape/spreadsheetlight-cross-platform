using System;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for setting spreadsheet document properties.
    /// </summary>
    public class SLDocumentProperties
    {
        /// <summary>
        /// The category of the document.
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// The status of the content.
        /// </summary>
        public string ContentStatus { get; set; }

        internal string Created { get; set; }

        /// <summary>
        /// The creator of the document.
        /// </summary>
        public string Creator { get; set; }

        /// <summary>
        /// The summary or abstract of the contents of the document. This might also be the comment section.
        /// </summary>
        public string Description { get; set; }

        internal string Identifier { get; set; }

        /// <summary>
        /// A word or set of words describing the document.
        /// </summary>
        public string Keywords { get; set; }

        internal string Language { get; set; }

        /// <summary>
        /// The document is last modified by this person.
        /// </summary>
        public string LastModifiedBy { get; set; }

        internal string LastPrinted { get; set; }

        internal string Modified { get; set; }

        internal string Revision { get; set; }

        /// <summary>
        /// The topic of the contents of the document.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// The title of the document.
        /// </summary>
        public string Title { get; set; }

        internal string Version { get; set; }

        internal SLDocumentProperties()
        {
            this.SetAllNull();
        }

        internal void SetAllNull()
        {
            this.Category = string.Empty;
            this.ContentStatus = string.Empty;
            this.Created = string.Empty;
            this.Creator = string.Empty;
            this.Description = string.Empty;
            this.Identifier = string.Empty;
            this.Keywords = string.Empty;
            this.Language = string.Empty;
            this.LastModifiedBy = string.Empty;
            this.LastPrinted = string.Empty;
            this.Modified = string.Empty;
            this.Revision = string.Empty;
            this.Subject = string.Empty;
            this.Title = string.Empty;
            this.Version = string.Empty;
        }
    }
}
