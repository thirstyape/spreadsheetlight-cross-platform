using System;

namespace SpreadsheetLight
{
    /// <summary>
    /// The type of hyperlink.
    /// </summary>
    public enum SLHyperlinkTypeValues
    {
        /// <summary>
        /// Hyperlink to an existing web page.
        /// </summary>
        Url = 0,
        /// <summary>
        /// Hyperlink to an existing file.
        /// </summary>
        FilePath,
        /// <summary>
        /// Hyperlink to a place within the spreadsheet (cell references or defined names).
        /// </summary>
        InternalDocumentLink,
        /// <summary>
        /// Hyperlink to an email address.
        /// </summary>
        EmailAddress
    }

    /// <summary>
    /// The type of measurement unit.
    /// </summary>
    public enum SLMeasureUnitTypeValues
    {
        /// <summary>
        /// English Metric Unit. No, not the bird...
        /// </summary>
        Emu = 0,
        /// <summary>
        /// Inch.
        /// </summary>
        Inch,
        /// <summary>
        /// Centimeter.
        /// </summary>
        Centimeter,
        /// <summary>
        /// Point.
        /// </summary>
        Point
    }

    /// <summary>
    /// The type of paste options.
    /// </summary>
    public enum SLPasteTypeValues
    {
        /// <summary>
        /// Just plain pasting. Fanfare and choral singing each sold separately. *smile*
        /// </summary>
        Paste = 0,
        /// <summary>
        /// Paste only values (no formulas).
        /// </summary>
        Values,
        /// <summary>
        /// Paste values and formulas. NOTE: Formulas are copied as is (no recalculating cell references).
        /// </summary>
        Formulas,
        /// <summary>
        /// Transpose.
        /// </summary>
        Transpose,
        /// <summary>
        /// Paste only formatting (styles).
        /// </summary>
        Formatting
    }
}
