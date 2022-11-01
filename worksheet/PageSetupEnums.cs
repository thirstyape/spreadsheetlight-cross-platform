using System;

namespace SpreadsheetLight
{
    /// <summary>
    /// Paper size
    /// </summary>
    public enum SLPaperSizeValues
    {
        /// <summary>
        /// Letter paper (8.5 in by 11 in)
        /// </summary>
        LetterPaper = 1,
        /// <summary>
        /// Letter small paper (8.5 in by 11 in)
        /// </summary>
        LetterSmallPaper = 2,
        /// <summary>
        /// Tabloid paper (11 in by 17 in)
        /// </summary>
        TabloidPaper = 3,
        /// <summary>
        /// Ledger paper (17 in by 11 in)
        /// </summary>
        LedgerPaper = 4,
        /// <summary>
        /// Legal paper (8.5 in by 14 in)
        /// </summary>
        LegalPaper = 5,
        /// <summary>
        /// Statement paper (5.5 in by 8.5 in)
        /// </summary>
        StatementPaper = 6,
        /// <summary>
        /// Executive paper (7.25 in by 10.5 in)
        /// </summary>
        ExecutivePaper = 7,
        /// <summary>
        /// A3 paper (297 mm by 420 mm)
        /// </summary>
        A3Paper = 8,
        /// <summary>
        /// A4 paper (210 mm by 297 mm)
        /// </summary>
        A4Paper = 9,
        /// <summary>
        /// A4 small paper (210 mm by 297 mm)
        /// </summary>
        A4SmallPaper = 10,
        /// <summary>
        /// A5 paper (148 mm by 210 mm)
        /// </summary>
        A5Paper = 11,
        /// <summary>
        /// B4 paper (250 mm by 353 mm)
        /// </summary>
        B4Paper = 12,
        /// <summary>
        /// B5 paper (176 mm by 250 mm)
        /// </summary>
        B5Paper = 13,
        /// <summary>
        /// Folio paper (8.5 in by 13 in)
        /// </summary>
        FolioPaper = 14,
        /// <summary>
        /// Quarto paper (215 mm by 275 mm)
        /// </summary>
        QuartoPaper = 15,
        /// <summary>
        /// Standard paper (10 in by 14 in)
        /// </summary>
        StandardPaper10by14 = 16,
        /// <summary>
        /// Standard paper (11 in by 17 in)
        /// </summary>
        StandardPaper11by17 = 17,
        /// <summary>
        /// Note paper (8.5 in by 11 in)
        /// </summary>
        NotePaper = 18,
        /// <summary>
        /// #9 envelope (3.875 in by 8.875 in)
        /// </summary>
        Number9Envelope = 19,
        /// <summary>
        /// #10 envelope (4.125 in by 9.5 in)
        /// </summary>
        Number10Envelope = 20,
        /// <summary>
        /// #11 envelope (4.5 in by 10.375 in)
        /// </summary>
        Number11Envelope = 21,
        /// <summary>
        /// #12 envelope (4.75 in by 11 in)
        /// </summary>
        Number12Envelope = 22,
        /// <summary>
        /// #14 envelope (5 in by 11.5 in)
        /// </summary>
        Number14Envelope = 23,
        /// <summary>
        /// C paper (17 in by 22 in)
        /// </summary>
        CPaper = 24,
        /// <summary>
        /// D paper (22 in by 34 in)
        /// </summary>
        DPaper = 25,
        /// <summary>
        /// E paper (34 in by 44 in)
        /// </summary>
        EPaper = 26,
        /// <summary>
        /// DL envelope (110 mm by 220 mm)
        /// </summary>
        DLEnvelope = 27,
        /// <summary>
        /// C5 envelope (162 mm by 229 mm)
        /// </summary>
        C5Envelope = 28,
        /// <summary>
        /// C3 envelope (324 mm by 458 mm)
        /// </summary>
        C3Envelope = 29,
        /// <summary>
        /// C4 envelope (229 mm by 324 mm)
        /// </summary>
        C4Envelope = 30,
        /// <summary>
        /// C6 envelope (114 mm by 162 mm)
        /// </summary>
        C6Envelope = 31,
        /// <summary>
        /// C65 envelope (114 mm by 229 mm)
        /// </summary>
        C65Envelope = 32,
        /// <summary>
        /// B4 envelope (250 mm by 353 mm)
        /// </summary>
        B4Envelope = 33,
        /// <summary>
        /// B5 envelope (176 mm by 250 mm)
        /// </summary>
        B5Envelope = 34,
        /// <summary>
        /// B6 envelope (176 mm by 125 mm)
        /// </summary>
        B6Envelope = 35,
        /// <summary>
        /// Italy envelope (110 mm by 230 mm)
        /// </summary>
        ItalyEnvelope = 36,
        /// <summary>
        /// Monarch envelope (3.875 in by 7.5 in)
        /// </summary>
        MonarchEnvelope = 37,
        /// <summary>
        /// 6 3/4 envelope (3.625 in by 6.5 in)
        /// </summary>
        SixThreeQuarterEnvelope = 38,
        /// <summary>
        /// US standard fanfold (14.875 in by 11 in)
        /// </summary>
        USStandardFanfold = 39,
        /// <summary>
        /// German standard fanfold (8.5 in by 12 in)
        /// </summary>
        GermanStandardFanfold = 40,
        /// <summary>
        /// German legal fanfold (8.5 in by 13 in)
        /// </summary>
        GermanLegalFanfold = 41,
        /// <summary>
        /// ISO B4 (250 mm by 353 mm)
        /// </summary>
        IsoB4 = 42,
        /// <summary>
        /// Japanese double postcard (200 mm by 148 mm)
        /// </summary>
        JapaneseDoublePostcard = 43,
        /// <summary>
        /// Standard paper (9 in by 11 in)
        /// </summary>
        StandardPaper9by11 = 44,
        /// <summary>
        /// Standard paper (10 in by 11 in)
        /// </summary>
        StandardPaper10by11 = 45,
        /// <summary>
        /// Standard paper (15 in by 11 in)
        /// </summary>
        StandardPaper15by11 = 46,
        /// <summary>
        /// Invite envelope (220 mm by 220 mm)
        /// </summary>
        InviteEnvelope = 47,
        /// <summary>
        /// Letter extra paper (9.275 in by 12 in)
        /// </summary>
        LetterExtraPaper = 50,
        /// <summary>
        /// Legal extra paper (9.275 in by 15 in)
        /// </summary>
        LegalExtraPaper = 51,
        /// <summary>
        /// Tabloid extra paper (11.69 in by 18 in)
        /// </summary>
        TabloidExtraPaper = 52,
        /// <summary>
        /// A4 extra paper (236 mm by 322 mm)
        /// </summary>
        A4ExtraPaper = 53,
        /// <summary>
        /// Letter transverse paper (8.275 in by 11 in)
        /// </summary>
        LetterTransversePaper = 54,
        /// <summary>
        /// A4 transverse paper (210 mm by 297 mm)
        /// </summary>
        A4TransversePaper = 55,
        /// <summary>
        /// Letter extra transverse paper (9.275 in by 12 in)
        /// </summary>
        LetterExtraTransversePaper = 56,
        /// <summary>
        /// SuperA/SuperA/A4 paper (227 mm by 356 mm)
        /// </summary>
        SuperASuperAA4Paper = 57,
        /// <summary>
        /// SuperB/SuperB/A3 paper (305 mm by 487 mm)
        /// </summary>
        SuperBSuperBA3Paper = 58,
        /// <summary>
        /// Letter plus paper (8.5 in by 12.69 in)
        /// </summary>
        LetterPlusPaper = 59,
        /// <summary>
        /// A4 plus paper (210 mm by 330 mm)
        /// </summary>
        A4PlusPaper = 60,
        /// <summary>
        /// A5 transverse paper (148 mm by 210 mm)
        /// </summary>
        A5TransversePaper = 61,
        /// <summary>
        /// JIS B5 transverse paper (182 mm by 257 mm)
        /// </summary>
        JisB5TransversePaper = 62,
        /// <summary>
        /// A3 extra paper (322 mm by 445 mm)
        /// </summary>
        A3ExtraPaper = 63,
        /// <summary>
        /// A5 extra paper (174 mm by 235 mm)
        /// </summary>
        A5ExtraPaper = 64,
        /// <summary>
        /// ISO B5 extra paper (201 mm by 276 mm)
        /// </summary>
        IsoB5ExtraPaper = 65,
        /// <summary>
        /// A2 paper (420 mm by 594 mm)
        /// </summary>
        A2Paper = 66,
        /// <summary>
        /// A3 transverse paper (297 mm by 420 mm)
        /// </summary>
        A3TransversePaper = 67,
        /// <summary>
        /// A3 extra transverse paper (322 mm by 445 mm)
        /// </summary>
        A3ExtraTransversePaper = 68
    }

    /// <summary>
    /// The type of header or footer.
    /// </summary>
    public enum SLHeaderFooterTypeValues
    {
        /// <summary>
        /// First page.
        /// </summary>
        First = 0,
        /// <summary>
        /// Odd-numbered pages.
        /// </summary>
        Odd,
        /// <summary>
        /// Even-numbered pages.
        /// </summary>
        Even
    }

    internal enum SLHeaderFooterSectionValues
    {
        Left,
        Center,
        Right
    }

    /// <summary>
    /// Header and footer format codes.
    /// </summary>
    public enum SLHeaderFooterFormatCodeValues
    {
        /// <summary>
        /// This is a positional code. Internally, it's "&amp;L" (without quotes).
        /// </summary>
        Left = 0,
        /// <summary>
        /// This is a positional code. Internally, it's "&amp;C" (without quotes).
        /// </summary>
        Center,
        /// <summary>
        /// This is a positional code. Internally, it's "&amp;R" (without quotes).
        /// </summary>
        Right,
        /// <summary>
        /// Page number. Excel interface displays "&amp;[Page]" but internally it's "&amp;P" (without quotes).
        /// </summary>
        PageNumber,
        /// <summary>
        /// Number of pages. Excel interface displays "&amp;[Pages]" but internally it's "&amp;N" (without quotes).
        /// </summary>
        NumberOfPages,
        /// <summary>
        /// Current date. Excel interface displays "&amp;[Date]" but internally it's "&amp;D" (without quotes).
        /// </summary>
        Date,
        /// <summary>
        /// Current time. Excel interface displays "&amp;[Time]" but internally it's "&amp;T" (without quotes).
        /// </summary>
        Time,
        /// <summary>
        /// File path. Excel interface displays "&amp;[Path]" but internally it's "&amp;Z" (without quotes).
        /// </summary>
        FilePath,
        /// <summary>
        /// File name. Excel interface displays "&amp;[File]" but internally it's "&amp;F" (without quotes).
        /// </summary>
        FileName,
        /// <summary>
        /// Sheet name. Excel interface displays "&amp;[Tab]" but internally it's "&amp;A" (without quotes).
        /// </summary>
        SheetName,
        /// <summary>
        /// This resets the font styles. Use this when the previous section of text has formatting and the next section of text should have normal font styles.
        /// </summary>
        ResetFont
    }
}
