using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml;
using Excel = DocumentFormat.OpenXml.Office.Excel;

// Targetting .NET Standard means Forms is no longer useful
// Keeping this as a comment as a reminder of pre-Standard days
// (or until this comment is no longer useful :)
//using System.Windows.Forms;

namespace SpreadsheetLight
{
    // Why is SLConvert similar to SLTool? Why is this made internal and not public?
    // Because I want SLTool to be consistent with say SLDrawingTool and that they're
    // really internal to SpreadsheetLight. SLConvert is meant as the "real" public interface.
    // This way, the end developer only need to remember SLConvert. And also SLConvert is
    // similar to using System.Convert class. Most .NET developers should be familiar with
    // that class.

    // Also, good design is as much of hiding things for clarity as much as exposing needed
    // features. "Need" depends on the individual developer I guess, but the target end
    // developer shouldn't see *everything*. Does the typical Windows user access the kernel?
    // Does the typical iPhone user access the phone internals?

    // Also, typing "SLTool" is faster than typing "SLConvert"...

    // Also, typing "SLTool" avoids confusion with typing "SLCon" and getting "SLConvert"
    // and "SLConstants" as valid Intellisense options. I try to help Intellisense filter out
    // options whenever I can. :)

    internal partial class SLTool
    {
        internal static bool CheckSheetChartName(string Name)
        {
            bool result = true;
            if (Name.Contains('\\') || Name.Contains('/') || Name.Contains('?') || Name.Contains('*') || Name.Contains('[') || Name.Contains(']'))
            {
                result = false;
            }
            if (Name.Length > 31 || Name.Length == 0)
            {
                result = false;
            }

            return result;
        }

        internal static bool CheckRowColumnIndexLimit(int RowIndex, int ColumnIndex)
        {
            bool result = true;
            if (RowIndex > SLConstants.RowLimit || RowIndex < 1)
            {
                result = false;
            }
            if (ColumnIndex > SLConstants.ColumnLimit || ColumnIndex < 1)
            {
                result = false;
            }
            return result;
        }

        internal static bool IsCellReference(string Input)
        {
            Input = Input.Trim();
            bool result = false;
            if (Regex.IsMatch(Input, @"^\$?[a-zA-Z]{1,3}\$?\d{1,7}$"))
            {
                int iRowIndex = -1, iColumnIndex = -1;
                Input = Input.Replace("$", "");
                if (SLTool.FormatCellReferenceToRowColumnIndex(Input, out iRowIndex, out iColumnIndex))
                {
                    result = true;
                }
            }

            return result;
        }

        internal static string ToColumnName(int ColumnIndex)
        {
            if (ColumnIndex < 1 || ColumnIndex > SLConstants.ColumnLimit) return SLConstants.ErrorReference;

            string result = SLConstants.ErrorReference;

            // As of this writing, there's a maximum of 3 letters, with the "largest" as "XFD".
            // We're using this information as an optimisation. If Excel increases the column size,
            // just add more. But why would anyone need more than 16384 columns?!?!
            // If you're using Excel to do multi-variable scenario planning so you need thousands of
            // columns as variable cells, using the entire spreadsheet like a giant matrix, you're probably
            // doing it wrong. Write a proper program to do that, or use a maths program to solve it.
            // I know Excel has the Solver add-in, but still... solving large systems of linear equations
            // is asking a little much from Excel, huh?
            int iLetterPos3, iLetterPos2, iLetterPos1;

            // The typical method seems to be incrementing numbers in a loop and getting the column
            // name that way. The downside is that all the additions, subtractions, divisions and
            // modulus functions are performed in every loop iteration. A typical spreadsheet has
            // maybe 10 to 20 columns. That means every time you need a column name, you're looping
            // anywhere between 10 to 20 iterations. Every single time.
            // That's a waste of CPU cycles. The worst case scenario is if you set cell values only
            // on column XFD and 1 million+ rows. You're accruing 16384 million iterations.
            
            // As a historical note, SpreadsheetLight originally used a static List<string> and Dictionary
            // to hold the column names and the reverse column name index. This is precalculated
            // at initialisation. This follows a game development tip: Precalculate everything when
            // possible. After setting >16384 cell values, the cost of precalculation would have paid
            // for itself.
            // However, the 2 structures need to be static so other classes can access them.
            // This has problems when multi-threading or multi-whatever comes in.

            // Did you know I tried writing a switch statement? Visual Studio Express choked when
            // I dumped 16384 case statements into it. Maybe Visual Studio Professional will be fine...
            // Yes, all this is written with the Express version.

            // So I came up with the following algorithm. Every column name calculation will thus
            // have a fixed number of operations. Which is as close to O(n) as when accessing with the
            // static List<string> structure as I could get.

            // Consider the column name "DYZ". Using straightforward division and modulus by 676 or 26,
            // we get (5, 0, 0). The correct representation is (4, 25, 26).
            // The problem is 676*5 + 26*0 + 0 = 676*4 + 26*25 + 26 = 3380
            // The solution is that you can only have leading zeroes, not trailing zeroes.
            // (5, 0, 0) has trailing zeroes, so it's the wrong interpretation.
            // "Trailing" zeroes include (5, 0, 1). That's invalid.
            // Basically, any zero after a non-zero entry is invalid.

            // Mathematically speaking, the range of values is [1, 26] rather than the [0, 25].
            // This is why it appears that we have multiple ways of representing the same value.
            // The value 0 is wrapped around to 26 with other values intact.

            // +65 to get "A" in ASCII value. Because our values are in the range [1, 26], we
            // don't +65 but +64.

            if (ColumnIndex <= 26)
            {
                //26 as "Z", the last single character column name
                iLetterPos3 = -1;
                iLetterPos2 = -1;
                iLetterPos1 = ColumnIndex % 26;
                if (iLetterPos1 == 0) iLetterPos1 = 26;

                result = string.Format("{0}", (char)(iLetterPos1 + 64));
            }
            else if (ColumnIndex <= 702)
            {
                // 702 = 676 + 26 as "ZZ", the last double character column name
                iLetterPos3 = -1;
                iLetterPos2 = ColumnIndex / 26;
                iLetterPos1 = ColumnIndex % 26;
                if (iLetterPos1 == 0)
                {
                    iLetterPos1 = 26;
                    --iLetterPos2;
                }

                result = string.Format("{0}{1}", (char)(iLetterPos2 + 64), (char)(iLetterPos1 + 64));
            }
            else
            {
                // note that at this point, ColumnIndex >676, meaning
                // iLetterPos3 is at least 1. So we can do borrowing to the other 2.
                iLetterPos3 = ColumnIndex / 676;
                iLetterPos2 = (ColumnIndex / 26) % 26;
                iLetterPos1 = ColumnIndex % 26;
                if (iLetterPos1 == 0)
                {
                    iLetterPos1 = 26;
                    --iLetterPos2;
                }
                if (iLetterPos2 <= 0)
                {
                    iLetterPos2 += 26;
                    --iLetterPos3;
                }
                if (iLetterPos3 <= 0) iLetterPos3 += 26;

                result = string.Format("{0}{1}{2}", (char)(iLetterPos3 + 64), (char)(iLetterPos2 + 64), (char)(iLetterPos1 + 64));
            }

            return result;
        }

        internal static int ToColumnIndex(string Input)
        {
            int iRowIndex = -1, iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(Input, out iRowIndex, out iColumnIndex))
            {
                // don't have to do anything because iColumnIndex is already correctly assigned.
            }
            else
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(Input, "^[a-zA-Z]{1,3}$"))
                {
                    Input = Input.ToUpperInvariant();
                    // I see no point in writing a loop when there's only 3 cases...
                    // Each character is already in upper case. The hack is "A" is 65. We subtract 64
                    // to get a value between 1 and 26 (both inclusive).
                    // So "A" to "Z" is 65 to 90, which becomes 1 to 26 after subtraction.
                    // And 676 = 26 * 26
                    if (Input.Length == 3)
                    {
                        // Simplifying calculations to reduce number of operations
                        //iColumnIndex = ((int)Input[0] - 64) * 676 + ((int)Input[1] - 64) * 26 + ((int)Input[2] - 64);
                        iColumnIndex = ((int)Input[0]) * 676 + ((int)Input[1]) * 26 + ((int)Input[2]) - 44992;
                    }
                    else if (Input.Length == 2)
                    {
                        // Simplifying calculations to reduce number of operations
                        //iColumnIndex = ((int)Input[0] - 64) * 26 + ((int)Input[1] - 64);
                        iColumnIndex = ((int)Input[0]) * 26 + ((int)Input[1]) - 1728;
                    }
                    else
                    {
                        // it can only be length 1 here because the regular expression already ensured it.
                        iColumnIndex = ((int)Input[0] - 64);
                    }
                }
            }

            return iColumnIndex;
        }

        internal static string FormatWorksheetNameForFormula(string WorksheetName)
        {
            //http://support.microsoft.com/kb/107468
            // If it originally has single quotes, make it 2 single quotes.
            // If include workbook path, or if start with digit, or has special characters,
            // then must surround with single quotes.
            // Based on empirical experiments, the underscore and period character don't
            // require the single quotes.
            string result = WorksheetName.Replace("'", "''");
            if ((result.IndexOfAny(" `~!@#$%^&*()-=+[{]}\\|;:\",<>/?".ToCharArray()) > -1)
                || Regex.IsMatch(result, "^\\d"))
            {
                result = string.Format("'{0}'", result);
            }

            return result;
        }

        internal static string ToCellReference(int RowIndex, int ColumnIndex)
        {
            return ToCellReference(string.Empty, RowIndex, ColumnIndex, false);
        }

        internal static string ToCellReference(int RowIndex, int ColumnIndex, bool IsAbsolute)
        {
            return ToCellReference(string.Empty, RowIndex, ColumnIndex, IsAbsolute);
        }

        internal static string ToCellReference(string WorksheetName, int RowIndex, int ColumnIndex)
        {
            return ToCellReference(WorksheetName, RowIndex, ColumnIndex, false);
        }

        internal static string ToCellReference(string WorksheetName, int RowIndex, int ColumnIndex, bool IsAbsolute)
        {
            int iRowIndex = RowIndex;
            int iColumnIndex = ColumnIndex;

            if (iRowIndex < 1) iRowIndex = 1;
            if (iRowIndex > SLConstants.RowLimit) iRowIndex = SLConstants.RowLimit;
            if (iColumnIndex < 1) iColumnIndex = 1;
            if (iColumnIndex > SLConstants.ColumnLimit) iColumnIndex = SLConstants.ColumnLimit;

            string sWorksheetName = WorksheetName;
            if (sWorksheetName.Length > 0)
            {
                sWorksheetName = string.Format("{0}!", SLTool.FormatWorksheetNameForFormula(sWorksheetName));
            }
            else sWorksheetName = string.Empty;

            string result = string.Empty;
            if (IsAbsolute) result = string.Format("{0}${1}${2}", sWorksheetName, SLTool.ToColumnName(iColumnIndex), iRowIndex.ToString(CultureInfo.InvariantCulture));
            else result = string.Format("{0}{1}{2}", sWorksheetName, SLTool.ToColumnName(iColumnIndex), iRowIndex.ToString(CultureInfo.InvariantCulture));

            return result;
        }

        internal static string ToCellRange(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return ToCellRange(string.Empty, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, false);
        }

        internal static string ToCellRange(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool IsAbsolute)
        {
            return ToCellRange(string.Empty, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, IsAbsolute);
        }

        internal static string ToCellRange(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return ToCellRange(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, false);
        }

        internal static string ToCellRange(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool IsAbsolute)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            string sWorksheetName = WorksheetName;
            if (sWorksheetName.Length > 0)
            {
                sWorksheetName = string.Format("{0}!", SLTool.FormatWorksheetNameForFormula(sWorksheetName));
            }
            else sWorksheetName = string.Empty;

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iStartRowIndex > SLConstants.RowLimit) iStartRowIndex = SLConstants.RowLimit;
            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iStartColumnIndex > SLConstants.ColumnLimit) iStartColumnIndex = SLConstants.ColumnLimit;
            if (iEndRowIndex < 1) iEndRowIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
            if (iEndColumnIndex < 1) iEndColumnIndex = 1;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            string result = string.Empty;
            if (IsAbsolute) result = string.Format("{0}${1}${2}:${3}${4}", sWorksheetName, SLTool.ToColumnName(iStartColumnIndex), iStartRowIndex.ToString(CultureInfo.InvariantCulture), SLTool.ToColumnName(iEndColumnIndex), iEndRowIndex.ToString(CultureInfo.InvariantCulture));
            else result = string.Format("{0}{1}{2}:{3}{4}", sWorksheetName, SLTool.ToColumnName(iStartColumnIndex), iStartRowIndex.ToString(CultureInfo.InvariantCulture), SLTool.ToColumnName(iEndColumnIndex), iEndRowIndex.ToString(CultureInfo.InvariantCulture));

            return result;
        }

        internal static bool FormatCellReferenceRangeToRowColumnIndex(string CellReferenceRange, out int StartRowIndex, out int StartColumnIndex, out int EndRowIndex, out int EndColumnIndex)
        {
            bool result = true;
            StartRowIndex = -1;
            StartColumnIndex = -1;
            EndRowIndex = -1;
            EndColumnIndex = -1;
            int index = -1;
            string sRef1 = string.Empty, sRef2 = string.Empty;
            if (Regex.IsMatch(CellReferenceRange, "^[a-zA-Z]{1,3}\\d{1,7}:[a-zA-Z]{1,3}\\d{1,7}$"))
            {
                index = CellReferenceRange.IndexOf(":");
                sRef1 = CellReferenceRange.Substring(0, index);
                sRef2 = CellReferenceRange.Substring(index + 1);

                result = SLTool.FormatCellReferenceToRowColumnIndex(sRef1, out StartRowIndex, out StartColumnIndex);
                result &= SLTool.FormatCellReferenceToRowColumnIndex(sRef2, out EndRowIndex, out EndColumnIndex);
            }
            else
            {
                result = false;
            }

            return result;
        }

        internal static bool FormatCellReferenceToRowColumnIndex(string CellReference, out int RowIndex, out int ColumnIndex)
        {
            bool result = true;
            RowIndex = -1;
            ColumnIndex = -1;
            int index = -1;
            if (Regex.IsMatch(CellReference, @"^[a-zA-Z]{1,3}\d{1,7}$"))
            {
                index = CellReference.IndexOfAny("0123456789".ToCharArray());
                if (!int.TryParse(CellReference.Substring(index), out RowIndex))
                {
                    result = false;
                }
                else
                {
                    string sPrefix = CellReference.Substring(0, index);
                    ColumnIndex = SLTool.ToColumnIndex(sPrefix);
                }
            }
            else
            {
                result = false;
            }

            if (!SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                result = false;
            }

            return result;
        }

        internal static System.Drawing.Color ToColor(string HexValue)
        {
            NumberStyles ns = NumberStyles.HexNumber;
            CultureInfo ci = CultureInfo.InvariantCulture;
            string sColor = HexValue;
            int iLength = 0;
            int iAlpha = 0, iRed = 0, iGreen = 0, iBlue = 0;
            if (sColor.Length < 8) sColor = sColor.PadLeft(8, 'F');
            iLength = sColor.Length;

            if (!int.TryParse(sColor.Substring(iLength - 8, 2), ns, ci, out iAlpha))
            {
                iAlpha = 255;
            }
            if (!int.TryParse(sColor.Substring(iLength - 6, 2), ns, ci, out iRed))
            {
                iRed = 0;
            }
            if (!int.TryParse(sColor.Substring(iLength - 4, 2), ns, ci, out iGreen))
            {
                iGreen = 0;
            }
            if (!int.TryParse(sColor.Substring(iLength - 2, 2), ns, ci, out iBlue))
            {
                iBlue = 0;
            }

            return System.Drawing.Color.FromArgb(iAlpha, iRed, iGreen, iBlue);
        }

        internal static System.Drawing.Color ToColor(double Hue, double Saturation, double Luminance)
        {
            double fChroma = (1.0 - Math.Abs(2.0 * Luminance - 1.0)) * Saturation;
            double fHue = Hue / 60.0;
            double fHueMod2 = fHue;
            while (fHueMod2 >= 2.0) fHueMod2 -= 2.0;
            double fTemp = fChroma * (1.0 - Math.Abs(fHueMod2 - 1.0));

            double fRed = 0, fGreen = 0, fBlue = 0;
            if (fHue < 1.0)
            {
                fRed = fChroma;
                fGreen = fTemp;
                fBlue = 0;
            }
            else if (fHue < 2.0)
            {
                fRed = fTemp;
                fGreen = fChroma;
                fBlue = 0;
            }
            else if (fHue < 3.0)
            {
                fRed = 0;
                fGreen = fChroma;
                fBlue = fTemp;
            }
            else if (fHue < 4.0)
            {
                fRed = 0;
                fGreen = fTemp;
                fBlue = fChroma;
            }
            else if (fHue < 5.0)
            {
                fRed = fTemp;
                fGreen = 0;
                fBlue = fChroma;
            }
            else if (fHue < 6.0)
            {
                fRed = fChroma;
                fGreen = 0;
                fBlue = fTemp;
            }
            else
            {
                fRed = 0;
                fGreen = 0;
                fBlue = 0;
            }

            double fMin = Luminance - 0.5 * fChroma;
            fRed += fMin;
            fGreen += fMin;
            fBlue += fMin;

            fRed *= 255.0;
            fGreen *= 255.0;
            fBlue *= 255.0;

            int iRed = 0, iGreen = 0, iBlue = 0;
            // the default seems to be to truncate rather than round
            iRed = Convert.ToInt32(Math.Truncate(fRed));
            iGreen = Convert.ToInt32(Math.Truncate(fGreen));
            iBlue = Convert.ToInt32(Math.Truncate(fBlue));
            if (iRed < 0) iRed = 0;
            if (iRed > 255) iRed = 255;
            if (iGreen < 0) iGreen = 0;
            if (iGreen > 255) iGreen = 255;
            if (iBlue < 0) iBlue = 0;
            if (iBlue > 255) iBlue = 255;

            return System.Drawing.Color.FromArgb(iRed, iGreen, iBlue);
        }

        internal static System.Drawing.Color ToColor(System.Drawing.Color clr, double Tint)
        {
            if (Tint < -1.0) Tint = -1.0;
            if (Tint > 1.0) Tint = 1.0;

            System.Drawing.Color clrRgb = clr;
            double fHue = clrRgb.GetHue();
            double fSat = clrRgb.GetSaturation();
            double fLum = clrRgb.GetBrightness();
            if (Tint < 0)
            {
                fLum = fLum * (1.0 + Tint);
            }
            else
            {
                //fLum = fLum * (1.0 - fTint) + (1.0 - 1.0 * (1.0 - fTint));
                // simplified to this
                fLum = fLum * (1.0 - Tint) + Tint;
            }
            clrRgb = SLTool.ToColor(fHue, fSat, fLum);

            return clrRgb;
        }

        internal static string RemoveNamespaceDeclaration(string Text)
        {
            string s = Text.Replace(string.Format(" xmlns:x=\"{0}\"", SLConstants.NamespaceX), "");
            s = s.Replace(string.Format(" xmlns:a=\"{0}\"", SLConstants.NamespaceA), "");

            return s;
        }

        internal static bool ToPreserveSpace(string Text)
        {
            bool result = false;
            if (Text != null)
            {
                if (Regex.IsMatch(Text, "^\\s") || Regex.IsMatch(Text, "\\s$") || Regex.IsMatch(Text, "^\\s+$"))
                {
                    // there's a space at the start
                    // or there's a space at the end
                    // or there's nothing but space
                    result = true;
                }
            }
            return result;
        }

        internal static System.Drawing.Imaging.ImageFormat TranslateImageContentType(string ContentType)
        {
            System.Drawing.Imaging.ImageFormat imgtype = System.Drawing.Imaging.ImageFormat.Jpeg;
            switch (ContentType.ToLowerInvariant())
            {
                case "image/bmp":
                    imgtype = System.Drawing.Imaging.ImageFormat.Bmp;
                    break;
                case "image/gif":
                    imgtype = System.Drawing.Imaging.ImageFormat.Gif;
                    break;
                case "image/png":
                    imgtype = System.Drawing.Imaging.ImageFormat.Png;
                    break;
                case "image/tiff":
                    imgtype = System.Drawing.Imaging.ImageFormat.Tiff;
                    break;
                case "image/x-icon":
                    imgtype = System.Drawing.Imaging.ImageFormat.Icon;
                    break;
                case "image/x-pcx":
                    // not supported by .NET! Just use bmp and pray...
                    imgtype = System.Drawing.Imaging.ImageFormat.Bmp;
                    break;
                case "image/jpeg":
                    imgtype = System.Drawing.Imaging.ImageFormat.Jpeg;
                    break;
                case "image/x-emf":
                    imgtype = System.Drawing.Imaging.ImageFormat.Emf;
                    break;
                case "image/x-wmf":
                    imgtype = System.Drawing.Imaging.ImageFormat.Wmf;
                    break;
            }

            return imgtype;
        }

        internal static double CalculateDaysFromEpoch(DateTime Data, bool For1904Epoch)
        {
            DateTime dtEpoch;
            // Microsoft Excel for Windows uses 1 Jan 1900 as the epoch
            // Microsoft Excel for Macintosh uses 1 Jan 1904 as the epoch
            if (For1904Epoch) dtEpoch = SLConstants.Epoch1904();
            else dtEpoch = SLConstants.Epoch1900();

            TimeSpan ts = Data - dtEpoch;
            double fDateTime = 0;
            if (For1904Epoch)
            {
                // Excel doesn't add an addition day...
                fDateTime = ts.TotalDays;
            }
            else
            {
                // for backwards compatibility, 29 Feb 1900 is a valid date,
                // even though 1900 is not a leap year.
                // 29 Feb 1900 is 59 days after 1 Jan 1900, so we just skip to 1 Mar 1900
                if (ts.Days >= 59)
                {
                    fDateTime = ts.TotalDays + 2.0;
                }
                else
                {
                    fDateTime = ts.TotalDays + 1.0;
                }
            }
            return fDateTime;
        }

        internal static DateTime CalculateDateTimeFromDaysFromEpoch(double Days, bool For1904Epoch)
        {
            DateTime dtEpoch;
            // Microsoft Excel for Windows uses 1 Jan 1900 as the epoch
            // Microsoft Excel for Macintosh uses 1 Jan 1904 as the epoch
            if (For1904Epoch) dtEpoch = SLConstants.Epoch1904();
            else dtEpoch = SLConstants.Epoch1900();

            DateTime dt;

            if (For1904Epoch)
            {
                dt = dtEpoch.AddDays(Days);
            }
            else
            {
                if (Days < 59)
                {
                    dt = dtEpoch.AddDays(Days - 1.0);
                }
                else
                {
                    dt = dtEpoch.AddDays(Days - 2.0);
                }
            }

            return dt;
        }

        internal static string XmlWrite(string XmlToBeEscaped, bool ThrowExceptionsIfAny)
        {
            string result = XmlToBeEscaped;

            try
            {
                StringWriter sw = new StringWriter();
                XmlTextWriter xtw = new XmlTextWriter(sw);

                xtw.WriteString(XmlToBeEscaped);
                result = sw.ToString();
                // apparently the escaped double quote escapes (haha) the XML escaping...
                result = result.Replace("\"", "&quot;"); 

                xtw.Close();
                sw.Close();
            }
            catch (Exception e)
            {
                if (ThrowExceptionsIfAny)
                {
                    throw e;
                }
                else
                {
                    // we don't care what really went wrong. The priority is to not throw errors...
                    result = XmlToBeEscaped;
                }
            }

            return result;
        }

        internal static string XmlRead(string XmlThatsEscaped, bool ThrowExceptionsIfAny)
        {
            string result = XmlThatsEscaped;

            try
            {
                // doesn't matter what the tag is. We just need a root XML tag.
                StringReader sr = new StringReader(string.Format("<sl>{0}</sl>", XmlThatsEscaped));
                XmlTextReader xtr = new XmlTextReader(sr);
                
                xtr.Read();
                result = xtr.ReadString();

                xtr.Close();
                sr.Close();
            }
            catch (Exception e)
            {
                if (ThrowExceptionsIfAny)
                {
                    throw e;
                }
                else
                {
                    // we don't care what really went wrong. The priority is to not throw errors...
                    result = XmlThatsEscaped;
                }
            }

            return result;
        }

        internal static Font GetUsableNormalFont(string FontName, double FontSize, FontStyle DrawStyle, bool ThrowExceptionsIfAny)
        {
            Font usablefont = new Font(FontFamily.GenericSansSerif, (float)FontSize);

            // there's this elaborate dance of try-catch-if-else because apparently certain typefaces
            // don't have a "normal" version but say the bold version works just fine.
            // For example, the typeface Aharoni will choke with FontStyle.Regular but FontStyle.Bold is fine.

            try
            {
                usablefont = new Font(FontName, (float)FontSize, DrawStyle);
            }
            catch (Exception e)
            {
                if (ThrowExceptionsIfAny)
                {
                    throw e;
                }
                else
                {
                    FontFamily ff = new FontFamily(FontName);
                    FontStyle fsLastDitch = FontStyle.Regular;
                    if ((DrawStyle & FontStyle.Bold) > 0)
                    {
                        if (ff.IsStyleAvailable(FontStyle.Bold)) fsLastDitch |= FontStyle.Bold;
                    }
                    if ((DrawStyle & FontStyle.Italic) > 0)
                    {
                        if (ff.IsStyleAvailable(FontStyle.Italic)) fsLastDitch |= FontStyle.Italic;
                    }
                    if ((DrawStyle & FontStyle.Strikeout) > 0)
                    {
                        if (ff.IsStyleAvailable(FontStyle.Strikeout)) fsLastDitch |= FontStyle.Strikeout;
                    }
                    if ((DrawStyle & FontStyle.Underline) > 0)
                    {
                        if (ff.IsStyleAvailable(FontStyle.Underline)) fsLastDitch |= FontStyle.Underline;
                    }
                    // do I need to check for combinations? Say bold and italic combo exists, but
                    // just bold or just italic doesn't. What kind of typeface does this?!?!
                    // If I didn't know better, I'd point my finger at Verdana, but that's because
                    // I've a verdant vendetta against it...

                    if (ff.IsStyleAvailable(fsLastDitch))
                    {
                        usablefont = new Font(FontName, (float)FontSize, fsLastDitch);
                        // else I-don't-care-anymore
                    }
                    else if (ff.IsStyleAvailable(FontStyle.Regular))
                    {
                        usablefont = new Font(FontName, (float)FontSize, FontStyle.Regular);
                    }
                    else if (ff.IsStyleAvailable(FontStyle.Bold))
                    {
                        usablefont = new Font(FontName, (float)FontSize, FontStyle.Bold);
                    }
                    else if (ff.IsStyleAvailable(FontStyle.Italic))
                    {
                        usablefont = new Font(FontName, (float)FontSize, FontStyle.Italic);
                    }
                    else if (ff.IsStyleAvailable(FontStyle.Bold | FontStyle.Italic))
                    {
                        usablefont = new Font(FontName, (float)FontSize, FontStyle.Bold | FontStyle.Italic);
                    }
                    else
                    {
                        // the font name or typeface might not be installed (say on a web server),
                        // so we'll use a generic sans serif font as a fallback. What if this fails?
                        // I don't know... *can* it fail?
                        usablefont = new Font(FontFamily.GenericSansSerif, (float)FontSize);
                    }
                }
            }

            return usablefont;
        }

        internal static string ToDotNetFormatCode(string FormatCode)
        {
            //http://office.microsoft.com/en-sg/excel-help/number-format-codes-HP005198679.aspx
            //http://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo%28v=vs.90%29.aspx
            //http://msdn.microsoft.com/en-us/library/0c899ak8%28v=vs.71%29.aspx

            // NOTE: This is *not* meant as a definitive translation from Excel format code
            // to .NET format code.

            // For example, given 12.33333 and the format code "0 ?/?", I'm going to just pass "0 ?/?"
            // straight to the .NET string formatter. It'll come out as "12 ?/?", but I don't care.
            // The actual result is "12 1/3", but I'm using the result to measure width and height.
            // I don't actually need the resulting string to be correct.
            // The width of "?" is *kinda* close to "0" (or whatever digit it's supposed to be).
            // Of course, "12 ????/????" results in "12    1/3   " (note the extra spaces),
            // because ? is a digit placeholder, so in this case hopefully it's also *kinda* close
            // to the space character in terms of character width.
            // *Can* I figure out how to get a proper fraction? Maybe. It's not worth the effort though.
            // Continued fractions? http://en.wikipedia.org/wiki/Continued_fraction
            // Try entering a value of 0.787. Then use the format code "?/???". The resulting string
            // is "569/723". I want you to tell me how the frac(tion) did Excel get that.

            // Similarly, for dates, Excel uses lower case "m" to denote both months and minutes,
            // differentiating between either based on context. "d/mm/yyyy" means the "m" is for months.
            // "h:mm" means the "m" is for minutes.
            // In either case, the resulting string has 2 digits for the "mm" part.
            // We're measuring the width, so it doesn't matter if the month or minute is used.
            // In the following code, I'll try to get the context right, but it's not important.

            // What I'm going to do is clean up the Excel format code and make it such that the .NET
            // string formatter can use it (and won't spew out exceptions like there's no tomorrow).
            // Some considerations:
            // 1) The "General" format appears to be .NET format string "G10".
            // 2) "_)" in Excel means "use the width of ')'". I'm going to just use ")", meaning I'll
            //    replace "_)" with just ")". Apparently, you can also do "_-", meaning "use the width
            //    of '-'". So I'm going to replace anything that looks like underscore-somechar with
            //    somechar.
            // 3) "* " means "fill up the whole cell with space". I'm going to ignore this. Mainly because
            //    we're measuring width and height for the purpose of autofitting. There's no "filling up"
            //    at all.
            // 4) [whatever colour] will be taken out because we need the resulting string but not the
            //    colour. So "[Red]($#,##0.00)" becomes "($#,##0.00)". For academic interest, these are
            //    the "built-in" ones: [Black], [Blue], [Cyan], [Green], [Magenta], [Red], [White],
            //    [Yellow]. Apparently you can do [Color 5], referencing the 5th colour in the palette
            //    (out of 56 palette colours). And it's case-insensitive, so [BLACK] is valid.

            string result = FormatCode;

            // we do this to separate all the literal strings in the format code.
            // Odd-numbered indices hold the literal strings. We do stuff on the even-numbered indices.
            // For example, "\"General\" 0.00 \"0.00\""
            // becomes
            // Index 0: empty string
            // Index 1: General (string literal)
            // Index 2: 0.00
            // Index 3: 0.00 (string literal)
            // Assuming an input value of -1234.5678, the resulting string output is
            // -General 1234.57 0.00
            string[] saFormat = FormatCode.Split("\"".ToCharArray());
            for (int i = 0; i < saFormat.Length; ++i)
            {
                if (i % 2 == 0)
                {
                    if (saFormat[i].Length == 0) continue;

                    // this is for "General". It's case-insensitive, hence the upper and lower case.
                    saFormat[i] = Regex.Replace(saFormat[i], "[Gg][Ee][Nn][Ee][Rr][Aa][Ll]", SLConstants.GeneralFormatPlaceholder);

                    // make this "_)" to this ")"
                    saFormat[i] = Regex.Replace(saFormat[i], "_(.)", "$1");

                    // this removes the "* " filling up part. Note that the format code can be
                    // "*i" meaning fill up with i's. So we match with any character just to be safe.
                    saFormat[i] = Regex.Replace(saFormat[i], "\\*.", "");

                    // this removes all the [whatever colour] parts.
                    saFormat[i] = Regex.Replace(saFormat[i], "\\[[Bb][Ll][Aa][Cc][Kk]\\]", "");
                    saFormat[i] = Regex.Replace(saFormat[i], "\\[[Bb][Ll][Uu][Ee]\\]", "");
                    saFormat[i] = Regex.Replace(saFormat[i], "\\[[Cc][Yy][Aa][Nn]\\]", "");
                    saFormat[i] = Regex.Replace(saFormat[i], "\\[[Gg][Rr][Ee][Ee][Nn]\\]", "");
                    saFormat[i] = Regex.Replace(saFormat[i], "\\[[Mm][Aa][Gg][Ee][Nn][Tt][Aa]\\]", "");
                    saFormat[i] = Regex.Replace(saFormat[i], "\\[[Rr][Ee][Dd]\\]", "");
                    saFormat[i] = Regex.Replace(saFormat[i], "\\[[Ww][Hh][Ii][Tt][Ee]\\]", "");
                    saFormat[i] = Regex.Replace(saFormat[i], "\\[[Yy][Ee][Ll][Ll][Oo][Ww]\\]", "");
                    saFormat[i] = Regex.Replace(saFormat[i], "\\[[Cc][Oo][Ll][Oo][Uu]?[Rr].+?\\]", "");

                    // apparently Excel can display the first letter of the month given 5 m's.
                    // I'm gonna just use "W", on account that W probably has the widest width out
                    // of the Latin/Anglo alphabets. Why not "M" since it can be March or May?
                    // Because "M" is a valid .NET format string. I don't care enough to make this
                    // correct... Remember, we're going for relative accuracy, not perfect accuracy.
                    saFormat[i] = saFormat[i].Replace("mmmmm", "W");

                    // ok fine, I'm gonna try to make the month/minute context work...

                    // this is for elapsed hours, minutes and seconds. If there's a hint of d, m or y
                    // in front, Excel seems to take [h] as normal h. Similarly for [mm] and [ss].
                    // Apparently, if the format code is "d [h]:mm", the [h] becomes 12. My guess is
                    // that it's 12 hours. This is regardless of the elapsed time.
                    // If the format code is "d [mm]:ss", the [mm] becomes 720 (minutes), which is 12 hours.
                    // If the format code is "d [ss]", the [ss] becomes 43200 (seconds), which is ... 12 hours.
                    // I'm not going to enforce this...
                    saFormat[i] = Regex.Replace(saFormat[i], "((^|;)\\s*[^dmy]*\\s*)\\[h\\]", string.Format("$1{0}", SLConstants.ElapsedHourFormatPlaceholder));
                    saFormat[i] = Regex.Replace(saFormat[i], "((^|;)\\s*[^dmy]*\\s*)\\[mm\\]", string.Format("$1{0}", SLConstants.ElapsedMinuteFormatPlaceholder));
                    saFormat[i] = Regex.Replace(saFormat[i], "((^|;)\\s*[^dmy]*\\s*)\\[ss\\]", string.Format("$1{0}", SLConstants.ElapsedSecondFormatPlaceholder));

                    saFormat[i] = saFormat[i].Replace("[h]", "h");
                    saFormat[i] = saFormat[i].Replace("[mm]", "mm");
                    saFormat[i] = saFormat[i].Replace("[ss]", "ss");

                    // This has to be done after the 5 m's because .NET doesn't have a 5 m's format.
                    // It has to be done after the elapsed time thing too.
                    // We assume all m's are months first.
                    saFormat[i] = saFormat[i].Replace("m", "M");

                    // minutes have to be after the hour format or before the second format
                    saFormat[i] = Regex.Replace(saFormat[i], "h(.*)MM", "h$1mm");
                    saFormat[i] = Regex.Replace(saFormat[i], "h(.*)M", "h$1m");
                    saFormat[i] = Regex.Replace(saFormat[i], "MM(.*)s", "mm$1s");
                    saFormat[i] = Regex.Replace(saFormat[i], "M(.*)s", "m$1s");

                    // this makes for 24-hour clock. Doesn't really matter, but we can try for
                    // correctness when it's easy...
                    saFormat[i] = saFormat[i].Replace("hh", "HH");

                    // we force 2-digit hour if the AM/PM designator is missing.
                    // Excel seems to do this...
                    if (!Regex.IsMatch(saFormat[i], "[Aa][Mm]\\s*?/\\s*?[Pp][Mm]"))
                    {
                        saFormat[i] = saFormat[i].Replace("h", "HH");
                    }

                    // this is for AM/PM designator
                    saFormat[i] = Regex.Replace(saFormat[i], "[Aa][Mm]\\s*?/\\s*?[Pp][Mm]", "tt");

                    // this is for A/P designator (the abbreviated version of AM/PM)
                    saFormat[i] = Regex.Replace(saFormat[i], "[Aa]\\s*?/\\s*?[Pp]", "t");
                }
                else
                {
                    // re-surround literal string with double quotes
                    saFormat[i] = string.Format("\"{0}\"", saFormat[i]);
                }
            }

            result = string.Join("", saFormat);

            return result;
        }

        internal static bool CheckIfFormatCodeIsDateRelated(string FormatCode)
        {
            // is this definitive enough to check for Excel date formats? I don't know...
            // Original section of code:
            //if (sFormat.IndexOfAny("dDmMyYhHsS".ToCharArray()) >= 0) bIsDate = true;
            // The capital S caused SLGENERAL to be treated as a date format... dang...
            return (FormatCode.IndexOfAny("dmyhs".ToCharArray()) >= 0) ? true : false;
        }

        internal static string ToSampleDisplayFormat(double Data, string FormatCode, bool ThrowExceptionsIfAny)
        {
            // NOTE: This is NOT meant to display the given value in the way Excel displays it.
            // It's meant to simulate the resulting string so that I can measure the width and height.
            // As such, it doesn't have to be exactly correct, only "relatively correct".

            string result = string.Empty;

            string sGeneralResult = Data.ToString("G10", CultureInfo.InvariantCulture);

            bool bIsDate = false;
            // remove quoted text for testing date format
            string sFormat = Regex.Replace(FormatCode, "\".*?\"", "");
            bIsDate = CheckIfFormatCodeIsDateRelated(sFormat);

            if (bIsDate)
            {
                // I don't think it matters which epoch it's from...
                DateTime dt = SLTool.CalculateDateTimeFromDaysFromEpoch(Data, false);
                // we're going to ignore the 29 Feb 1900 bug because it's not worth getting it right...
                TimeSpan ts = dt - SLConstants.Epoch1900();
                try
                {
                    // will this blow up? Only time will tell...
                    result = dt.ToString(FormatCode);
                    result = result.Replace(SLConstants.GeneralFormatPlaceholder, sGeneralResult);
                    result = result.Replace(SLConstants.ElapsedHourFormatPlaceholder, ts.TotalHours.ToString("f0", CultureInfo.InvariantCulture));
                    result = result.Replace(SLConstants.ElapsedMinuteFormatPlaceholder, ts.TotalMinutes.ToString("f0", CultureInfo.InvariantCulture));
                    result = result.Replace(SLConstants.ElapsedSecondFormatPlaceholder, ts.TotalSeconds.ToString("f0", CultureInfo.InvariantCulture));
                }
                catch (Exception e)
                {
                    if (ThrowExceptionsIfAny)
                    {
                        throw e;
                    }
                    else
                    {
                        result = dt.ToString("dd/MM/yyyy");
                    }
                }
            }
            else
            {
                try
                {
                    // will this blow up? Only time will tell...
                    result = Data.ToString(FormatCode);
                    result = result.Replace(SLConstants.GeneralFormatPlaceholder, sGeneralResult);
                }
                catch (Exception e)
                {
                    if (ThrowExceptionsIfAny)
                    {
                        throw e;
                    }
                    else
                    {
                        result = Data.ToString("G10");
                    }
                }
            }

            return result;
        }

        internal static double ToRadian(double AngleInDegree)
        {
            return AngleInDegree * Math.PI / 180.0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Width"></param>
        /// <param name="Height"></param>
        /// <param name="Angle">This is in degrees, not radians.</param>
        /// <returns></returns>
        internal static SizeF CalculateOuterBoundsOfRotatedRectangle(float Width, float Height, double Angle)
        {
            // Read this for explanation of maths calculations
            // http://polymathprogrammer.com/2012/02/22/optimal-width-height-after-image-rotation/

            // get Angle to within 0 to 360 degrees
            while (Angle < 0) Angle += 360.0;
            while (Angle > 360) Angle -= 360.0;

            double fAngleBeta;
            // we can check like this because if an angle is greater than say 270, it's automatically
            // greater than 180. So we check larger angles first. Otherwise we'll have to check a range.
            if (Angle > 270) fAngleBeta = 360 - Angle;
            else if (Angle > 180) fAngleBeta = Angle - 180;
            else if (Angle > 90) fAngleBeta = 180 - Angle;
            else fAngleBeta = 90 - Angle;

            Angle = SLTool.ToRadian(Angle);
            fAngleBeta = SLTool.ToRadian(fAngleBeta);

            SizeF szf = new SizeF(0, 0);

            szf.Width = (float)(Width * Math.Cos(Angle) + Height * Math.Cos(fAngleBeta));
            szf.Height = (float)(Width * Math.Sin(Angle) + Height * Math.Sin(fAngleBeta));

            return szf;
        }

        internal static SizeF MeasureText(Bitmap bm, Graphics g, string Text, Font UsableFont)
        {
            string[] saText = Text.Replace("\r\n", "\n").Split('\n');
            SizeF[] szfa = new SizeF[saText.Length];

            float fDoubleUnderscoreWidth = g.MeasureString("__", UsableFont).Width;

            for (int i = 0; i < saText.Length; ++i)
            {
                szfa[i] = g.MeasureString(string.Format("_{0}_", saText[i]), UsableFont);
                szfa[i].Width = szfa[i].Width - fDoubleUnderscoreWidth;
                
                // Targetting .NET Standard means TextRenderer is no longer available
                // We will just have to make do with Graphics only for width and height.

                //// It seems that using Graphics.MeasureString() for the width and
                //// TextRenderer.MeasureText() for the height works out.
                //// No, I don't know why. Yes, I'm just as confused.

                //szfa[i] = g.MeasureString(string.Format("_{0}_", saText[i]), UsableFont);
                //szfa[i].Width = szfa[i].Width - fDoubleUnderscoreWidth;

                //// why do we use the average? I don't know. It appears to give the closest
                //// approximation to Excel's calculation formula.
                //// No, I don't know why. Yes, I'm just as confused.
                //// The result should be at most a few pixels off, so it isn't too bad.
                //szfa[i].Height = (szfa[i].Height + (float)TextRenderer.MeasureText(Text, UsableFont).Height) / 2.0f;
            }

            SizeF szf = new SizeF(0, 0);
            for (int i = 0; i < szfa.Length; ++i)
            {
                if (szfa[i].Height > szf.Height) szf.Height = szfa[i].Height;
                if (szfa[i].Width > szf.Width) szf.Width = szfa[i].Width;
            }
            // if there are 4 lines, then we multiply the height by 4
            szf.Height *= saText.Length;

            return szf;
        }

        internal static SLCellPointRange TranslateReferenceToCellPointRange(string Reference)
        {
            SLCellPointRange pt = new SLCellPointRange();
            int index;
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iSwap = 0;

            index = Reference.IndexOf(":");
            if (index > -1)
            {
                if (SLTool.FormatCellReferenceRangeToRowColumnIndex(Reference, out iStartRowIndex, out iStartColumnIndex, out iEndRowIndex, out iEndColumnIndex))
                {
                    // in case the cell ranges aren't top-left to bottom-right
                    // But if you're using Excel, Excel would've taken care of this.
                    if (iStartRowIndex > iEndRowIndex)
                    {
                        iSwap = iStartRowIndex;
                        iStartRowIndex = iEndRowIndex;
                        iEndRowIndex = iSwap;
                    }
                    if (iStartColumnIndex > iEndColumnIndex)
                    {
                        iSwap = iStartColumnIndex;
                        iStartColumnIndex = iEndColumnIndex;
                        iEndColumnIndex = iSwap;
                    }

                    pt = new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
                }
            }
            else
            {
                if (SLTool.FormatCellReferenceToRowColumnIndex(Reference, out iStartRowIndex, out iStartColumnIndex))
                {
                    pt = new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iStartRowIndex, iStartColumnIndex);
                }
            }

            return pt;
        }

        internal static List<SLCellPointRange> TranslateReferenceToCellPointRangeList(string Reference)
        {
            List<SLCellPointRange> pts = new List<SLCellPointRange>();
            SLCellPointRange pt;
            int index;
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iSwap = 0;

            string[] sa = Reference.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < sa.Length; ++i)
            {
                index = Reference.IndexOf(":");
                if (index > -1)
                {
                    if (SLTool.FormatCellReferenceRangeToRowColumnIndex(Reference, out iStartRowIndex, out iStartColumnIndex, out iEndRowIndex, out iEndColumnIndex))
                    {
                        // in case the cell ranges aren't top-left to bottom-right
                        // But if you're using Excel, Excel would've taken care of this.
                        if (iStartRowIndex > iEndRowIndex)
                        {
                            iSwap = iStartRowIndex;
                            iStartRowIndex = iEndRowIndex;
                            iEndRowIndex = iSwap;
                        }
                        if (iStartColumnIndex > iEndColumnIndex)
                        {
                            iSwap = iStartColumnIndex;
                            iStartColumnIndex = iEndColumnIndex;
                            iEndColumnIndex = iSwap;
                        }

                        pt = new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
                        pts.Add(pt);
                    }
                }
                else
                {
                    if (SLTool.FormatCellReferenceToRowColumnIndex(Reference, out iStartRowIndex, out iStartColumnIndex))
                    {
                        pt = new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iStartRowIndex, iStartColumnIndex);
                        pts.Add(pt);
                    }
                }
            }

            return pts;
        }

        internal static List<SLCellPointRange> TranslateSeqRefToCellPointRange(ListValue<StringValue> SeqRef)
        {
            List<SLCellPointRange> pts = new List<SLCellPointRange>();

            SLCellPointRange pt;
            int index;
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iSwap = 0;

            foreach (StringValue s in SeqRef.Items)
            {
                index = s.Value.IndexOf(":");
                if (index > -1)
                {
                    if (SLTool.FormatCellReferenceRangeToRowColumnIndex(s.Value, out iStartRowIndex, out iStartColumnIndex, out iEndRowIndex, out iEndColumnIndex))
                    {
                        // in case the cell ranges aren't top-left to bottom-right
                        // But if you're using Excel, Excel would've taken care of this.
                        if (iStartRowIndex > iEndRowIndex)
                        {
                            iSwap = iStartRowIndex;
                            iStartRowIndex = iEndRowIndex;
                            iEndRowIndex = iSwap;
                        }
                        if (iStartColumnIndex > iEndColumnIndex)
                        {
                            iSwap = iStartColumnIndex;
                            iStartColumnIndex = iEndColumnIndex;
                            iEndColumnIndex = iSwap;
                        }

                        pt = new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
                        pts.Add(pt);
                    }
                }
                else
                {
                    if (SLTool.FormatCellReferenceToRowColumnIndex(s.Value, out iStartRowIndex, out iStartColumnIndex))
                    {
                        pt = new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iStartRowIndex, iStartColumnIndex);
                        pts.Add(pt);
                    }
                }
            }

            return pts;
        }

        internal static ListValue<StringValue> TranslateCellPointRangeToSeqRef(List<SLCellPointRange> PointRange)
        {
            ListValue<StringValue> list = new ListValue<StringValue>();

            string sRef = string.Empty;
            foreach (SLCellPointRange pt in PointRange)
            {
                if (pt.StartRowIndex == pt.EndRowIndex && pt.StartColumnIndex == pt.EndColumnIndex)
                {
                    sRef = SLTool.ToCellReference(pt.StartRowIndex, pt.StartColumnIndex);
                }
                else
                {
                    sRef = string.Format("{0}:{1}", SLTool.ToCellReference(pt.StartRowIndex, pt.StartColumnIndex), SLTool.ToCellReference(pt.EndRowIndex, pt.EndColumnIndex));
                }
                list.Items.Add(new StringValue(sRef));
            }

            return list;
        }

        internal static List<SLCellPointRange> TranslateRefSeqToCellPointRange(Excel.ReferenceSequence RefSeq)
        {
            List<SLCellPointRange> pts = new List<SLCellPointRange>();

            SLCellPointRange pt;
            int index;
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iSwap = 0;

            string[] saRef = RefSeq.Text.Split(" ".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            foreach (string s in saRef)
            {
                index = s.IndexOf(":");
                if (index > -1)
                {
                    if (SLTool.FormatCellReferenceRangeToRowColumnIndex(s, out iStartRowIndex, out iStartColumnIndex, out iEndRowIndex, out iEndColumnIndex))
                    {
                        // in case the cell ranges aren't top-left to bottom-right
                        // But if you're using Excel, Excel would've taken care of this.
                        if (iStartRowIndex > iEndRowIndex)
                        {
                            iSwap = iStartRowIndex;
                            iStartRowIndex = iEndRowIndex;
                            iEndRowIndex = iSwap;
                        }
                        if (iStartColumnIndex > iEndColumnIndex)
                        {
                            iSwap = iStartColumnIndex;
                            iStartColumnIndex = iEndColumnIndex;
                            iEndColumnIndex = iSwap;
                        }

                        pt = new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
                        pts.Add(pt);
                    }
                }
                else
                {
                    if (SLTool.FormatCellReferenceToRowColumnIndex(s, out iStartRowIndex, out iStartColumnIndex))
                    {
                        pt = new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iStartRowIndex, iStartColumnIndex);
                        pts.Add(pt);
                    }
                }
            }

            return pts;
        }

        internal static string TranslateCellPointRangeToRefSeq(List<SLCellPointRange> PointRange)
        {
            string result = string.Empty;

            string sRef = string.Empty;
            foreach (SLCellPointRange pt in PointRange)
            {
                if (pt.StartRowIndex == pt.EndRowIndex && pt.StartColumnIndex == pt.EndColumnIndex)
                {
                    sRef = SLTool.ToCellReference(pt.StartRowIndex, pt.StartColumnIndex);
                }
                else
                {
                    sRef = string.Format("{0}:{1}", SLTool.ToCellReference(pt.StartRowIndex, pt.StartColumnIndex), SLTool.ToCellReference(pt.EndRowIndex, pt.EndColumnIndex));
                }

                result += string.Format(" {0}", sRef);
            }
            // get rid of the starting space. Or we could just check within the loop.
            result = result.TrimStart();

            return result;
        }

        internal static string TranslateCellPointRangeToReference(SLCellPointRange Range)
        {
            string result = string.Empty;
            if (Range.StartRowIndex == Range.EndRowIndex && Range.StartColumnIndex == Range.EndColumnIndex)
            {
                result = SLTool.ToCellReference(Range.StartRowIndex, Range.StartColumnIndex);
            }
            else
            {
                result = SLTool.ToCellRange(Range.StartRowIndex, Range.StartColumnIndex, Range.EndRowIndex, Range.EndColumnIndex);
            }

            return result;
        }
    }
}
