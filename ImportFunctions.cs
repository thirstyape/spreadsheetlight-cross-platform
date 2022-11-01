using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        private List<string> GetPossibleDateFormatsForImportParsing(List<string> list1, List<string> list2, List<string> list3)
        {
            List<string> result = new List<string>();
            int i, j, k;
            for (i = 0; i < list1.Count; ++i)
            {
                for (j = 0; j < list2.Count; ++j)
                {
                    for (k = 0; k < list3.Count; ++k)
                    {
                        // we'll use a space as a separator. When the date text is obtained,
                        // we'll remove any date separators and replace it with a space.
                        // And just for good measure, we add a version without separators too.
                        // Example: "--//05 / /-.-/ Oct -+...-+- 2013" becomes "05 Oct 2013".
                        // Yes I know typically it's either "05-Oct-2013" or "05.Oct.2013" or
                        // whatever date formats other parts of the world use. I'm just trying
                        // to be thorough...
                        // "But some of the formats don't make sense!"
                        // You wanna filter those that don't make sense out of all the possible
                        // combinations? Be my guest. I'm not worldly enough to know which
                        // aren't possible in the first place, so I'll work with everything and
                        // let the .NET DateTime parser figure it out.
                        result.Add(string.Format("{0} {1} {2}", list1[i], list2[j], list3[k]));
                        result.Add(string.Format("{0}{1}{2}", list1[i], list2[j], list3[k]));
                    }
                }
            }

            return result;
        }

        // You think I'm *good* at naming functions? I just needed a function so I don't have to type so much...
        private void SetDateIfFailThenSetAsTextForImportParsing(string TextData, int RowIndex, int ColumnIndex, List<string> CustomDateFormats, CultureInfo Culture, bool HasTextQualifier, char TextQualifier, bool PreserveSpace)
        {
            DateTime dtData = DateTime.Now;
            bool bSuccess = false;
            if (CustomDateFormats.Count > 0)
            {
                bSuccess = DateTime.TryParseExact(TextData, CustomDateFormats.ToArray(), Culture, System.Globalization.DateTimeStyles.None, out dtData);
            }

            if (bSuccess)
            {
                SetCellValue(RowIndex, ColumnIndex, dtData);
            }
            else
            {
                // try parsing in a generic way since the custom date formats failed
                // (or that there are no custom date formats given).
                if (DateTime.TryParse(TextData, Culture, System.Globalization.DateTimeStyles.None, out dtData))
                {
                    SetCellValue(RowIndex, ColumnIndex, dtData);
                }
                else
                {
                    // Apparently there's no standard way to deal with consecutive text qualifiers.
                    // Excel treats ""honeydew"" as honeydew""
                    // LibreOffice Calc treats ""honeydew"" as "honeydew", mooshing any columns after it
                    // as the same column.
                    // I'm gonna just treat ""honeydew"" as honeydew (trimming off everything).
                    string sData = TextData;
                    if (HasTextQualifier) sData = sData.Trim(TextQualifier);
                    if (!PreserveSpace) sData = sData.Trim();
                    SetCellValue(RowIndex, ColumnIndex, sData);
                }
            }
        }

        /// <summary>
        /// Import a text file as a data source, with the first data row and first data column at a specific cell.
        /// </summary>
        /// <param name="FileName">The file name.</param>
        /// <param name="AnchorCellReference">The anchor cell reference, such as "A1".</param>
        public void ImportText(string FileName, string AnchorCellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iRowIndex, out iColumnIndex))
            {
                this.ImportText(FileName, iRowIndex, iColumnIndex, null);
            }
        }

        /// <summary>
        /// Import a text file as a data source, with the first data row and first data column at a specific cell.
        /// </summary>
        /// <param name="FileName">The file name.</param>
        /// <param name="AnchorCellReference">The anchor cell reference, such as "A1".</param>
        /// <param name="Options">Text import options.</param>
        public void ImportText(string FileName, string AnchorCellReference, SLTextImportOptions Options)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iRowIndex, out iColumnIndex))
            {
                this.ImportText(FileName, iRowIndex, iColumnIndex, Options);
            }
        }

        /// <summary>
        /// Import a text file as a data source, with the first data row and first data column at a specific cell.
        /// </summary>
        /// <param name="FileName">The file name.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor cell.</param>
        public void ImportText(string FileName, int AnchorRowIndex, int AnchorColumnIndex)
        {
            this.ImportText(FileName, AnchorRowIndex, AnchorColumnIndex, null);
        }

        /// <summary>
        /// Import a text file as a data source, with the first data row and first data column at a specific cell.
        /// </summary>
        /// <param name="FileName">The file name.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor cell.</param>
        /// <param name="Options">Text import options.</param>
        public void ImportText(string FileName, int AnchorRowIndex, int AnchorColumnIndex, SLTextImportOptions Options)
        {
            if (Options == null) Options = new SLTextImportOptions();
            if (AnchorRowIndex < 1) AnchorRowIndex = 1;
            if (AnchorColumnIndex < 1) AnchorColumnIndex = 1;

            List<char> listDelimiters = new List<char>();
            if (Options.UseTabDelimiter) listDelimiters.Add('\t');
            if (Options.UseSemicolonDelimiter) listDelimiters.Add(';');
            if (Options.UseCommaDelimiter) listDelimiters.Add(',');
            if (Options.UseSpaceDelimiter) listDelimiters.Add(' ');
            if (Options.UseCustomDelimiter) listDelimiters.Add(Options.CustomDelimiter);
            char[] caDelimiters = listDelimiters.ToArray();
            StringSplitOptions sso = Options.MergeDelimiters ? StringSplitOptions.RemoveEmptyEntries : StringSplitOptions.None;

            string sDataLine;
            List<string> listData = new List<string>();
            double fData;
            DateTime dtData = DateTime.Now;
            string sDateData = string.Empty;

            // There's the space separator: \s
            // Example: 05 Oct 2013
            // Then the typical slash separator: /
            // Example: 05/10/2013 or 10/05/2013
            // Then the dash: - (we do \- to escape for regex)
            // Example: 05-10-2013 or 10-05-2013
            // Then the plus sign: + (we do \+ to escape for regex).
            // The plus sign occurs for expanded year representions, such as the year +12345
            // Will SpreadsheetLight survive beyond the year 9999? Who knows?
            // See this for more information:
            //http://en.wikipedia.org/wiki/ISO_8601
            // The dot (or period) is also used: . (we do \. to escape for regex)
            // Then the single quote is used when the century portion of the year is shortened.
            // Example: 05.10.2013 or 05.10.'13
            // See this for more information:
            //http://en.wikipedia.org/wiki/Date_and_time_notation_in_Europe
            string sDateSeparatorRegex = @"[\s/\-\+\.']+";

            string sData;
            int i;
            int iRowIndex, iColumnIndex;
            int iNextSubstringIndex, iFixedWidth;
            SLTextImportColumnFormatValues tiColumnFormat = SLTextImportColumnFormatValues.General;

            List<string> listYears = new List<string> { "yyyy", "yy" };
            List<string> listMonths = new List<string> { "MMMM", "MMM", "MM", "M" };
            List<string> listDays = new List<string> { "dd", "d" };
            List<string> listMDY = GetPossibleDateFormatsForImportParsing(listMonths, listDays, listYears);
            List<string> listDMY = GetPossibleDateFormatsForImportParsing(listDays, listMonths, listYears);
            List<string> listYMD = GetPossibleDateFormatsForImportParsing(listYears, listMonths, listDays);
            List<string> listMYD = GetPossibleDateFormatsForImportParsing(listMonths, listYears, listDays);
            List<string> listDYM = GetPossibleDateFormatsForImportParsing(listDays, listYears, listMonths);
            List<string> listYDM = GetPossibleDateFormatsForImportParsing(listYears, listDays, listMonths);

            iRowIndex = AnchorRowIndex;
            int iRowCounter = 0;
            using (StreamReader sr = new StreamReader(FileName, Options.Encoding))
            {
                while (sr.Peek() > -1)
                {
                    sDataLine = sr.ReadLine();
                    ++iRowCounter;
                    if (iRowCounter < Options.ImportStartRowIndex) continue;

                    listData.Clear();
                    if (Options.DataFieldType == SLTextImportDataFieldTypeValues.Delimited)
                    {
                        listData = new List<string>(sDataLine.Split(caDelimiters, sso));
                    }
                    else
                    {
                        // else is fixed width

                        iNextSubstringIndex = 0;
                        iFixedWidth = Options.DefaultFixedWidth;
                        // use i temporarily for tracking column indices
                        i = 1;
                        if (Options.dictFixedWidth.ContainsKey(i)) iFixedWidth = Options.dictFixedWidth[i];
                        while (iNextSubstringIndex + iFixedWidth <= sDataLine.Length)
                        {
                            listData.Add(sDataLine.Substring(iNextSubstringIndex, iFixedWidth));
                            iNextSubstringIndex += iFixedWidth;
                            ++i;
                            iFixedWidth = Options.DefaultFixedWidth;
                            if (Options.dictFixedWidth.ContainsKey(i)) iFixedWidth = Options.dictFixedWidth[i];
                        }

                        // if still need to do substring, but the fixed width exceeded the string length...
                        if (iNextSubstringIndex < sDataLine.Length)
                        {
                            // then take the rest of the string
                            listData.Add(sDataLine.Substring(iNextSubstringIndex));
                        }
                        // no else because all the data has been fixed-width-separated by now.
                    }

                    iColumnIndex = AnchorColumnIndex;
                    for (i = 0; i < listData.Count; ++i)
                    {
                        tiColumnFormat = SLTextImportColumnFormatValues.General;
                        // +1 because i is zero-based
                        if (Options.dictColumnFormat.ContainsKey(i + 1))
                        {
                            tiColumnFormat = Options.dictColumnFormat[i + 1];
                        }

                        switch (tiColumnFormat)
                        {
                            case SLTextImportColumnFormatValues.General:
                                // We try to parse as a floating point number first.
                                // Failing that, we try to parse as date with any given custom date formats.
                                // If fail that or there are no custom date formats given, we try to parse
                                // as date in a general manner.
                                // If fail *that*, then we throw in the towel and just set as text.
                                if (double.TryParse(listData[i], Options.NumberStyles, Options.Culture, out fData))
                                {
                                    SetCellValue(iRowIndex, iColumnIndex, fData);
                                }
                                else
                                {
                                    SetDateIfFailThenSetAsTextForImportParsing(listData[i], iRowIndex, iColumnIndex, Options.listCustomDateFormats, Options.Culture, Options.HasTextQualifier, Options.TextQualifier, Options.PreserveSpace);
                                }
                                break;
                            case SLTextImportColumnFormatValues.Text:
                                sData = listData[i];
                                if (Options.HasTextQualifier) sData = sData.Trim(Options.TextQualifier);
                                if (!Options.PreserveSpace) sData = sData.Trim();
                                SetCellValue(iRowIndex, iColumnIndex, sData);
                                break;
                            case SLTextImportColumnFormatValues.DateMDY:
                                // we try to make the date string as compact and as close to what
                                // we have as date formats as possible, before trying the date combos.
                                sDateData = Regex.Replace(listData[i], sDateSeparatorRegex, " ").Trim();
                                if (DateTime.TryParseExact(sDateData, listMDY.ToArray(), Options.Culture, System.Globalization.DateTimeStyles.None, out dtData))
                                {
                                    SetCellValue(iRowIndex, iColumnIndex, dtData);
                                }
                                else
                                {
                                    SetDateIfFailThenSetAsTextForImportParsing(listData[i], iRowIndex, AnchorColumnIndex + i, Options.listCustomDateFormats, Options.Culture, Options.HasTextQualifier, Options.TextQualifier, Options.PreserveSpace);
                                }
                                break;
                            case SLTextImportColumnFormatValues.DateDMY:
                                sDateData = Regex.Replace(listData[i], sDateSeparatorRegex, " ").Trim();
                                if (DateTime.TryParseExact(sDateData, listDMY.ToArray(), Options.Culture, System.Globalization.DateTimeStyles.None, out dtData))
                                {
                                    SetCellValue(iRowIndex, iColumnIndex, dtData);
                                }
                                else
                                {
                                    SetDateIfFailThenSetAsTextForImportParsing(listData[i], iRowIndex, AnchorColumnIndex + i, Options.listCustomDateFormats, Options.Culture, Options.HasTextQualifier, Options.TextQualifier, Options.PreserveSpace);
                                }
                                break;
                            case SLTextImportColumnFormatValues.DateYMD:
                                sDateData = Regex.Replace(listData[i], sDateSeparatorRegex, " ").Trim();
                                if (DateTime.TryParseExact(sDateData, listYMD.ToArray(), Options.Culture, System.Globalization.DateTimeStyles.None, out dtData))
                                {
                                    SetCellValue(iRowIndex, iColumnIndex, dtData);
                                }
                                else
                                {
                                    SetDateIfFailThenSetAsTextForImportParsing(listData[i], iRowIndex, AnchorColumnIndex + i, Options.listCustomDateFormats, Options.Culture, Options.HasTextQualifier, Options.TextQualifier, Options.PreserveSpace);
                                }
                                break;
                            case SLTextImportColumnFormatValues.DateMYD:
                                sDateData = Regex.Replace(listData[i], sDateSeparatorRegex, " ").Trim();
                                if (DateTime.TryParseExact(sDateData, listMYD.ToArray(), Options.Culture, System.Globalization.DateTimeStyles.None, out dtData))
                                {
                                    SetCellValue(iRowIndex, iColumnIndex, dtData);
                                }
                                else
                                {
                                    SetDateIfFailThenSetAsTextForImportParsing(listData[i], iRowIndex, AnchorColumnIndex + i, Options.listCustomDateFormats, Options.Culture, Options.HasTextQualifier, Options.TextQualifier, Options.PreserveSpace);
                                }
                                break;
                            case SLTextImportColumnFormatValues.DateDYM:
                                sDateData = Regex.Replace(listData[i], sDateSeparatorRegex, " ").Trim();
                                if (DateTime.TryParseExact(sDateData, listDYM.ToArray(), Options.Culture, System.Globalization.DateTimeStyles.None, out dtData))
                                {
                                    SetCellValue(iRowIndex, iColumnIndex, dtData);
                                }
                                else
                                {
                                    SetDateIfFailThenSetAsTextForImportParsing(listData[i], iRowIndex, iColumnIndex, Options.listCustomDateFormats, Options.Culture, Options.HasTextQualifier, Options.TextQualifier, Options.PreserveSpace);
                                }
                                break;
                            case SLTextImportColumnFormatValues.DateYDM:
                                sDateData = Regex.Replace(listData[i], sDateSeparatorRegex, " ").Trim();
                                if (DateTime.TryParseExact(sDateData, listYDM.ToArray(), Options.Culture, System.Globalization.DateTimeStyles.None, out dtData))
                                {
                                    SetCellValue(iRowIndex, iColumnIndex, dtData);
                                }
                                else
                                {
                                    SetDateIfFailThenSetAsTextForImportParsing(listData[i], iRowIndex, iColumnIndex, Options.listCustomDateFormats, Options.Culture, Options.HasTextQualifier, Options.TextQualifier, Options.PreserveSpace);
                                }
                                break;
                        }

                        if (tiColumnFormat != SLTextImportColumnFormatValues.Skip)
                        {
                            ++iColumnIndex;
                        }
                    }

                    ++iRowIndex;
                }
            }
        }

        // merging spreadsheets is kinda like importing data, right? So it's here then.

//        public void MergeSpreadsheet(string SpreadsheetFileName)
//        {
//            this.MergeSpreadsheet(true, SpreadsheetFileName, null);
//        }

//        public void MergeSpreadsheet(Stream SpreadsheetStream)
//        {
//            this.MergeSpreadsheet(false, null, SpreadsheetStream);
//        }

//        private void MergeSpreadsheet(bool IsFile, string SpreadsheetFileName, Stream SpreadsheetStream)
//        {
//            using (MemoryStream msAnother = new MemoryStream())
//            {
//                if (IsFile)
//                {
//                    byte[] baData = File.ReadAllBytes(SpreadsheetFileName);
//                    msAnother.Write(baData, 0, baData.Length);
//                }
//                else
//                {
//                    SpreadsheetStream.Position = 0;
//                    byte[] baData = new byte[SpreadsheetStream.Length];
//                    SpreadsheetStream.Read(baData, 0, baData.Length);
//                    msAnother.Write(baData, 0, baData.Length);
//                }

//                using (SpreadsheetDocument xlAnother = SpreadsheetDocument.Open(msAnother, false))
//                {
//                    HashSet<string> hsCurrentSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
//                    foreach (SLSheet sheet in slwb.Sheets)
//                    {
//                        // current sheet names supposed to be unique, so I'm not checking for collisions.
//                        hsCurrentSheetNames.Add(sheet.Name);
//                    }

//                    HashSet<string> hsAnotherSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
//                    List<string> listAnotherSheetNames = new List<string>();
//                    using (OpenXmlReader oxr = OpenXmlReader.Create(xlAnother.WorkbookPart.Workbook.Sheets))
//                    {
//                        string sSheetName;
//                        while (oxr.Read())
//                        {
//                            if (oxr.ElementType == typeof(Sheet))
//                            {
//                                sSheetName = ((Sheet)oxr.LoadCurrentElement()).Name.Value;
//                                hsAnotherSheetNames.Add(sSheetName);
//                                listAnotherSheetNames.Add(sSheetName);
//                            }
//                        }
//                    }

////Sheet1
////Sheet2
////Sheet3

////Sheet1 -> Sheet7 -> Sheet8
////Sheet6
////Sheet7

//                    Dictionary<string, string> dictAnotherNewSheetNames = new Dictionary<string, string>();
//                    foreach (string s in listAnotherSheetNames)
//                    {
//                    }
//                }
//                // end of using SpreadsheetDocument
//            }
//        }

        /// <summary>
        /// Import a System.Data.DataTable as a data source, with the first data row and first data column at a specific cell.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The data table.</param>
        /// <param name="IncludeHeader">True if the data table's column names are to be used in the first row as a header row. False otherwise.</param>
        public void ImportDataTable(string CellReference, DataTable Data, bool IncludeHeader)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return;
            }

            ImportDataTable(iRowIndex, iColumnIndex, Data, IncludeHeader);
        }

        /// <summary>
        /// Import a System.Data.DataTable as a data source, with the first data row and first data column at a specific cell.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The data table.</param>
        /// <param name="IncludeHeader">True if the data table's column names are to be used in the first row as a header row. False otherwise.</param>
        public void ImportDataTable(int RowIndex, int ColumnIndex, DataTable Data, bool IncludeHeader)
        {
            int i, j;
            Type[] taColumns;
            string[] saColumnNames;
            int iDefaultColumnLength = 10;
            int iColumnLength = 0;

            if (Data.Columns.Count == 0)
            {
                iColumnLength = iDefaultColumnLength;
                taColumns = new Type[iColumnLength];
                saColumnNames = new string[iColumnLength];
                for (i = 0; i < iColumnLength; ++i)
                {
                    taColumns[i] = typeof(string);
                    saColumnNames[i] = string.Format("Column{0}", i + 1);
                }
            }
            else
            {
                iColumnLength = Data.Columns.Count;
                taColumns = new Type[iColumnLength];
                saColumnNames = new string[iColumnLength];
                for (i = 0; i < iColumnLength; ++i)
                {
                    taColumns[i] = Data.Columns[i].DataType;
                    saColumnNames[i] = Data.Columns[i].ColumnName;
                }
            }

            // "Optimisation" order:
            // double, float, decimal, int, long, string, DateTime,
            // short, ushort, uint, ulong, char, byte, sbyte, bool,
            // TimeSpan, byte[]

            if (IncludeHeader)
            {
                for (i = 0; i < iColumnLength; ++i)
                {
                    this.SetCellValue(RowIndex, ColumnIndex + i, saColumnNames[i]);
                }

                // get to the next row for the data part
                ++RowIndex;
            }

            int iRowCount = Data.Rows.Count;
            int iItemCount;
            int iRowIndex, iColumnIndex;
            DataRow dr;
            for (i = 0; i < iRowCount; ++i)
            {
                iRowIndex = RowIndex + i;
                dr = Data.Rows[i];
                iItemCount = dr.ItemArray.Length;
                for (j = 0; j < iItemCount; ++j)
                {
                    iColumnIndex = ColumnIndex + j;
                    if (j <= iColumnLength)
                    {
                        // in case the the data table cell is DBNull
                        // This code part sent in by Troye Stonich.
                        if (dr.ItemArray[j].GetType() == typeof(System.DBNull))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, string.Empty);
                            continue;
                        }

                        if (taColumns[j] == typeof(double))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (double)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(float))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (float)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(decimal))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (decimal)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(int))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (int)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(long))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (long)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(string))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (string)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(DateTime))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (DateTime)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(short))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (short)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(ushort))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (ushort)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(uint))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (uint)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(ulong))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (ulong)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(char))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (char)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(byte))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (byte)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(sbyte))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (sbyte)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(bool))
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, (bool)dr.ItemArray[j]);
                        }
                        else if (taColumns[j] == typeof(TimeSpan))
                        {
                            // what do you do with TimeSpans?
                            this.SetCellValue(iRowIndex, iColumnIndex, ((TimeSpan)dr.ItemArray[j]).ToString());
                        }
                        // what do you do with byte[]?
                        //else if (taColumns[j] == typeof(byte[]))
                        //{
                        //}
                        else
                        {
                            this.SetCellValue(iRowIndex, iColumnIndex, dr.ItemArray[j].ToString());
                        }
                    }
                    else
                    {
                        // this value is in the data row, but isn't defined in the columns
                        this.SetCellValue(iRowIndex, iColumnIndex, (string)dr.ItemArray[j]);
                    }
                }
            }
        }
    }
}
