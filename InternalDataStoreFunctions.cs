using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    partial class SLDocument
    {
        internal int SaveToStylesheet(string Hash)
        {
            int index = 0;
            if (dictStyleHash.ContainsKey(Hash))
            {
                index = dictStyleHash[Hash];
            }
            else
            {
                index = listStyle.Count;
                listStyle.Add(Hash);
                dictStyleHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheet(string Hash)
        {
            int index = listStyle.Count;
            listStyle.Add(Hash);
            dictStyleHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetNumberingFormat(string Hash)
        {
            int index = 0;
            if (dictStyleNumberingFormatHash.ContainsKey(Hash))
            {
                index = dictStyleNumberingFormatHash[Hash];
            }
            else if (dictBuiltInNumberingFormatHash.ContainsKey(Hash))
            {
                index = dictBuiltInNumberingFormatHash[Hash];
            }
            else
            {
                index = NextNumberFormatId;
                ++NextNumberFormatId;
                dictStyleNumberingFormat[index] = Hash;
                dictStyleNumberingFormatHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetNumberingFormat(int index, string Hash)
        {
            dictStyleNumberingFormat[index] = Hash;
            dictStyleNumberingFormatHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetFont(string Hash)
        {
            int index = 0;
            if (dictStyleFontHash.ContainsKey(Hash))
            {
                index = dictStyleFontHash[Hash];
            }
            else
            {
                index = listStyleFont.Count;
                listStyleFont.Add(Hash);
                dictStyleFontHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetFont(string Hash)
        {
            int index = listStyleFont.Count;
            listStyleFont.Add(Hash);
            dictStyleFontHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetFill(string Hash)
        {
            int index = 0;
            if (dictStyleFillHash.ContainsKey(Hash))
            {
                index = dictStyleFillHash[Hash];
            }
            else
            {
                index = listStyleFill.Count;
                listStyleFill.Add(Hash);
                dictStyleFillHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetFill(string Hash)
        {
            int index = listStyleFill.Count;
            listStyleFill.Add(Hash);
            dictStyleFillHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetBorder(string Hash)
        {
            int index = 0;
            if (dictStyleBorderHash.ContainsKey(Hash))
            {
                index = dictStyleBorderHash[Hash];
            }
            else
            {
                index = listStyleBorder.Count;
                listStyleBorder.Add(Hash);
                dictStyleBorderHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetBorder(string Hash)
        {
            int index = listStyleBorder.Count;
            listStyleBorder.Add(Hash);
            dictStyleBorderHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetCellStylesFormat(string Hash)
        {
            int index = 0;
            if (dictStyleCellStyleFormatHash.ContainsKey(Hash))
            {
                index = dictStyleCellStyleFormatHash[Hash];
            }
            else
            {
                index = listStyleCellStyleFormat.Count;
                listStyleCellStyleFormat.Add(Hash);
                dictStyleCellStyleFormatHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetCellStylesFormat(string Hash)
        {
            int index = listStyleCellStyleFormat.Count;
            listStyleCellStyleFormat.Add(Hash);
            dictStyleCellStyleFormatHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetCellStyle(string Hash)
        {
            int index = 0;
            if (dictStyleCellStyleHash.ContainsKey(Hash))
            {
                index = dictStyleCellStyleHash[Hash];
            }
            else
            {
                index = listStyleCellStyle.Count;
                listStyleCellStyle.Add(Hash);
                dictStyleCellStyleHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetCellStyle(string Hash)
        {
            int index = listStyleCellStyle.Count;
            listStyleCellStyle.Add(Hash);
            dictStyleCellStyleHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetDifferentialFormat(string Hash)
        {
            int index = 0;
            if (dictStyleDifferentialFormatHash.ContainsKey(Hash))
            {
                index = dictStyleDifferentialFormatHash[Hash];
            }
            else
            {
                index = listStyleDifferentialFormat.Count;
                listStyleDifferentialFormat.Add(Hash);
                dictStyleDifferentialFormatHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetDifferentialFormat(string Hash)
        {
            int index = listStyleDifferentialFormat.Count;
            listStyleDifferentialFormat.Add(Hash);
            dictStyleDifferentialFormatHash[Hash] = index;

            return index;
        }

        internal void ForceSaveToSharedStringTable(SharedStringItem ssi)
        {
            int index = listSharedString.Count;
            string sHash = SLTool.RemoveNamespaceDeclaration(ssi.InnerXml);
            listSharedString.Add(sHash);
            dictSharedStringHash[sHash] = index;
        }

        internal int DirectSaveToSharedStringTable(string Data)
        {
            int index = 0;
            string sHash;
            if (SLTool.ToPreserveSpace(Data))
            {
                sHash = string.Format("<x:t xml:space=\"preserve\">{0}</x:t>", Data);
            }
            else
            {
                sHash = string.Format("<x:t>{0}</x:t>", Data);
            }

            if (dictSharedStringHash.ContainsKey(sHash))
            {
                index = dictSharedStringHash[sHash];
            }
            else
            {
                index = listSharedString.Count;
                listSharedString.Add(sHash);
                dictSharedStringHash[sHash] = index;
            }

            if (!hsUniqueSharedString.Contains(Data))
            {
                hsUniqueSharedString.Add(Data);
            }

            return index;
        }

        internal int DirectSaveToSharedStringTable(InlineString Data)
        {
            int index = 0;
            string sHash = SLTool.RemoveNamespaceDeclaration(Data.InnerXml);
            if (dictSharedStringHash.ContainsKey(sHash))
            {
                index = dictSharedStringHash[sHash];
            }
            else
            {
                index = listSharedString.Count;
                listSharedString.Add(sHash);
                dictSharedStringHash[sHash] = index;
            }

            if (!hsUniqueSharedString.Contains(sHash))
            {
                hsUniqueSharedString.Add(sHash);
            }

            #region To be deleted once stable
            //Text t;
            //Run r;
            //PhoneticRun phr;
            //int iTextCount = 0;
            //string sText = string.Empty;
            //foreach (var child in Data)
            //{
            //    if (child is Text)
            //    {
            //        t = (Text)child;
            //        sText = t.Text;
            //        ++iTextCount;
            //    }
            //    else if (child is Run)
            //    {
            //        r = (Run)child;
            //        foreach (var runchild in r.ChildElements)
            //        {
            //            if (runchild is Text)
            //            {
            //                t = (Text)runchild;
            //                sText = t.Text;
            //                ++iTextCount;
            //            }
            //        }
            //    }
            //    else if (child is PhoneticRun)
            //    {
            //        phr = (PhoneticRun)child;
            //        foreach (var phrchild in phr.ChildElements)
            //        {
            //            if (phrchild is Text)
            //            {
            //                t = (Text)phrchild;
            //                sText = t.Text;
            //                ++iTextCount;
            //            }
            //        }
            //    }
            //}

            //if (iTextCount == 1)
            //{
            //    // this should track unique text strings, disregarding styles.
            //    // But only if the entire text is formatted, and not partially formatted.
            //    if (!hsUniqueSharedString.Contains(sText))
            //    {
            //        hsUniqueSharedString.Add(sText);
            //    }
            //}
            //else if (iTextCount > 1)
            //{
            //    // if there are multiple text elements, we assume different parts of the text
            //    // are formatted, therefore we just hash the whole internal XML thing
            //    if (!hsUniqueSharedString.Contains(sHash))
            //    {
            //        hsUniqueSharedString.Add(sHash);
            //    }
            //}
            #endregion

            return index;
        }
    }
}
