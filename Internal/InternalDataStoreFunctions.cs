using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight;

internal static class InternalDataStoreFunctions
{
    internal static int SaveToStylesheet(string hash, SLDocument document)
    {
        if (document.dictStyleHash.TryGetValue(hash, out int index) == false)
        {
            index = document.listStyle.Count;

			document.listStyle.Add(hash);
			document.dictStyleHash[hash] = index;
        }

        return index;
    }

	internal static int ForceSaveToStylesheet(string hash, SLDocument document)
	{
		int index = document.listStyle.Count;

		document.listStyle.Add(hash);
		document.dictStyleHash[hash] = index;

		return index;
	}

    internal static int SaveToStylesheetNumberingFormat(string hash, SLDocument document)
    {
		if (document.dictStyleNumberingFormatHash.TryGetValue(hash, out int index) == false && document.dictBuiltInNumberingFormatHash.TryGetValue(hash, out index) == false)
		{
			index = document.NextNumberFormatId;

			++document.NextNumberFormatId;
			document.dictStyleNumberingFormat[index] = hash;
			document.dictStyleNumberingFormatHash[hash] = index;
		}

		return index;
    }

    internal static int SaveToStylesheetFont(string hash, SLDocument document)
    {
        if (document.dictStyleFontHash.TryGetValue(hash, out int index) == false)
        {
            index = document.listStyleFont.Count;

			document.listStyleFont.Add(hash);
			document.dictStyleFontHash[hash] = index;
        }

        return index;
    }

    internal static int ForceSaveToStylesheetFont(string hash, SLDocument document)
    {
        int index = document.listStyleFont.Count;

		document.listStyleFont.Add(hash);
		document.dictStyleFontHash[hash] = index;

        return index;
    }

    internal static int SaveToStylesheetFill(string hash, SLDocument document)
    {
        if (document.dictStyleFillHash.TryGetValue(hash, out int index) == false)
        {
            index = document.listStyleFill.Count;

			document.listStyleFill.Add(hash);
			document.dictStyleFillHash[hash] = index;
        }

        return index;
    }

    internal static int ForceSaveToStylesheetFill(string hash, SLDocument document)
    {
        int index = document.listStyleFill.Count;

		document.listStyleFill.Add(hash);
		document.dictStyleFillHash[hash] = index;

        return index;
    }

    internal static int SaveToStylesheetBorder(string hash, SLDocument document)
    {
        if (document.dictStyleBorderHash.TryGetValue(hash, out int index) == false)
        {
            index = document.listStyleBorder.Count;

			document.listStyleBorder.Add(hash);
			document.dictStyleBorderHash[hash] = index;
        }

        return index;
    }

    internal static int ForceSaveToStylesheetBorder(string hash, SLDocument document)
    {
        int index = document.listStyleBorder.Count;

		document.listStyleBorder.Add(hash);
		document.dictStyleBorderHash[hash] = index;

        return index;
    }

    internal static int SaveToStylesheetCellStylesFormat(string hash, SLDocument document)
    {
        if (document.dictStyleCellStyleFormatHash.TryGetValue(hash, out int index) == false)
        {
            index = document.listStyleCellStyleFormat.Count;

			document.listStyleCellStyleFormat.Add(hash);
			document.dictStyleCellStyleFormatHash[hash] = index;
        }

        return index;
    }

    internal static int ForceSaveToStylesheetCellStylesFormat(string hash, SLDocument document)
    {
        int index = document.listStyleCellStyleFormat.Count;

		document.listStyleCellStyleFormat.Add(hash);
		document.dictStyleCellStyleFormatHash[hash] = index;

        return index;
    }

    internal static int SaveToStylesheetCellStyle(string hash, SLDocument document)
    {
        if (document.dictStyleCellStyleHash.TryGetValue(hash, out int index) == false)
        {
            index = document.listStyleCellStyle.Count;

			document.listStyleCellStyle.Add(hash);
			document.dictStyleCellStyleHash[hash] = index;
        }

        return index;
    }

    internal static int ForceSaveToStylesheetCellStyle(string hash, SLDocument document)
    {
        int index = document.listStyleCellStyle.Count;

		document.listStyleCellStyle.Add(hash);
		document.dictStyleCellStyleHash[hash] = index;

        return index;
    }

    internal static int SaveToStylesheetDifferentialFormat(string hash, SLDocument document)
    {
        if (document.dictStyleDifferentialFormatHash.TryGetValue(hash, out int index) == false)
        {
            index = document.listStyleDifferentialFormat.Count;

			document.listStyleDifferentialFormat.Add(hash);
			document.dictStyleDifferentialFormatHash[hash] = index;
        }

        return index;
    }

    internal static int ForceSaveToStylesheetDifferentialFormat(string hash, SLDocument document)
    {
        int index = document.listStyleDifferentialFormat.Count;

        document.listStyleDifferentialFormat.Add(hash);
		document.dictStyleDifferentialFormatHash[hash] = index;

        return index;
    }

    internal static void ForceSaveToSharedStringTable(SharedStringItem item, SLDocument document)
    {
        var index = document.listSharedString.Count;
        var hash = SLTool.RemoveNamespaceDeclaration(item.InnerXml);

		document.listSharedString.Add(hash);
		document.dictSharedStringHash[hash] = index;
    }

    internal static int DirectSaveToSharedStringTable(string data, SLDocument document)
    {
        var hash = SLTool.ToPreserveSpace(data) ? string.Format("<x:t xml:space=\"preserve\">{0}</x:t>", data) : string.Format("<x:t>{0}</x:t>", data);

        if (document.dictSharedStringHash.TryGetValue(hash, out int index) == false)
        {
            index = document.listSharedString.Count;

			document.listSharedString.Add(hash);
			document.dictSharedStringHash[hash] = index;
        }

        if (document.hsUniqueSharedString.Contains(data) == false)
			document.hsUniqueSharedString.Add(data);

        return index;
    }

    internal static int DirectSaveToSharedStringTable(InlineString data, SLDocument document)
    {
        var hash = SLTool.RemoveNamespaceDeclaration(data.InnerXml);

        if (document.dictSharedStringHash.TryGetValue(hash, out int index) == false)
        {
            index = document.listSharedString.Count;

			document.listSharedString.Add(hash);
			document.dictSharedStringHash[hash] = index;
        }

        if (document.hsUniqueSharedString.Contains(hash) == false)
			document.hsUniqueSharedString.Add(hash);

        return index;
    }
}
