using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight;

internal static class SharedStringFunctions
{
	internal static List<SLRstType> GetSharedStrings(SLDocument document)
	{
		var result = new List<SLRstType>();
		var rst = new SLRstType();

		for (int i = 0; i < document.listSharedString.Count; ++i)
		{
			rst.FromHash(document.listSharedString[i]);
			result.Add(rst.Clone());
		}

		return result;
	}

	internal static List<SharedStringItem> GetSharedStringItems(SLDocument document)
	{
		var result = new List<SharedStringItem>();

		for (int i = 0; i < document.listSharedString.Count; ++i)
		{
			var ssi = new SharedStringItem
			{
				InnerXml = document.listSharedString[i]
			};

			result.Add(ssi);
		}

		return result;
	}

	internal static void LoadSharedStringTable(SLDocument document)
	{
		document.countSharedString = 0;
		document.listSharedString = new();
		document.dictSharedStringHash = new();
		document.hsUniqueSharedString = new();

		if (document.wbp.SharedStringTablePart == null)
			return;

		var oxr = OpenXmlReader.Create(document.wbp.SharedStringTablePart);

		while (oxr.Read())
		{
			if (oxr.ElementType == typeof(SharedStringItem))
				InternalDataStoreFunctions.ForceSaveToSharedStringTable((SharedStringItem)oxr.LoadCurrentElement(), document);
		}

		oxr.Dispose();
		document.countSharedString = document.listSharedString.Count;
	}

	internal static void WriteSharedStringTable(SLDocument document)
	{
		if (document.wbp.SharedStringTablePart != null)
		{
			if (document.listSharedString.Count <= document.countSharedString)
				return;

			if (document.WriteUniqueSharedStringCount)
			{
				document.wbp.SharedStringTablePart.SharedStringTable.Count = (uint)document.listSharedString.Count;
				document.wbp.SharedStringTablePart.SharedStringTable.UniqueCount = (uint)document.hsUniqueSharedString.Count;
			}
			else
			{
				document.wbp.SharedStringTablePart.SharedStringTable.Count = null;
				document.wbp.SharedStringTablePart.SharedStringTable.UniqueCount = null;
			}

			var diff = document.listSharedString.Count - document.countSharedString;

			for (int i = 0; i < diff; ++i)
			{
				document.wbp.SharedStringTablePart.SharedStringTable.Append(new SharedStringItem()
				{
					InnerXml = document.listSharedString[i + document.countSharedString]
				});
			}

			document.wbp.SharedStringTablePart.SharedStringTable.Save();
		}
		else if (document.listSharedString.Count > 0)
		{
			var sstp = document.wbp.AddNewPart<SharedStringTablePart>();

			using var ms = new MemoryStream();
			using var sw = new StreamWriter(ms);

			if (document.WriteUniqueSharedStringCount)
			{
				sw.Write("<x:sst count=\"{0}\" uniqueCount=\"{1}\" xmlns:x=\"{2}\">", document.listSharedString.Count, document.hsUniqueSharedString.Count, SLConstants.NamespaceX);
			}
			else
			{
				sw.Write("<x:sst xmlns:x=\"{0}\">", SLConstants.NamespaceX);
			}

			for (int i = 0; i < document.listSharedString.Count; ++i)
			{
				sw.Write("<x:si>{0}</x:si>", document.listSharedString[i]);
			}

			sw.Write("</x:sst>");
			sw.Flush();
			ms.Position = 0;
			sstp.FeedData(ms);
		}
	}
}
