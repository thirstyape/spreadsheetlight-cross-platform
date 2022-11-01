using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLSheet
    {
        internal string Name { get; set; }
        internal uint SheetId { get; set; }
        internal SheetStateValues State { get; set; }
        internal string Id { get; set; }
        internal SLSheetType SheetType { get; set; }

        internal SLSheet(string Name, uint SheetId, string Id, SLSheetType SheetType)
        {
            this.Name = Name;
            this.SheetId = SheetId;
            this.State = SheetStateValues.Visible;
            this.Id = Id;
            this.SheetType = SheetType;
        }

        internal Sheet ToSheet()
        {
            Sheet s = new Sheet();
            s.Name = this.Name;
            s.SheetId = this.SheetId;
            if (this.State != SheetStateValues.Visible) s.State = this.State;
            s.Id = this.Id;

            return s;
        }

        internal SLSheet Clone()
        {
            SLSheet s = new SLSheet(this.Name, this.SheetId, this.Id, this.SheetType);
            s.State = this.State;
            return s;
        }
    }
}
