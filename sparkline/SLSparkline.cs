using System;
using System.Text;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Excel = DocumentFormat.OpenXml.Office.Excel;

namespace SpreadsheetLight
{
    internal class SLSparkline
    {
        internal string WorksheetName;
        internal int StartRowIndex;
        internal int StartColumnIndex;
        internal int EndRowIndex;
        internal int EndColumnIndex;
        internal int LocationRowIndex;
        internal int LocationColumnIndex;

        internal SLSparkline()
        {
            this.WorksheetName = string.Empty;
            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;
            this.LocationRowIndex = 1;
            this.LocationColumnIndex = 1;
        }

        internal X14.Sparkline ToSparkline()
        {
            X14.Sparkline spk = new X14.Sparkline();

            if (this.StartRowIndex == this.EndRowIndex && this.StartColumnIndex == this.EndColumnIndex)
            {
                spk.Formula = new Excel.Formula();
                spk.Formula.Text = SLTool.ToCellReference(this.WorksheetName, this.StartRowIndex, this.StartColumnIndex);
            }
            else
            {
                spk.Formula = new Excel.Formula();
                spk.Formula.Text = SLTool.ToCellRange(this.WorksheetName, this.StartRowIndex, this.StartColumnIndex, this.EndRowIndex, this.EndColumnIndex);
            }

            spk.ReferenceSequence = new Excel.ReferenceSequence();
            spk.ReferenceSequence.Text = SLTool.ToCellReference(this.LocationRowIndex, this.LocationColumnIndex);

            return spk;
        }

        internal SLSparkline Clone()
        {
            SLSparkline spk = new SLSparkline();
            spk.WorksheetName = this.WorksheetName;
            spk.StartRowIndex = this.StartRowIndex;
            spk.StartColumnIndex = this.StartColumnIndex;
            spk.EndRowIndex = this.EndRowIndex;
            spk.EndColumnIndex = this.EndColumnIndex;
            spk.LocationRowIndex = this.LocationRowIndex;
            spk.LocationColumnIndex = this.LocationColumnIndex;

            return spk;
        }
    }
}
