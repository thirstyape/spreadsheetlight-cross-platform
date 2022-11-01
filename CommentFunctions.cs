using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// Get existing comments in the currently selected worksheet. WARNING: This is only a snapshot. Any changes made to the returned result are not used.
        /// </summary>
        /// <returns>A Dictionary of existing comments.</returns>
        public Dictionary<SLCellPoint, SLRstType> GetCommentText()
        {
            Dictionary<SLCellPoint, SLRstType> result = new Dictionary<SLCellPoint, SLRstType>();

            // we don't add to existing comments, so it's either get existing comments
            // or use the newly inserted comments.
            if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
            {
                WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                if (wsp.WorksheetCommentsPart != null)
                {
                    Comment comm;
                    int iRowIndex, iColumnIndex;
                    SLRstType rst = new SLRstType();
                    using (OpenXmlReader oxr = OpenXmlReader.Create(wsp.WorksheetCommentsPart.Comments.CommentList))
                    {
                        while (oxr.Read())
                        {
                            if (oxr.ElementType == typeof(Comment))
                            {
                                comm = (Comment)oxr.LoadCurrentElement();
                                SLTool.FormatCellReferenceToRowColumnIndex(comm.Reference.Value, out iRowIndex, out iColumnIndex);
                                rst.FromCommentText(comm.CommentText);
                                result[new SLCellPoint(iRowIndex, iColumnIndex)] = rst.Clone();
                            }
                        }
                    }
                }
                else
                {
                    List<SLCellPoint> pts = slws.Comments.Keys.ToList<SLCellPoint>();
                    foreach (SLCellPoint pt in pts)
                    {
                        result[pt] = slws.Comments[pt].rst.Clone();
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Insert comment given the cell reference of the cell it's based on. This will overwrite any existing comment.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Comment">The cell comment.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool InsertComment(string CellReference, SLComment Comment)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return InsertComment(iRowIndex, iColumnIndex, Comment);
        }

        /// <summary>
        /// Insert comment given the row index and column index of the cell it's based on. This will overwrite any existing comment.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Comment">The cell comment.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool InsertComment(int RowIndex, int ColumnIndex, SLComment Comment)
        {
            if (!SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                return false;
            }

            if (!slws.Authors.Contains(Comment.Author))
            {
                slws.Authors.Add(Comment.Author);
            }

            SLCellPoint pt = new SLCellPoint(RowIndex, ColumnIndex);
            SLComment comm = Comment.Clone();
            slws.Comments[pt] = comm;

            return true;
        }

        private void WriteCommentPart(WorksheetCommentsPart wcp, VmlDrawingPart vdp)
        {
            List<SLCellPoint> listCommentKeys = slws.Comments.Keys.ToList<SLCellPoint>();
            listCommentKeys.Sort(new SLCellReferencePointComparer());
            bool bAuthorFound = false;
            int iAuthorIndex = 0;
            SLComment comm;

            int i = 0;
            SLRowProperties rp;
            SLColumnProperties cp;
            SLCellPoint pt;

            // just in case
            if (slws.Authors.Count == 0)
            {
                if (this.DocumentProperties.Creator.Length > 0)
                {
                    slws.Authors.Add(this.DocumentProperties.Creator);
                }
                else
                {
                    slws.Authors.Add(SLConstants.ApplicationName);
                }
            }

            int iDataRange = 1;
            // hah! optional... we'll see...
            int iOptionalShapeTypeId = 202;
            string sShapeTypeId = string.Format("_x0000_t{0}", iOptionalShapeTypeId);
            int iShapeIdBase = iDataRange * 1024;

            int iRowID = 0;
            int iColumnID = 0;
            double fRowRemainder = 0;
            double fColumnRemainder = 0;
            long lRowEMU = 0;
            long lColumnEMU = 0;
            long lRowRemainder = 0;
            long lColumnRemainder = 0;
            int iEMU = 0;
            double fMargin = 0;

            double fFrac = 0;
            int iFrac = 0;
            string sFrac = string.Empty;

            ImagePart imgp;
            string sFileName = string.Empty;

            // image data in base 64, relationship ID
            Dictionary<string, string> dictImageData = new Dictionary<string, string>();
            // not supporting existing VML drawings. But if supporting, process the VmlDrawingPart for
            // ImageParts.

            // Apparently, Excel chokes if a "non-standard" relationship ID is given
            // to VML drawings. It seems to only accept the form "rId{num}", and {num} seems
            // to have to start from 1. I don't know if Excel will also choke if you jump
            // from 1 to 3, but I *do* know you can't even start from rId3.
            // Excel VML seems to be particularly strict on this...

            // The error originated by having a "non-standard" relationship ID for a
            // VML image. Say "R2dk723lgsjg2" or whatever. Then fire up Excel and open
            // that spreadsheet. Then save. Then open it again. You'll get an error.
            // Apparently, Excel will put "rId1" on the tag o:relid. The error is that
            // the original o:relid with "R2dk723lgsjg2" as the value is still there.
            // Meaning the o:relid attribute is duplicated, hence the error.

            // Why don't I just use the number of vdp.Parts or even vdp.ImageParts to
            // get the next valid relationship ID? I don't know. Paranoia?
            // The existing relationship IDs *might* be in sequential order, but you never
            // know what Excel accepts... If you can get me the Microsoft Excel developer
            // who can explain this, I'll gladly change the algorithm...

            // So why the dictionary? Apparently, Excel also chokes if there are duplicates of
            // the VML image. So even 2 unique relationship IDs that *happens* to have identical
            // image data will tie Excel into knots. I am so upset with Excel right now...
            // I know it will keep file size down if only unique image data is stored, but still...

            StringBuilder sbVml = new StringBuilder();
            sbVml.Append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
            sbVml.Append("<o:shapelayout v:ext=\"edit\">");
            sbVml.AppendFormat("<o:idmap v:ext=\"edit\" data=\"{0}\"/>", iDataRange);
            sbVml.Append("</o:shapelayout>");

            sbVml.AppendFormat("<v:shapetype id=\"{0}\" coordsize=\"21600,21600\" o:spt=\"{1}\" path=\"m,l,21600r21600,l21600,xe\">", sShapeTypeId, iOptionalShapeTypeId);
            sbVml.Append("<v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>");
            sbVml.Append("</v:shapetype>");

            using (OpenXmlWriter oxwComment = OpenXmlWriter.Create(wcp))
            {
                oxwComment.WriteStartElement(new Comments());

                oxwComment.WriteStartElement(new Authors());
                for (i = 0; i < slws.Authors.Count; ++i)
                {
                    oxwComment.WriteElement(new Author(slws.Authors[i]));
                }
                oxwComment.WriteEndElement();

                oxwComment.WriteStartElement(new CommentList());
                for (i = 0; i < listCommentKeys.Count; ++i)
                {
                    pt = listCommentKeys[i];
                    comm = slws.Comments[pt];

                    bAuthorFound = false;
                    for (iAuthorIndex = 0; iAuthorIndex < slws.Authors.Count; ++iAuthorIndex)
                    {
                        if (comm.Author.Equals(slws.Authors[iAuthorIndex]))
                        {
                            bAuthorFound = true;
                            break;
                        }
                    }
                    if (!bAuthorFound) iAuthorIndex = 0;

                    oxwComment.WriteStartElement(new Comment()
                    {
                        Reference = SLTool.ToCellReference(pt.RowIndex, pt.ColumnIndex),
                        AuthorId = (uint)iAuthorIndex
                    });
                    oxwComment.WriteElement(comm.rst.ToCommentText());
                    oxwComment.WriteEndElement();

                    sbVml.AppendFormat("<v:shape id=\"_x0000_s{0}\" type=\"#{1}\"", iShapeIdBase + i + 1, sShapeTypeId);
                    sbVml.Append(" style='position:absolute;");

                    if (!comm.HasSetPosition)
                    {
                        comm.Top = pt.RowIndex - 1 + SLConstants.DefaultCommentTopOffset;
                        comm.Left = pt.ColumnIndex + SLConstants.DefaultCommentLeftOffset;
                        if (comm.Top < 0) comm.Top = 0;
                        if (comm.Left < 0) comm.Left = 0;
                    }

                    if (comm.UsePositionMargin)
                    {
                        sbVml.AppendFormat("margin-left:{0}pt;", comm.LeftMargin.ToString("0.##", CultureInfo.InvariantCulture));
                        sbVml.AppendFormat("margin-top:{0}pt;", comm.TopMargin.ToString("0.##", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        iRowID = (int)Math.Floor(comm.Top);
                        fRowRemainder = comm.Top - iRowID;
                        iColumnID = (int)Math.Floor(comm.Left);
                        fColumnRemainder = comm.Left - iColumnID;
                        lRowEMU = 0;
                        lColumnEMU = 0;

                        for (iEMU = 1; iEMU <= iRowID; ++iEMU)
                        {
                            if (slws.RowProperties.ContainsKey(iEMU))
                            {
                                rp = slws.RowProperties[iEMU];
                                lRowEMU += rp.HeightInEMU;
                            }
                            else
                            {
                                lRowEMU += slws.SheetFormatProperties.DefaultRowHeightInEMU;
                            }
                        }

                        if (slws.RowProperties.ContainsKey(iRowID + 1))
                        {
                            rp = slws.RowProperties[iRowID + 1];
                            lRowRemainder = Convert.ToInt64(fRowRemainder * rp.HeightInEMU);
                            lRowEMU += lRowRemainder;
                        }
                        else
                        {
                            lRowRemainder = Convert.ToInt64(fRowRemainder * slws.SheetFormatProperties.DefaultRowHeightInEMU);
                            lRowEMU += lRowRemainder;
                        }

                        for (iEMU = 1; iEMU <= iColumnID; ++iEMU)
                        {
                            if (slws.ColumnBreaks.ContainsKey(iEMU))
                            {
                                cp = slws.ColumnProperties[iEMU];
                                lColumnEMU += cp.WidthInEMU;
                            }
                            else
                            {
                                lColumnEMU += slws.SheetFormatProperties.DefaultColumnWidthInEMU;
                            }
                        }

                        if (slws.ColumnProperties.ContainsKey(iColumnID + 1))
                        {
                            cp = slws.ColumnProperties[iColumnID + 1];
                            lColumnRemainder = Convert.ToInt64(fColumnRemainder * cp.WidthInEMU);
                            lColumnEMU += lColumnRemainder;
                        }
                        else
                        {
                            lColumnRemainder = Convert.ToInt64(fColumnRemainder * slws.SheetFormatProperties.DefaultColumnWidthInEMU);
                            lColumnEMU += lColumnRemainder;
                        }

                        fMargin = (double)lColumnEMU / (double)SLConstants.PointToEMU;
                        sbVml.AppendFormat("margin-left:{0}pt;", fMargin.ToString("0.##", CultureInfo.InvariantCulture));
                        fMargin = (double)lRowEMU / (double)SLConstants.PointToEMU;
                        sbVml.AppendFormat("margin-top:{0}pt;", fMargin.ToString("0.##", CultureInfo.InvariantCulture));
                    }

                    if (comm.AutoSize)
                    {
                        sbVml.Append("width:auto;height:auto;");
                    }
                    else
                    {
                        sbVml.AppendFormat("width:{0}pt;", comm.Width.ToString("0.##", CultureInfo.InvariantCulture));
                        sbVml.AppendFormat("height:{0}pt;", comm.Height.ToString("0.##", CultureInfo.InvariantCulture));
                    }

                    sbVml.AppendFormat("z-index:{0};", i + 1);

                    sbVml.AppendFormat("visibility:{0}'", comm.Visible ? "visible" : "hidden");

                    if (!comm.Fill.HasFill)
                    {
                        // use #ffffff ?
                        sbVml.Append(" fillcolor=\"window [65]\"");
                    }
                    else if (comm.Fill.Type == SLA.SLFillType.NoFill)
                    {
                        sbVml.Append(" filled=\"f\"");
                        sbVml.AppendFormat(" fillcolor=\"#{0}{1}{2}\"",
                            comm.Fill.SolidColor.DisplayColor.R.ToString("x2"),
                            comm.Fill.SolidColor.DisplayColor.G.ToString("x2"),
                            comm.Fill.SolidColor.DisplayColor.B.ToString("x2"));
                    }
                    else if (comm.Fill.Type == SLA.SLFillType.SolidFill)
                    {
                        sbVml.AppendFormat(" fillcolor=\"#{0}{1}{2}\"",
                            comm.Fill.SolidColor.DisplayColor.R.ToString("x2"),
                            comm.Fill.SolidColor.DisplayColor.G.ToString("x2"),
                            comm.Fill.SolidColor.DisplayColor.B.ToString("x2"));
                    }
                    else if (comm.Fill.Type == SLA.SLFillType.GradientFill)
                    {
                        if (comm.Fill.GradientColor.GradientStops.Count > 0)
                        {
                            sbVml.AppendFormat(" fillcolor=\"#{0}{1}{2}\"",
                                comm.Fill.GradientColor.GradientStops[0].Color.DisplayColor.R.ToString("x2"),
                                comm.Fill.GradientColor.GradientStops[0].Color.DisplayColor.G.ToString("x2"),
                                comm.Fill.GradientColor.GradientStops[0].Color.DisplayColor.B.ToString("x2"));
                        }
                    }
                    else if (comm.Fill.Type == SLA.SLFillType.BlipFill)
                    {
                        // don't have to do anything
                    }
                    else if (comm.Fill.Type == SLA.SLFillType.PatternFill)
                    {
                        sbVml.AppendFormat(" fillcolor=\"#{0}{1}{2}\"",
                            comm.Fill.PatternForegroundColor.DisplayColor.R.ToString("x2"),
                            comm.Fill.PatternForegroundColor.DisplayColor.G.ToString("x2"),
                            comm.Fill.PatternForegroundColor.DisplayColor.B.ToString("x2"));
                    }
                    
                    if (comm.LineColor != null)
                    {
                        sbVml.AppendFormat(" strokecolor=\"#{0}{1}{2}\"",
                            comm.LineColor.Value.R.ToString("x2"),
                            comm.LineColor.Value.G.ToString("x2"),
                            comm.LineColor.Value.B.ToString("x2"));
                    }

                    if (comm.fLineWeight != null)
                    {
                        sbVml.AppendFormat(" strokeweight=\"{0}pt\"", comm.fLineWeight.Value.ToString("0.##", CultureInfo.InvariantCulture));
                    }

                    sbVml.Append(" o:insetmode=\"auto\">");

                    sbVml.Append("<v:fill");
                    if (comm.Fill.Type == SLA.SLFillType.SolidFill || comm.Fill.Type == SLA.SLFillType.GradientFill)
                    {
                        if (comm.Fill.Type == SLA.SLFillType.SolidFill) fFrac = 100.0 - (double)comm.Fill.SolidColor.Transparency;
                        else fFrac = 100.0 - comm.bFromTransparency;
                        iFrac = Convert.ToInt32(fFrac * 65536.0 / 100.0);
                        if (iFrac <= 0)
                        {
                            sFrac = "0";
                        }
                        else if (iFrac >= 65536)
                        {
                            sFrac = "1";
                        }
                        else
                        {
                            sFrac = string.Format("{0}f", iFrac.ToString(CultureInfo.InvariantCulture));
                        }
                        // default is 1
                        if (!sFrac.Equals("1")) sbVml.AppendFormat(" opacity=\"{0}\"", sFrac);
                    }

                    if (comm.Fill.Type == SLA.SLFillType.SolidFill)
                    {
                        sbVml.AppendFormat(" color2=\"#{0}{1}{2}\"",
                            comm.Fill.SolidColor.DisplayColor.R.ToString("x2"),
                            comm.Fill.SolidColor.DisplayColor.G.ToString("x2"),
                            comm.Fill.SolidColor.DisplayColor.B.ToString("x2"));
                    }
                    else if (comm.Fill.Type == SLA.SLFillType.GradientFill)
                    {
                        if (comm.Fill.GradientColor.GradientStops.Count > 0)
                        {
                            sbVml.AppendFormat(" color2=\"#{0}{1}{2}\"",
                                comm.Fill.GradientColor.GradientStops[0].Color.DisplayColor.R.ToString("x2"),
                                comm.Fill.GradientColor.GradientStops[0].Color.DisplayColor.G.ToString("x2"),
                                comm.Fill.GradientColor.GradientStops[0].Color.DisplayColor.B.ToString("x2"));
                        }
                        else
                        {
                            // shouldn't happen, but you know, in case...
                            sbVml.AppendFormat(" color2=\"#{0}{1}{2}\"",
                                comm.Fill.SolidColor.DisplayColor.R.ToString("x2"),
                                comm.Fill.SolidColor.DisplayColor.G.ToString("x2"),
                                comm.Fill.SolidColor.DisplayColor.B.ToString("x2"));
                        }

                        fFrac = 100.0 - comm.bToTransparency;
                        iFrac = Convert.ToInt32(fFrac * 65536.0 / 100.0);
                        if (iFrac <= 0)
                        {
                            sFrac = "0";
                        }
                        else if (iFrac >= 65536)
                        {
                            sFrac = "1";
                        }
                        else
                        {
                            sFrac = string.Format("{0}f", iFrac.ToString(CultureInfo.InvariantCulture));
                        }
                        // default is 1
                        if (!sFrac.Equals("1")) sbVml.AppendFormat(" o:opacity=\"{0}\"", sFrac);

                        sbVml.Append(" rotate=\"t\"");

                        if (comm.Fill.GradientColor.GradientStops.Count > 0)
                        {
                            sbVml.Append(" colors=\"");
                            for (int iGradient = 0; iGradient < comm.Fill.GradientColor.GradientStops.Count; ++iGradient)
                            {
                                // you take the position/gradient value straight
                                fFrac = (double)comm.Fill.GradientColor.GradientStops[iGradient].Position;
                                iFrac = Convert.ToInt32(fFrac * 65536.0 / 100.0);
                                if (iFrac <= 0)
                                {
                                    sFrac = "0";
                                }
                                else if (iFrac >= 65536)
                                {
                                    sFrac = "1";
                                }
                                else
                                {
                                    sFrac = string.Format("{0}f", iFrac.ToString(CultureInfo.InvariantCulture));
                                }

                                if (iGradient > 0) sbVml.Append(";");
                                sbVml.AppendFormat("{0} #{1}{2}{3}", sFrac,
                                    comm.Fill.GradientColor.GradientStops[iGradient].Color.DisplayColor.R.ToString("x2"),
                                    comm.Fill.GradientColor.GradientStops[iGradient].Color.DisplayColor.G.ToString("x2"),
                                    comm.Fill.GradientColor.GradientStops[iGradient].Color.DisplayColor.B.ToString("x2"));
                            }
                            sbVml.Append("\"");
                        }

                        if (comm.Fill.GradientColor.IsLinear)
                        {
                            // use temporarily
                            // VML increases angles in counter-clockwise direction,
                            // otherwise we'd just use the angle straight from the property
                            //...fFrac = 360.0 - (double)comm.Fill.GradientColor.Angle;
                            fFrac = (double)comm.Fill.GradientColor.Angle;
                            sbVml.AppendFormat(" angle=\"{0}\"", fFrac.ToString("0.##", CultureInfo.InvariantCulture));
                            sbVml.Append(" focus=\"100%\" type=\"gradient\"");
                        }
                        else
                        {
                            switch (comm.Fill.GradientColor.PathType)
                            {
                                case A.PathShadeValues.Shape:
                                    sbVml.Append(" focusposition=\"50%,50%\" focus=\"100%\" type=\"gradientradial\"");
                                    break;
                                case A.PathShadeValues.Rectangle:
                                case A.PathShadeValues.Circle:
                                    // because there's no way to do a circular gradient with VML...
                                    switch (comm.Fill.GradientColor.Direction)
                                    {
                                        case SLA.SLGradientDirectionValues.Center:
                                            sbVml.Append(" focusposition=\"50%,50%\"");
                                            break;
                                        case SLA.SLGradientDirectionValues.CenterToBottomLeftCorner:
                                            // so the "centre" is at the top-right
                                            sbVml.Append(" focusposition=\"100%,0%\"");
                                            break;
                                        case SLA.SLGradientDirectionValues.CenterToBottomRightCorner:
                                            // so the "centre" is at the top-left
                                            sbVml.Append(" focusposition=\"0%,0%\"");
                                            break;
                                        case SLA.SLGradientDirectionValues.CenterToTopLeftCorner:
                                            // so the "centre" is at the bottom-right
                                            sbVml.Append(" focusposition=\"100%,100%\"");
                                            break;
                                        case SLA.SLGradientDirectionValues.CenterToTopRightCorner:
                                            // so the "centre" is at the bottom-left
                                            sbVml.Append(" focusposition=\"0%,100%\"");
                                            break;
                                    }
                                    sbVml.Append(" focus=\"100%\" type=\"gradientradial\"");
                                    break;
                            }
                        }
                    }
                    else if (comm.Fill.Type == SLA.SLFillType.BlipFill)
                    {
                        string sRelId = "rId1";
                        using (FileStream fs = new FileStream(comm.Fill.BlipFileName, FileMode.Open))
                        {
                            byte[] ba = new byte[fs.Length];
                            fs.Read(ba, 0, ba.Length);
                            string sImageData = Convert.ToBase64String(ba);
                            if (dictImageData.ContainsKey(sImageData))
                            {
                                sRelId = dictImageData[sImageData];
                                comm.Fill.BlipRelationshipID = sRelId;
                            }
                            else
                            {
                                // if we haven't found a viable relationship ID by 10 million iterations,
                                // then we have serious issues...
                                for (int iIDNum = 1; iIDNum <= SLConstants.VmlTenMillionIterations; ++iIDNum)
                                {
                                    sRelId = string.Format("rId{0}", iIDNum.ToString(CultureInfo.InvariantCulture));
                                    // we could use a hashset to store the relationship IDs so we
                                    // don't use the ContainsValue() because ContainsValue() is supposedly
                                    // slow... I'm not gonna care because if this algorithm slows enough
                                    // that ContainsValue() is inefficient, that means there are enough VML
                                    // drawings to choke a modestly sized art museum.
                                    if (!dictImageData.ContainsValue(sRelId)) break;
                                }
                                imgp = vdp.AddImagePart(SLA.SLDrawingTool.GetImagePartType(comm.Fill.BlipFileName), sRelId);
                                fs.Position = 0;
                                imgp.FeedData(fs);
                                comm.Fill.BlipRelationshipID = vdp.GetIdOfPart(imgp);

                                dictImageData[sImageData] = sRelId;
                            }
                        }
                        
                        sbVml.AppendFormat(" o:relid=\"{0}\"", comm.Fill.BlipRelationshipID);

                        // all this to get from "myawesomepicture.jpg" to "myawesomepicture"
                        sFileName = comm.Fill.BlipFileName;
                        // use temporarily
                        iFrac = sFileName.LastIndexOfAny("\\/".ToCharArray());
                        sFileName = sFileName.Substring(iFrac + 1);
                        iFrac = sFileName.LastIndexOf(".");
                        sFileName = sFileName.Substring(0, iFrac);
                        sbVml.AppendFormat(" o:title=\"{0}\"", sFileName);

                        sbVml.AppendFormat(" color2=\"#{0}{1}{2}\"",
                            comm.Fill.SolidColor.DisplayColor.R.ToString("x2"),
                            comm.Fill.SolidColor.DisplayColor.G.ToString("x2"),
                            comm.Fill.SolidColor.DisplayColor.B.ToString("x2"));

                        sbVml.Append(" recolor=\"t\" rotate=\"t\"");

                        fFrac = 100.0 - (double)comm.Fill.BlipTransparency;
                        iFrac = Convert.ToInt32(fFrac * 65536.0 / 100.0);
                        if (iFrac <= 0)
                        {
                            sFrac = "0";
                        }
                        else if (iFrac >= 65536)
                        {
                            sFrac = "1";
                        }
                        else
                        {
                            sFrac = string.Format("{0}f", iFrac.ToString(CultureInfo.InvariantCulture));
                        }
                        // default is 1
                        if (!sFrac.Equals("1")) sbVml.AppendFormat(" o:opacity=\"{0}\"", sFrac);

                        if (comm.Fill.BlipTile)
                        {
                            sbVml.Append(" type=\"tile\"");
                            sbVml.AppendFormat(" size=\"{0}%,{1}%\"",
                                comm.Fill.BlipScaleX.ToString("0.##", CultureInfo.InvariantCulture),
                                comm.Fill.BlipScaleY.ToString("0.##", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            sbVml.Append(" type=\"frame\"");
                            // use temporarily
                            //fFrac = (50.0 - (double)comm.Fill.BlipLeftOffset) + (50.0 - (double)comm.Fill.BlipRightOffset);
                            fFrac = 100.0 - (double)comm.Fill.BlipLeftOffset - (double)comm.Fill.BlipRightOffset;
                            sbVml.AppendFormat(" size=\"{0}%,", fFrac.ToString("0.##", CultureInfo.InvariantCulture));
                            fFrac = 100.0 - (double)comm.Fill.BlipTopOffset - (double)comm.Fill.BlipBottomOffset;
                            sbVml.AppendFormat("{0}%\"", fFrac.ToString("0.##",CultureInfo.InvariantCulture));
                        }
                    }
                    else if (comm.Fill.Type == SLA.SLFillType.PatternFill)
                    {
                        string sRelId = "rId1";
                        using (MemoryStream ms = new MemoryStream())
                        {
                            using (System.Drawing.Bitmap bm = SLA.SLDrawingTool.GetVmlPatternFill(comm.Fill.PatternPreset))
                            {
                                bm.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                            }

                            byte[] ba = new byte[ms.Length];
                            ms.Read(ba, 0, ba.Length);
                            string sImageData = Convert.ToBase64String(ba);
                            if (dictImageData.ContainsKey(sImageData))
                            {
                                sRelId = dictImageData[sImageData];
                                comm.Fill.BlipRelationshipID = sRelId;
                            }
                            else
                            {
                                // go check the "normal" image part additions for comments...
                                for (int iIDNum = 1; iIDNum <= SLConstants.VmlTenMillionIterations; ++iIDNum)
                                {
                                    sRelId = string.Format("rId{0}", iIDNum.ToString(CultureInfo.InvariantCulture));
                                    if (!dictImageData.ContainsValue(sRelId)) break;
                                }
                                imgp = vdp.AddImagePart(ImagePartType.Png, sRelId);
                                ms.Position = 0;
                                imgp.FeedData(ms);
                                comm.Fill.BlipRelationshipID = vdp.GetIdOfPart(imgp);

                                dictImageData[sImageData] = sRelId;
                            }
                        }
                        
                        sbVml.AppendFormat(" o:relid=\"{0}\"", comm.Fill.BlipRelationshipID);

                        sbVml.AppendFormat(" o:title=\"{0}\"", SLA.SLDrawingTool.ConvertToVmlTitle(comm.Fill.PatternPreset));

                        sbVml.AppendFormat(" color2=\"#{0}{1}{2}\"",
                            comm.Fill.PatternBackgroundColor.DisplayColor.R.ToString("x2"),
                            comm.Fill.PatternBackgroundColor.DisplayColor.G.ToString("x2"),
                            comm.Fill.PatternBackgroundColor.DisplayColor.B.ToString("x2"));

                        sbVml.Append(" recolor=\"t\" type=\"pattern\"");
                    }
                    sbVml.Append("/>");

                    if (comm.LineStyle != StrokeLineStyleValues.Single || comm.vLineDashStyle != null)
                    {
                        sbVml.Append("<v:stroke");

                        switch (comm.LineStyle)
                        {
                            case StrokeLineStyleValues.Single:
                                // don't have to do anything
                                break;
                            case StrokeLineStyleValues.ThickBetweenThin:
                                sbVml.Append(" linestyle=\"thickBetweenThin\"/>");
                                break;
                            case StrokeLineStyleValues.ThickThin:
                                sbVml.Append(" linestyle=\"thickThin\"/>");
                                break;
                            case StrokeLineStyleValues.ThinThick:
                                sbVml.Append(" linestyle=\"thinThick\"/>");
                                break;
                            case StrokeLineStyleValues.ThinThin:
                                sbVml.Append(" linestyle=\"thinThin\"/>");
                                break;
                        }

                        if (comm.vLineDashStyle != null)
                        {
                            switch (comm.vLineDashStyle.Value)
                            {
                                case SLDashStyleValues.Solid:
                                    sbVml.Append(" dashstyle=\"solid\"/>");
                                    break;
                                case SLDashStyleValues.ShortDash:
                                    sbVml.Append(" dashstyle=\"shortdash\"/>");
                                    break;
                                case SLDashStyleValues.ShortDot:
                                    sbVml.Append(" dashstyle=\"shortdot\"/>");
                                    break;
                                case SLDashStyleValues.ShortDashDot:
                                    sbVml.Append(" dashstyle=\"shortdashdot\"/>");
                                    break;
                                case SLDashStyleValues.ShortDashDotDot:
                                    sbVml.Append(" dashstyle=\"shortdashdotdot\"/>");
                                    break;
                                case SLDashStyleValues.Dot:
                                    sbVml.Append(" dashstyle=\"dot\"/>");
                                    break;
                                case SLDashStyleValues.Dash:
                                    sbVml.Append(" dashstyle=\"dash\"/>");
                                    break;
                                case SLDashStyleValues.LongDash:
                                    sbVml.Append(" dashstyle=\"longdash\"/>");
                                    break;
                                case SLDashStyleValues.DashDot:
                                    sbVml.Append(" dashstyle=\"dashdot\"/>");
                                    break;
                                case SLDashStyleValues.LongDashDot:
                                    sbVml.Append(" dashstyle=\"longdashdot\"/>");
                                    break;
                                case SLDashStyleValues.LongDashDotDot:
                                    sbVml.Append(" dashstyle=\"longdashdotdot\"/>");
                                    break;
                            }
                        }

                        if (comm.vEndCap != null)
                        {
                            switch (comm.vEndCap.Value)
                            {
                                case StrokeEndCapValues.Flat:
                                    sbVml.Append(" endcap=\"flat\"/>");
                                    break;
                                case StrokeEndCapValues.Round:
                                    sbVml.Append(" endcap=\"round\"/>");
                                    break;
                                case StrokeEndCapValues.Square:
                                    sbVml.Append(" endcap=\"square\"/>");
                                    break;
                            }
                        }

                        sbVml.Append("/>");
                    }

                    if (comm.HasShadow)
                    {
                        sbVml.AppendFormat("<v:shadow on=\"t\" color=\"#{0}{1}{2}\" obscured=\"t\"/>",
                            comm.ShadowColor.R.ToString("x2"),
                            comm.ShadowColor.G.ToString("x2"),
                            comm.ShadowColor.B.ToString("x2"));
                    }

                    sbVml.Append("<v:path o:connecttype=\"none\"/>");

                    sbVml.Append("<v:textbox style='mso-direction-alt:auto;");

                    switch (comm.Orientation)
                    {
                        case SLCommentOrientationValues.Horizontal:
                            // don't have to do anything
                            break;
                        case SLCommentOrientationValues.TopDown:
                            sbVml.Append("layout-flow:vertical;mso-layout-flow-alt:top-to-bottom;");
                            break;
                        case SLCommentOrientationValues.Rotated270Degrees:
                            sbVml.Append("layout-flow:vertical;mso-layout-flow-alt:bottom-to-top;");
                            break;
                        case SLCommentOrientationValues.Rotated90Degrees:
                            sbVml.Append("layout-flow:vertical;");
                            break;
                    }

                    if (comm.TextDirection == SLAlignmentReadingOrderValues.RightToLeft)
                    {
                        sbVml.Append("direction:RTL;");
                    }
                    // no else because don't have to do anything

                    if (comm.AutoSize) sbVml.Append("mso-fit-shape-to-text:t;");
                    sbVml.Append("'><div");

                    if (comm.HorizontalTextAlignment != SLHorizontalTextAlignmentValues.Distributed
                        || comm.TextDirection == SLAlignmentReadingOrderValues.RightToLeft)
                    {
                        sbVml.Append(" style='");
                        switch (comm.HorizontalTextAlignment)
                        {
                            case SLHorizontalTextAlignmentValues.Left:
                                sbVml.Append("text-align:left;");
                                break;
                            case SLHorizontalTextAlignmentValues.Justify:
                                sbVml.Append("text-align:justify;");
                                break;
                            case SLHorizontalTextAlignmentValues.Center:
                                sbVml.Append("text-align:center;");
                                break;
                            case SLHorizontalTextAlignmentValues.Right:
                                sbVml.Append("text-align:right;");
                                break;
                            case SLHorizontalTextAlignmentValues.Distributed:
                                // don't have to do anything
                                break;
                        }

                        if (comm.TextDirection == SLAlignmentReadingOrderValues.RightToLeft)
                        {
                            sbVml.Append("direction:rtl;");
                        }
                        sbVml.Append("'");
                    }

                    sbVml.Append("></div>");
                    sbVml.Append("</v:textbox>");

                    sbVml.Append("<x:ClientData ObjectType=\"Note\">");
                    sbVml.Append("<x:MoveWithCells/>");
                    sbVml.Append("<x:SizeWithCells/>");
                    // anchors are bloody hindering awkward inconvenient to calculate...
                    //sbVml.Append("<x:Anchor>");
                    //sbVml.Append("2, 15, 2, 14, 4, 23, 6, 19");
                    //sbVml.Append("</x:Anchor>");
                    sbVml.Append("<x:AutoFill>False</x:AutoFill>");
                    
                    switch (comm.HorizontalTextAlignment)
                    {
                        case SLHorizontalTextAlignmentValues.Left:
                            // don't have to do anything
                            break;
                        case SLHorizontalTextAlignmentValues.Justify:
                            sbVml.Append("<x:TextHAlign>Justify</x:TextHAlign>");
                            break;
                        case SLHorizontalTextAlignmentValues.Center:
                            sbVml.Append("<x:TextHAlign>Center</x:TextHAlign>");
                            break;
                        case SLHorizontalTextAlignmentValues.Right:
                            sbVml.Append("<x:TextHAlign>Right</x:TextHAlign>");
                            break;
                        case SLHorizontalTextAlignmentValues.Distributed:
                            sbVml.Append("<x:TextHAlign>Distributed</x:TextHAlign>");
                            break;
                    }

                    switch (comm.VerticalTextAlignment)
                    {
                        case SLVerticalTextAlignmentValues.Top:
                            // don't have to do anything
                            break;
                        case SLVerticalTextAlignmentValues.Justify:
                            sbVml.Append("<x:TextVAlign>Justify</x:TextVAlign>");
                            break;
                        case SLVerticalTextAlignmentValues.Center:
                            sbVml.Append("<x:TextVAlign>Center</x:TextVAlign>");
                            break;
                        case SLVerticalTextAlignmentValues.Bottom:
                            sbVml.Append("<x:TextVAlign>Bottom</x:TextVAlign>");
                            break;
                        case SLVerticalTextAlignmentValues.Distributed:
                            sbVml.Append("<x:TextVAlign>Distributed</x:TextVAlign>");
                            break;
                    }

                    sbVml.AppendFormat("<x:Row>{0}</x:Row>", pt.RowIndex - 1);
                    sbVml.AppendFormat("<x:Column>{0}</x:Column>", pt.ColumnIndex - 1);
                    if (comm.Visible) sbVml.Append("<x:Visible/>");
                    sbVml.Append("</x:ClientData>");

                    sbVml.Append("</v:shape>");
                }
                oxwComment.WriteEndElement();

                // this is for Comments
                oxwComment.WriteEndElement();
            }

            sbVml.Append("</xml>");

            using (MemoryStream mem = new MemoryStream(Encoding.ASCII.GetBytes(sbVml.ToString())))
            {
                vdp.FeedData(mem);
            }
        }
    }
}
