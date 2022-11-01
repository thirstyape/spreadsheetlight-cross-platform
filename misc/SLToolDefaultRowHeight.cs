using System;

namespace SpreadsheetLight
{
    internal partial class SLTool
    {
        internal static double GetDefaultRowHeight(string FontName)
        {
            // Yes, this is a long list of hard-coded values. The reason is that
            // I don't know how Excel calculates the default row height. I've tried to
            // measure the font ascents and descents and looked at script terms such as
            // cap height and x-height and what-not. Verdana and Gautami are the biggest culprits.
            // Verdana is supposed to fill up the space, being big and all, but it doesn't.
            // The default row height for Verdana is normal.
            // Gautami is kinda small, but it takes up a larger row height than the average typeface.
            // So I can't figure out the formula... There's no consistency at all.

            // So where did I get all these values? I got a list of installed typefaces on my computer.
            // There are 264 of them as of the writing of this comment. Then I used SpreadsheetLight
            // and generated a theme with the installed fonts as the minor font and generated a spreadsheet.
            // So I had 3 folders, dpi96, dpi120 and dpi144. And each folder had 264 Excel files,
            // each named something like Arial.xlsx with Arial as the minor font and Calibri.xlsx
            // with Calibri as the minor font and so on and so forth.

            // Then I changed my screen resolution to 96 DPI and then went into the folder of dpi96.
            // And then opened each Excel file. And then saved it. Excel will automatically calculate
            // the default row height based on the theme's minor font and save that.
            // Then I changed my screen resolution to 120 DPI and did the same thing for dpi120 folder.
            // And I did the similar thing for 144 DPI. I was considering doing it for another DPI,
            // but thought it's not worth it...
            // So I spent the better part of a night opening an Excel file, saving it and closing it.
            // I did that for 792 Excel files. It was... tedious to say the least...

            // Then I wrote a program to read all those 792 files and get the default row height,
            // and write out case statements based on that default row height value.
            // So now you know. There's no secret. Just lots of painstakingly tedious work.

            // What if the parameter is a typeface that's not in the list? We'll use Calibri.
            // Yeah, my list of installed typefaces is going to be the canonical one.

            // EDIT: Office 2013 added a new typeface, Calibri Light. Fortunately, the row height
            // for that is the same as that of Calibri.

            // EDIT: 26 July 2013. Office 2013 snuck in 8 new themes. But the new font Gadugi isn't
            // one of the theme fonts... whatever.

            string fontcheck = string.Empty;
            double fDefaultRowHeight = 15;
            float fResolution = 96;
            // it seems the resolution obtained this way is either 96 or 120,
            // but whatever... I'm including 144 as well.
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(32, 32);
            fResolution = bm.VerticalResolution;
            bm.Dispose();

            //96,120,144
            if (fResolution < 108)
            {
                fDefaultRowHeight = 15;
                fontcheck = string.Format("{0} [96]", FontName);
            }
            else if (fResolution < 132)
            {
                fDefaultRowHeight = 14.4;
                fontcheck = string.Format("{0} [120]", FontName);
            }
            else
            {
                fDefaultRowHeight = 14.5;
                fontcheck = string.Format("{0} [144]", FontName);
            }

            switch (fontcheck)
            {
                case "Agency FB [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Agency FB [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Agency FB [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Aharoni [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Aharoni [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Aharoni [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Algerian [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Algerian [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Algerian [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Andalus [96]":
                    fDefaultRowHeight = 21;
                    break;
                case "Andalus [120]":
                    fDefaultRowHeight = 19.8;
                    break;
                case "Andalus [144]":
                    fDefaultRowHeight = 19;
                    break;
                case "Angsana New [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Angsana New [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Angsana New [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "AngsanaUPC [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "AngsanaUPC [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "AngsanaUPC [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Aparajita [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Aparajita [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "Aparajita [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Arabic Typesetting [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Arabic Typesetting [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "Arabic Typesetting [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Arial [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Arial [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Arial [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Arial Black [96]":
                    fDefaultRowHeight = 18.75;
                    break;
                case "Arial Black [120]":
                    fDefaultRowHeight = 17.4;
                    break;
                case "Arial Black [144]":
                    fDefaultRowHeight = 17;
                    break;
                case "Arial Narrow [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Arial Narrow [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Arial Narrow [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Arial Rounded MT Bold [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Arial Rounded MT Bold [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Arial Rounded MT Bold [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Arial Unicode MS [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Arial Unicode MS [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Arial Unicode MS [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Baskerville Old Face [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Baskerville Old Face [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Baskerville Old Face [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Batang [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Batang [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Batang [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "BatangChe [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "BatangChe [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "BatangChe [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Bauhaus 93 [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Bauhaus 93 [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "Bauhaus 93 [144]":
                    fDefaultRowHeight = 17;
                    break;
                case "Bell MT [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Bell MT [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Bell MT [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Berlin Sans FB [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Berlin Sans FB [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Berlin Sans FB [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Berlin Sans FB Demi [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Berlin Sans FB Demi [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Berlin Sans FB Demi [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Bernard MT Condensed [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Bernard MT Condensed [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Bernard MT Condensed [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Blackadder ITC [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Blackadder ITC [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Blackadder ITC [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Bodoni MT [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Bodoni MT [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Bodoni MT [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Bodoni MT Black [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Bodoni MT Black [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Bodoni MT Black [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Bodoni MT Condensed [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Bodoni MT Condensed [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Bodoni MT Condensed [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Bodoni MT Poster Compressed [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Bodoni MT Poster Compressed [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Bodoni MT Poster Compressed [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Book Antiqua [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Book Antiqua [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Book Antiqua [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Bookman Old Style [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Bookman Old Style [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Bookman Old Style [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Bookshelf Symbol 7 [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Bookshelf Symbol 7 [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "Bookshelf Symbol 7 [144]":
                    fDefaultRowHeight = 13;
                    break;
                case "Bradley Hand ITC [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Bradley Hand ITC [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Bradley Hand ITC [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Britannic Bold [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Britannic Bold [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Britannic Bold [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Broadway [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Broadway [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Broadway [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Browallia New [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Browallia New [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Browallia New [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "BrowalliaUPC [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "BrowalliaUPC [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "BrowalliaUPC [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Brush Script MT [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Brush Script MT [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Brush Script MT [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Calibri [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Calibri [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Calibri [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Calibri Light [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Calibri Light [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Calibri Light [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Californian FB [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Californian FB [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Californian FB [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Calisto MT [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Calisto MT [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Calisto MT [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Cambria [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Cambria [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Cambria [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Cambria Math [96]":
                    fDefaultRowHeight = 87.75;
                    break;
                case "Cambria Math [120]":
                    fDefaultRowHeight = 83.4;
                    break;
                case "Cambria Math [144]":
                    fDefaultRowHeight = 85.5;
                    break;
                case "Candara [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Candara [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Candara [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Castellar [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Castellar [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Castellar [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Centaur [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Centaur [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Centaur [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Century [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Century [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Century [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Century Gothic [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Century Gothic [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Century Gothic [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Century Schoolbook [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Century Schoolbook [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Century Schoolbook [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Chiller [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Chiller [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Chiller [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Colonna MT [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Colonna MT [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Colonna MT [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Comic Sans MS [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Comic Sans MS [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Comic Sans MS [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Consolas [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Consolas [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Consolas [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Constantia [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Constantia [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Constantia [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Cooper Black [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Cooper Black [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Cooper Black [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Copperplate Gothic Bold [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Copperplate Gothic Bold [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Copperplate Gothic Bold [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Copperplate Gothic Light [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Copperplate Gothic Light [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Copperplate Gothic Light [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Corbel [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Corbel [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Corbel [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Cordia New [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Cordia New [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Cordia New [144]":
                    fDefaultRowHeight = 17;
                    break;
                case "CordiaUPC [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "CordiaUPC [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "CordiaUPC [144]":
                    fDefaultRowHeight = 17;
                    break;
                case "Courier New [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Courier New [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Courier New [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Curlz MT [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Curlz MT [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Curlz MT [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "DaunPenh [96]":
                    fDefaultRowHeight = 19.5;
                    break;
                case "DaunPenh [120]":
                    fDefaultRowHeight = 18.6;
                    break;
                case "DaunPenh [144]":
                    fDefaultRowHeight = 19;
                    break;
                case "David [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "David [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "David [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "DejaVu Sans [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "DejaVu Sans [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "DejaVu Sans [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "DejaVu Sans Condensed [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "DejaVu Sans Condensed [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "DejaVu Sans Condensed [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "DejaVu Sans Light [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "DejaVu Sans Light [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "DejaVu Sans Light [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "DejaVu Sans Mono [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "DejaVu Sans Mono [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "DejaVu Sans Mono [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "DejaVu Serif [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "DejaVu Serif [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "DejaVu Serif [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "DejaVu Serif Condensed [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "DejaVu Serif Condensed [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "DejaVu Serif Condensed [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "DFKai-SB [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "DFKai-SB [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "DFKai-SB [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "DilleniaUPC [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "DilleniaUPC [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "DilleniaUPC [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "DokChampa [96]":
                    fDefaultRowHeight = 27;
                    break;
                case "DokChampa [120]":
                    fDefaultRowHeight = 25.8;
                    break;
                case "DokChampa [144]":
                    fDefaultRowHeight = 26.5;
                    break;
                case "Dotum [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Dotum [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Dotum [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "DotumChe [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "DotumChe [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "DotumChe [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Ebrima [96]":
                    fDefaultRowHeight = 20.25;
                    break;
                case "Ebrima [120]":
                    fDefaultRowHeight = 19.2;
                    break;
                case "Ebrima [144]":
                    fDefaultRowHeight = 19;
                    break;
                case "Edwardian Script ITC [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Edwardian Script ITC [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Edwardian Script ITC [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Elephant [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Elephant [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Elephant [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Engravers MT [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Engravers MT [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Engravers MT [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Eras Bold ITC [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Eras Bold ITC [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Eras Bold ITC [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Eras Demi ITC [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Eras Demi ITC [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Eras Demi ITC [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Eras Light ITC [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Eras Light ITC [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Eras Light ITC [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Eras Medium ITC [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Eras Medium ITC [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Eras Medium ITC [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Estrangelo Edessa [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Estrangelo Edessa [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Estrangelo Edessa [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "EucrosiaUPC [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "EucrosiaUPC [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "EucrosiaUPC [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Euphemia [96]":
                    fDefaultRowHeight = 18.75;
                    break;
                case "Euphemia [120]":
                    fDefaultRowHeight = 17.4;
                    break;
                case "Euphemia [144]":
                    fDefaultRowHeight = 17;
                    break;
                case "FangSong [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "FangSong [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "FangSong [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Felix Titling [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Felix Titling [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Felix Titling [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Footlight MT Light [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Footlight MT Light [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Footlight MT Light [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Forte [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Forte [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "Forte [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Franklin Gothic Book [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Franklin Gothic Book [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Book [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Demi [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Franklin Gothic Demi [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Demi [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Demi Cond [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Franklin Gothic Demi Cond [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Demi Cond [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Heavy [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Franklin Gothic Heavy [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Heavy [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Medium [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Franklin Gothic Medium [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Medium [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Medium Cond [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Franklin Gothic Medium Cond [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Franklin Gothic Medium Cond [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "FrankRuehl [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "FrankRuehl [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "FrankRuehl [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "FreesiaUPC [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "FreesiaUPC [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "FreesiaUPC [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Freestyle Script [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Freestyle Script [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Freestyle Script [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "French Script MT [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "French Script MT [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "French Script MT [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Gabriola [96]":
                    fDefaultRowHeight = 24;
                    break;
                case "Gabriola [120]":
                    fDefaultRowHeight = 22.2;
                    break;
                case "Gabriola [144]":
                    fDefaultRowHeight = 22.5;
                    break;
                case "Gadugi [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Gadugi [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Gadugi [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Garamond [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Garamond [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Garamond [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Gautami [96]":
                    fDefaultRowHeight = 25.5;
                    break;
                case "Gautami [120]":
                    fDefaultRowHeight = 25.2;
                    break;
                case "Gautami [144]":
                    fDefaultRowHeight = 25.5;
                    break;
                case "Gentium Basic [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Gentium Basic [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Gentium Basic [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Gentium Book Basic [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Gentium Book Basic [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Gentium Book Basic [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Georgia [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Georgia [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Georgia [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Gigi [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Gigi [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "Gigi [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Gill Sans MT [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Gill Sans MT [120]":
                    fDefaultRowHeight = 18;
                    break;
                case "Gill Sans MT [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Gill Sans MT Condensed [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Gill Sans MT Condensed [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Gill Sans MT Condensed [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Gill Sans MT Ext Condensed Bold [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Gill Sans MT Ext Condensed Bold [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Gill Sans MT Ext Condensed Bold [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Gill Sans Ultra Bold [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Gill Sans Ultra Bold [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Gill Sans Ultra Bold [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Gill Sans Ultra Bold Condensed [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Gill Sans Ultra Bold Condensed [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Gill Sans Ultra Bold Condensed [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Gisha [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Gisha [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Gisha [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Gloucester MT Extra Condensed [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Gloucester MT Extra Condensed [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Gloucester MT Extra Condensed [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Goudy Old Style [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Goudy Old Style [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Goudy Old Style [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Goudy Stout [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Goudy Stout [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Goudy Stout [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Gulim [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Gulim [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Gulim [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "GulimChe [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "GulimChe [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "GulimChe [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Gungsuh [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Gungsuh [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Gungsuh [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "GungsuhChe [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "GungsuhChe [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "GungsuhChe [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Haettenschweiler [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Haettenschweiler [120]":
                    fDefaultRowHeight = 12.6;
                    break;
                case "Haettenschweiler [144]":
                    fDefaultRowHeight = 13;
                    break;
                case "Harlow Solid Italic [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Harlow Solid Italic [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Harlow Solid Italic [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Harrington [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Harrington [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Harrington [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "High Tower Text [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "High Tower Text [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "High Tower Text [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Impact [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Impact [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Impact [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Imprint MT Shadow [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Imprint MT Shadow [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Imprint MT Shadow [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Informal Roman [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Informal Roman [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Informal Roman [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "IrisUPC [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "IrisUPC [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "IrisUPC [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Iskoola Pota [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Iskoola Pota [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Iskoola Pota [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "JasmineUPC [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "JasmineUPC [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "JasmineUPC [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Jokerman [96]":
                    fDefaultRowHeight = 18.75;
                    break;
                case "Jokerman [120]":
                    fDefaultRowHeight = 17.4;
                    break;
                case "Jokerman [144]":
                    fDefaultRowHeight = 18.5;
                    break;
                case "Juice ITC [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Juice ITC [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Juice ITC [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "KaiTi [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "KaiTi [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "KaiTi [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Kalinga [96]":
                    fDefaultRowHeight = 18;
                    break;
                case "Kalinga [120]":
                    fDefaultRowHeight = 18;
                    break;
                case "Kalinga [144]":
                    fDefaultRowHeight = 17.5;
                    break;
                case "Kartika [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Kartika [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "Kartika [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Khmer UI [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Khmer UI [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Khmer UI [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "KodchiangUPC [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "KodchiangUPC [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "KodchiangUPC [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Kokila [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Kokila [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Kokila [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Kristen ITC [96]":
                    fDefaultRowHeight = 18;
                    break;
                case "Kristen ITC [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Kristen ITC [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Kunstler Script [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Kunstler Script [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Kunstler Script [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Lao UI [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Lao UI [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Lao UI [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Latha [96]":
                    fDefaultRowHeight = 22.5;
                    break;
                case "Latha [120]":
                    fDefaultRowHeight = 21;
                    break;
                case "Latha [144]":
                    fDefaultRowHeight = 21.5;
                    break;
                case "Leelawadee [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Leelawadee [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Leelawadee [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Levenim MT [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Levenim MT [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Levenim MT [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Liberation Mono [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Liberation Mono [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Liberation Mono [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Liberation Sans [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Liberation Sans [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Liberation Sans [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Liberation Sans Narrow [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Liberation Sans Narrow [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Liberation Sans Narrow [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Liberation Serif [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Liberation Serif [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Liberation Serif [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "LilyUPC [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "LilyUPC [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "LilyUPC [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Linux Biolinum G [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Linux Biolinum G [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Linux Biolinum G [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Linux Libertine G [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Linux Libertine G [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Linux Libertine G [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Lucida Bright [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Lucida Bright [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Lucida Bright [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Lucida Calligraphy [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Lucida Calligraphy [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Lucida Calligraphy [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Lucida Console [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Lucida Console [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Lucida Console [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Lucida Fax [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Lucida Fax [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Lucida Fax [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Lucida Handwriting [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Lucida Handwriting [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Lucida Handwriting [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Lucida Sans [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Lucida Sans [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Lucida Sans [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Lucida Sans Typewriter [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Lucida Sans Typewriter [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Lucida Sans Typewriter [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Lucida Sans Unicode [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Lucida Sans Unicode [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Lucida Sans Unicode [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Magneto [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Magneto [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Magneto [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Maiandra GD [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Maiandra GD [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Maiandra GD [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Malgun Gothic [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Malgun Gothic [120]":
                    fDefaultRowHeight = 17.4;
                    break;
                case "Malgun Gothic [144]":
                    fDefaultRowHeight = 17;
                    break;
                case "Mangal [96]":
                    fDefaultRowHeight = 25.5;
                    break;
                case "Mangal [120]":
                    fDefaultRowHeight = 24;
                    break;
                case "Mangal [144]":
                    fDefaultRowHeight = 24;
                    break;
                case "Marlett [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Marlett [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Marlett [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Matura MT Script Capitals [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Matura MT Script Capitals [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Matura MT Script Capitals [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Meiryo [96]":
                    fDefaultRowHeight = 18.75;
                    break;
                case "Meiryo [120]":
                    fDefaultRowHeight = 17.4;
                    break;
                case "Meiryo [144]":
                    fDefaultRowHeight = 17.5;
                    break;
                case "Meiryo UI [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Meiryo UI [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Meiryo UI [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Microsoft Himalaya [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Microsoft Himalaya [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Microsoft Himalaya [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Microsoft JhengHei [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Microsoft JhengHei [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Microsoft JhengHei [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Microsoft New Tai Lue [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Microsoft New Tai Lue [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Microsoft New Tai Lue [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Microsoft PhagsPa [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Microsoft PhagsPa [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Microsoft PhagsPa [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Microsoft Sans Serif [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Microsoft Sans Serif [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Microsoft Sans Serif [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Microsoft Tai Le [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Microsoft Tai Le [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Microsoft Tai Le [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Microsoft Uighur [96]":
                    fDefaultRowHeight = 18;
                    break;
                case "Microsoft Uighur [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Microsoft Uighur [144]":
                    fDefaultRowHeight = 17;
                    break;
                case "Microsoft YaHei [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Microsoft YaHei [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Microsoft YaHei [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Microsoft Yi Baiti [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Microsoft Yi Baiti [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "Microsoft Yi Baiti [144]":
                    fDefaultRowHeight = 13;
                    break;
                case "MingLiU [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "MingLiU [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "MingLiU [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "MingLiU-ExtB [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "MingLiU-ExtB [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "MingLiU-ExtB [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "MingLiU_HKSCS [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "MingLiU_HKSCS [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "MingLiU_HKSCS [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "MingLiU_HKSCS-ExtB [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "MingLiU_HKSCS-ExtB [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "MingLiU_HKSCS-ExtB [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Miriam [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Miriam [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Miriam [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Miriam Fixed [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Miriam Fixed [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Miriam Fixed [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Mistral [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Mistral [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Mistral [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Modern No. 20 [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Modern No. 20 [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Modern No. 20 [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Mongolian Baiti [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Mongolian Baiti [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Mongolian Baiti [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Monotype Corsiva [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Monotype Corsiva [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Monotype Corsiva [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "MoolBoran [96]":
                    fDefaultRowHeight = 19.5;
                    break;
                case "MoolBoran [120]":
                    fDefaultRowHeight = 18.6;
                    break;
                case "MoolBoran [144]":
                    fDefaultRowHeight = 19;
                    break;
                case "MS Gothic [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "MS Gothic [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "MS Gothic [144]":
                    fDefaultRowHeight = 13;
                    break;
                case "MS Mincho [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "MS Mincho [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "MS Mincho [144]":
                    fDefaultRowHeight = 13;
                    break;
                case "MS Outlook [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "MS Outlook [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "MS Outlook [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "MS PGothic [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "MS PGothic [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "MS PGothic [144]":
                    fDefaultRowHeight = 13;
                    break;
                case "MS PMincho [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "MS PMincho [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "MS PMincho [144]":
                    fDefaultRowHeight = 13;
                    break;
                case "MS Reference Sans Serif [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "MS Reference Sans Serif [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "MS Reference Sans Serif [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "MS Reference Specialty [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "MS Reference Specialty [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "MS Reference Specialty [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "MS UI Gothic [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "MS UI Gothic [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "MS UI Gothic [144]":
                    fDefaultRowHeight = 13;
                    break;
                case "MT Extra [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "MT Extra [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "MT Extra [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "MV Boli [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "MV Boli [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "MV Boli [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Narkisim [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Narkisim [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Narkisim [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Niagara Engraved [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Niagara Engraved [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Niagara Engraved [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Niagara Solid [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Niagara Solid [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Niagara Solid [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "NSimSun [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "NSimSun [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "NSimSun [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Nyala [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Nyala [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Nyala [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "OCR A Extended [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "OCR A Extended [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "OCR A Extended [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Old English Text MT [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Old English Text MT [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Old English Text MT [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Onyx [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Onyx [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "Onyx [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "OpenSymbol [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "OpenSymbol [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "OpenSymbol [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Palace Script MT [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Palace Script MT [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Palace Script MT [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Palatino Linotype [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Palatino Linotype [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Palatino Linotype [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Papyrus [96]":
                    fDefaultRowHeight = 19.5;
                    break;
                case "Papyrus [120]":
                    fDefaultRowHeight = 18;
                    break;
                case "Papyrus [144]":
                    fDefaultRowHeight = 18;
                    break;
                case "Parchment [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Parchment [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Parchment [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Perpetua [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Perpetua [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Perpetua [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Perpetua Titling MT [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Perpetua Titling MT [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Perpetua Titling MT [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Plantagenet Cherokee [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Plantagenet Cherokee [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Plantagenet Cherokee [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Playbill [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Playbill [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "Playbill [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "PMingLiU [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "PMingLiU [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "PMingLiU [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "PMingLiU-ExtB [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "PMingLiU-ExtB [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "PMingLiU-ExtB [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Poor Richard [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Poor Richard [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Poor Richard [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Pristina [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Pristina [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "Pristina [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Raavi [96]":
                    fDefaultRowHeight = 22.5;
                    break;
                case "Raavi [120]":
                    fDefaultRowHeight = 21.6;
                    break;
                case "Raavi [144]":
                    fDefaultRowHeight = 21;
                    break;
                case "Rage Italic [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Rage Italic [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Rage Italic [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Ravie [96]":
                    fDefaultRowHeight = 18;
                    break;
                case "Ravie [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Ravie [144]":
                    fDefaultRowHeight = 17;
                    break;
                case "Rockwell [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Rockwell [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Rockwell [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Rockwell Condensed [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Rockwell Condensed [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Rockwell Condensed [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Rockwell Extra Bold [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Rockwell Extra Bold [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Rockwell Extra Bold [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Rod [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Rod [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Rod [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Sakkal Majalla [96]":
                    fDefaultRowHeight = 18;
                    break;
                case "Sakkal Majalla [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Sakkal Majalla [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Script MT Bold [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Script MT Bold [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Script MT Bold [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Segoe Print [96]":
                    fDefaultRowHeight = 23.25;
                    break;
                case "Segoe Print [120]":
                    fDefaultRowHeight = 21.6;
                    break;
                case "Segoe Print [144]":
                    fDefaultRowHeight = 22;
                    break;
                case "Segoe Script [96]":
                    fDefaultRowHeight = 18.75;
                    break;
                case "Segoe Script [120]":
                    fDefaultRowHeight = 19.2;
                    break;
                case "Segoe Script [144]":
                    fDefaultRowHeight = 19;
                    break;
                case "Segoe UI [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Segoe UI [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Segoe UI [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Segoe UI Light [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Segoe UI Light [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Segoe UI Light [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Segoe UI Semibold [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Segoe UI Semibold [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Segoe UI Semibold [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Segoe UI Symbol [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Segoe UI Symbol [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Segoe UI Symbol [144]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Shonar Bangla [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Shonar Bangla [120]":
                    fDefaultRowHeight = 16.8;
                    break;
                case "Shonar Bangla [144]":
                    fDefaultRowHeight = 17;
                    break;
                case "Showcard Gothic [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Showcard Gothic [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Showcard Gothic [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Shruti [96]":
                    fDefaultRowHeight = 21;
                    break;
                case "Shruti [120]":
                    fDefaultRowHeight = 19.8;
                    break;
                case "Shruti [144]":
                    fDefaultRowHeight = 19;
                    break;
                case "SimHei [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "SimHei [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "SimHei [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Simplified Arabic [96]":
                    fDefaultRowHeight = 23.25;
                    break;
                case "Simplified Arabic [120]":
                    fDefaultRowHeight = 21.6;
                    break;
                case "Simplified Arabic [144]":
                    fDefaultRowHeight = 22;
                    break;
                case "Simplified Arabic Fixed [96]":
                    fDefaultRowHeight = 17.25;
                    break;
                case "Simplified Arabic Fixed [120]":
                    fDefaultRowHeight = 16.2;
                    break;
                case "Simplified Arabic Fixed [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "SimSun [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "SimSun [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "SimSun [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "SimSun-ExtB [96]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "SimSun-ExtB [120]":
                    fDefaultRowHeight = 13.2;
                    break;
                case "SimSun-ExtB [144]":
                    fDefaultRowHeight = 13;
                    break;
                case "Snap ITC [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Snap ITC [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Snap ITC [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Stencil [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Stencil [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Stencil [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Sylfaen [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Sylfaen [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Sylfaen [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Symbol [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Symbol [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Symbol [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Tahoma [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Tahoma [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Tahoma [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Tempus Sans ITC [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Tempus Sans ITC [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Tempus Sans ITC [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Times New Roman [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Times New Roman [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Times New Roman [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Traditional Arabic [96]":
                    fDefaultRowHeight = 22.5;
                    break;
                case "Traditional Arabic [120]":
                    fDefaultRowHeight = 19.2;
                    break;
                case "Traditional Arabic [144]":
                    fDefaultRowHeight = 19;
                    break;
                case "Trebuchet MS [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Trebuchet MS [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Trebuchet MS [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Tunga [96]":
                    fDefaultRowHeight = 22.5;
                    break;
                case "Tunga [120]":
                    fDefaultRowHeight = 21;
                    break;
                case "Tunga [144]":
                    fDefaultRowHeight = 21;
                    break;
                case "Tw Cen MT [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Tw Cen MT [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Tw Cen MT [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Tw Cen MT Condensed [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Tw Cen MT Condensed [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Tw Cen MT Condensed [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Tw Cen MT Condensed Extra Bold [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Tw Cen MT Condensed Extra Bold [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Tw Cen MT Condensed Extra Bold [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Utsaah [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Utsaah [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Utsaah [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Vani [96]":
                    fDefaultRowHeight = 21.75;
                    break;
                case "Vani [120]":
                    fDefaultRowHeight = 20.4;
                    break;
                case "Vani [144]":
                    fDefaultRowHeight = 20.5;
                    break;
                case "Verdana [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Verdana [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Verdana [144]":
                    fDefaultRowHeight = 13.5;
                    break;
                case "Vijaya [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Vijaya [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Vijaya [144]":
                    fDefaultRowHeight = 15.5;
                    break;
                case "Viner Hand ITC [96]":
                    fDefaultRowHeight = 18.75;
                    break;
                case "Viner Hand ITC [120]":
                    fDefaultRowHeight = 18;
                    break;
                case "Viner Hand ITC [144]":
                    fDefaultRowHeight = 18.5;
                    break;
                case "Vivaldi [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Vivaldi [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Vivaldi [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Vladimir Script [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Vladimir Script [120]":
                    fDefaultRowHeight = 15;
                    break;
                case "Vladimir Script [144]":
                    fDefaultRowHeight = 15;
                    break;
                case "Vrinda [96]":
                    fDefaultRowHeight = 16.5;
                    break;
                case "Vrinda [120]":
                    fDefaultRowHeight = 15.6;
                    break;
                case "Vrinda [144]":
                    fDefaultRowHeight = 16;
                    break;
                case "Webdings [96]":
                    fDefaultRowHeight = 15.75;
                    break;
                case "Webdings [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Webdings [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Wide Latin [96]":
                    fDefaultRowHeight = 15;
                    break;
                case "Wide Latin [120]":
                    fDefaultRowHeight = 14.4;
                    break;
                case "Wide Latin [144]":
                    fDefaultRowHeight = 14.5;
                    break;
                case "Wingdings [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Wingdings [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Wingdings [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Wingdings 2 [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Wingdings 2 [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Wingdings 2 [144]":
                    fDefaultRowHeight = 14;
                    break;
                case "Wingdings 3 [96]":
                    fDefaultRowHeight = 14.25;
                    break;
                case "Wingdings 3 [120]":
                    fDefaultRowHeight = 13.8;
                    break;
                case "Wingdings 3 [144]":
                    fDefaultRowHeight = 14;
                    break;
            }

            return fDefaultRowHeight;
        }
    }
}
