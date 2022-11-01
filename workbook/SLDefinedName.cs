using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// This simulates the DocumentFormat.OpenXml.Spreadsheet.DefinedName class.
    /// </summary>
    public class SLDefinedName
    {
        /// <summary>
        /// The text of the defined name.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// The name of the defined name. Names starting with "_xlnm" are reserved.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// User comment.
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        /// Custom menu text.
        /// </summary>
        public string CustomMenu { get; set; }

        /// <summary>
        /// Description text.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Help topic for display.
        /// </summary>
        public string Help { get; set; }

        /// <summary>
        /// Status bar text.
        /// </summary>
        public string StatusBar { get; set; }

        /// <summary>
        /// The sheet index (0-based indexing) that's the scope of the defined name. If null, the defined name applies to the entire spreadsheet.
        /// </summary>
        public uint? LocalSheetId { get; set; }

        /// <summary>
        /// Specifies if the defined name is hidden in the user interface. The default value is false.
        /// </summary>
        public bool? Hidden { get; set; }

        /// <summary>
        /// Specifies if the defined name refers to a user-defined function. The default value is false.
        /// </summary>
        public bool? Function { get; set; }

        /// <summary>
        /// Specifies if the defined name is related to an external function, command or executable code. The default value is false.
        /// </summary>
        public bool? VbProcedure { get; set; }

        /// <summary>
        /// Specifies if the defined name is related to an external function, command or executable code. The default value is false.
        /// </summary>
        public bool? Xlm { get; set; }

        /// <summary>
        /// Specifies the function group index if the defined name refers to a function. Refer to Open XML specifications for the meaning of the values. For example, 1 is for "Financial" and 2 is for "Date and Time".
        /// </summary>
        public uint? FunctionGroupId { get; set; }

        /// <summary>
        /// Specifies the keyboard shortcut for the defined name.
        /// </summary>
        public string ShortcutKey { get; set; }

        /// <summary>
        /// Specifies if the defined name is included in a spreadsheet that's published or rendered on a web or application server. The default value is false.
        /// </summary>
        public bool? PublishToServer { get; set; }

        /// <summary>
        /// Specifies that the defined name is used as a parameter of a spreadsheet that's published or rendered on a web or application server. The default value is false.
        /// </summary>
        public bool? WorkbookParameter { get; set; }

        internal SLDefinedName(string Name)
        {
            this.Text = string.Empty;
            this.Name = Name;
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Comment = null;
            this.CustomMenu = null;
            this.Description = null;
            this.Help = null;
            this.StatusBar = null;
            this.LocalSheetId = null;
            this.Hidden = null;
            this.Function = null;
            this.VbProcedure = null;
            this.Xlm = null;
            this.FunctionGroupId = null;
            this.ShortcutKey = null;
            this.PublishToServer = null;
            this.WorkbookParameter = null;
        }

        internal void FromDefinedName(DefinedName dn)
        {
            this.SetAllNull();
            this.Text = dn.Text ?? string.Empty;
            this.Name = dn.Name.Value;
            if (dn.Comment != null) this.Comment = dn.Comment.Value;
            if (dn.CustomMenu != null) this.CustomMenu = dn.CustomMenu.Value;
            if (dn.Description != null) this.Description = dn.Description.Value;
            if (dn.Help != null) this.Help = dn.Help.Value;
            if (dn.StatusBar != null) this.StatusBar = dn.StatusBar.Value;
            if (dn.LocalSheetId != null) this.LocalSheetId = dn.LocalSheetId.Value;
            if (dn.Hidden != null) this.Hidden = dn.Hidden.Value;
            if (dn.Function != null) this.Function = dn.Function.Value;
            if (dn.VbProcedure != null) this.VbProcedure = dn.VbProcedure.Value;
            if (dn.Xlm != null) this.Xlm = dn.Xlm.Value;
            if (dn.FunctionGroupId != null) this.FunctionGroupId = dn.FunctionGroupId.Value;
            if (dn.ShortcutKey != null) this.ShortcutKey = dn.ShortcutKey.Value;
            if (dn.PublishToServer != null) this.PublishToServer = dn.PublishToServer.Value;
            if (dn.WorkbookParameter != null) this.WorkbookParameter = dn.WorkbookParameter.Value;
        }

        internal DefinedName ToDefinedName()
        {
            DefinedName dn = new DefinedName();
            dn.Text = this.Text;
            dn.Name = this.Name;
            if (this.Comment != null) dn.Comment = this.Comment;
            if (this.CustomMenu != null) dn.CustomMenu = this.CustomMenu;
            if (this.Description != null) dn.Description = this.Description;
            if (this.Help != null) dn.Help = this.Help;
            if (this.StatusBar != null) dn.StatusBar = this.StatusBar;
            if (this.LocalSheetId != null) dn.LocalSheetId = this.LocalSheetId.Value;
            if (this.Hidden != null && this.Hidden != false) dn.Hidden = this.Hidden.Value;
            if (this.Function != null && this.Function != false) dn.Function = this.Function.Value;
            if (this.VbProcedure != null && this.VbProcedure != false) dn.VbProcedure = this.VbProcedure.Value;
            if (this.Xlm != null && this.Xlm != false) dn.Xlm = this.Xlm.Value;
            if (this.FunctionGroupId != null) dn.FunctionGroupId = this.FunctionGroupId.Value;
            if (this.ShortcutKey != null) dn.ShortcutKey = this.ShortcutKey;
            if (this.PublishToServer != null && this.PublishToServer != false) dn.PublishToServer = this.PublishToServer.Value;
            if (this.WorkbookParameter != null && this.WorkbookParameter != false) dn.WorkbookParameter = this.WorkbookParameter.Value;

            return dn;
        }

        internal SLDefinedName Clone()
        {
            SLDefinedName dn = new SLDefinedName(this.Name);
            dn.Text = this.Text;
            dn.Name = this.Name;
            dn.Comment = this.Comment;
            dn.CustomMenu = this.CustomMenu;
            dn.Description = this.Description;
            dn.Help = this.Help;
            dn.StatusBar = this.StatusBar;
            dn.LocalSheetId = this.LocalSheetId;
            dn.Hidden = this.Hidden;
            dn.Function = this.Function;
            dn.VbProcedure = this.VbProcedure;
            dn.Xlm = this.Xlm;
            dn.FunctionGroupId = this.FunctionGroupId;
            dn.ShortcutKey = this.ShortcutKey;
            dn.PublishToServer = this.PublishToServer;
            dn.WorkbookParameter = this.WorkbookParameter;

            return dn;
        }
    }
}
