using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Nav.Common.VSPackages.VSFormatFile
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid(PackageGuids.GuidVSFormatFileOptionsString)]
    public class OptionsPage : DialogPage
    {
        public enum LineBreakStyle
        {
            Unix = 0,
            Windows = 1,
        };

        [Category("Format")]
        [Description("Enable remove and sort on save, only apply to .cs file.")]
        public bool EnableRemoveAndSort { get; set; } = true;

        [Category("Format")]
        [Description("Apply remove and sort to .cs without #if. Remove and sort must be enabled first.")]
        public bool EnableSmartRemoveAndSort { get; set; } = true;

        [Category("Format")]
        [Description("Enable format document on save.")]
        public bool EnableFormatDocument { get; set; } = true;

        [Category("Format")]
        [Description("Allow extentions for FormatDocument only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowFormatDocumentExtentions { get; set; } = string.Empty;

        [Category("Format")]
        [Description("Deny extentions for FormatDocument only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyFormatDocumentExtentions { get; set; } = string.Empty;

        [Category("Format")]
        [Description("Enable unify line break on save.")]
        public bool EnableUnifyLineBreak { get; set; } = true;

        [Category("Format")]
        [Description("Allow extentions for all except FormatDocument. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowExtentions { get; set; } = string.Empty;

        [Category("Format")]
        [Description("Deny extentions for all except FormatDocument. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyExtentions { get; set; } = string.Empty;

        [Category("Format")]
        [Description("Line break style.")]
        public LineBreakStyle LineBreak { get; set; } = LineBreakStyle.Windows;

        [Category("Format")]
        [Description("Enable unify end of file to one empty line on save.")]
        public bool EnableUnifyEndOfFile { get; set; } = true;

        [Category("Format")]
        [Description("Enable tab to space on save. Depends on tabs options for the type of file.")]
        public bool EnableTabToSpace { get; set; } = true;

        [Category("Format")]
        [Description("Enable force file encoding to UTF8 without BOM on save.")]
        public bool EnableForceUtf8WithoutBom { get; set; } = true;

        [Category("Format")]
        [Description("Allow extentions for ForceUtf8WithoutBom only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string AllowForceUtf8WithoutBomExtentions { get; set; } = string.Empty;

        [Category("Format")]
        [Description("Deny extentions for ForceUtf8WithoutBom only. Space separated list. For example: .cs .html .cshtml .vb")]
        public string DenyForceUtf8WithoutBomExtentions { get; set; } = string.Empty;

        [Category("Format")]
        [Description("Enable remove trailing spaces. It is mostly for Visual Sutdio 2012, which won't remove trailing spaces when formatting. In higher version than 2012, this will do nothing when FormatDocument is enabled.")]
        public bool EnableRemoveTrailingSpaces { get; set; } = true;

        public AllowDenyDocumentFilter AllowDenyFormatDocumentFilter;
        public AllowDenyDocumentFilter AllowDenyForceUtf8WithoutBomFilter;
        public AllowDenyDocumentFilter AllowDenyFilter;

        void UpdateSettings()
        {
            AllowDenyFormatDocumentFilter = new AllowDenyDocumentFilter(
                AllowFormatDocumentExtentions.Split(' '), DenyFormatDocumentExtentions.Split(' '));

            AllowDenyForceUtf8WithoutBomFilter = new AllowDenyDocumentFilter(
                AllowForceUtf8WithoutBomExtentions.Split(' '), DenyForceUtf8WithoutBomExtentions.Split(' '));

            AllowDenyFilter = new AllowDenyDocumentFilter(
                AllowExtentions.Split(' '), DenyExtentions.Split(' '));
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            base.OnApply(e);
            UpdateSettings();
        }

        public override void LoadSettingsFromStorage()
        {
            base.LoadSettingsFromStorage();
            UpdateSettings();
        }
    }
}
