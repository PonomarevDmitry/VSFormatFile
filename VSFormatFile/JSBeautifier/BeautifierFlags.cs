namespace Nav.Common.VSPackages.VSFormatFile.JSBeautifier
{
    using System;

    public class BeautifierFlags
    {
        public BeautifierFlags(string mode)
        {
            PreviousMode = "BLOCK";
            Mode = mode;
            VarLine = false;
            VarLineTainted = false;
            VarLineReindented = false;
            InHtmlComment = false;
            IfLine = false;
            ChainExtraIndentation = 0;
            InCase = false;
            InCaseStatement = false;
            CaseBody = false;
            IndentationLevel = 0;
            TernaryDepth = 0;
        }

        public string PreviousMode { get; set; }

        public string Mode { get; set; }

        public bool VarLine { get; set; }

        public bool VarLineTainted { get; set; }

        public bool VarLineReindented { get; set; }

        public bool InHtmlComment { get; set; }

        public bool IfLine { get; set; }

        public int ChainExtraIndentation { get; set; }

        public bool InCase { get; set; }

        public bool InCaseStatement { get; set; }

        public bool CaseBody { get; set; }

        public int IndentationLevel { get; set; }

        public int TernaryDepth { get; set; }
    }
}