namespace GovDocFormatter;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static int Main(string[] args)
    {
        if (args.Length >= 3 && string.Equals(args[0], "--format", StringComparison.OrdinalIgnoreCase))
        {
            var options = new FormattingOptions
            {
                InsertFrontMatter = false,
                InsertEndingMatter = false,
                ApplyPageNumbers = true,
                FormatExistingContent = true,
                AutoDetectTitle = true
            };

            new GovDocumentFormatter().Format(args[1], args[2], options);
            return 0;
        }

        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
        return 0;
    }    
}
