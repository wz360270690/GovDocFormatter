using GovDocFormatter;

var appVersion = typeof(GovDocumentFormatter).Assembly.GetName().Version?.ToString(3) ?? "1.0.12";

if (args.Length == 0 || HasFlag(args, "--help") || HasFlag(args, "-h"))
{
    PrintUsage(appVersion);
    return 0;
}

if (HasFlag(args, "--version") || HasFlag(args, "-v"))
{
    Console.WriteLine("GovDocFormatter Linux CLI v" + appVersion);
    return 0;
}

if (args.Length < 2)
{
    Console.Error.WriteLine("缺少输入或输出文件路径。");
    PrintUsage(appVersion);
    return 2;
}

var inputPath = args[0];
var outputPath = args[1];

var options = new FormattingOptions
{
    InsertFrontMatter = false,
    InsertEndingMatter = false,
    ApplyPageNumbers = !HasFlag(args, "--no-page-numbers"),
    FormatExistingContent = !HasFlag(args, "--no-format-existing"),
    AutoDetectTitle = !HasFlag(args, "--no-auto-title")
};

try
{
    var result = new GovDocumentFormatter().Format(inputPath, outputPath, options);
    Console.WriteLine("排版完成：" + result.OutputPath);
    foreach (var message in result.Messages)
    {
        Console.WriteLine(" - " + message);
    }

    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine("排版失败：" + ex.Message);
    return 1;
}

static bool HasFlag(string[] args, string flag)
{
    return args.Any(arg => string.Equals(arg, flag, StringComparison.OrdinalIgnoreCase));
}

static void PrintUsage(string version)
{
    Console.WriteLine("GovDocFormatter Linux CLI v" + version);
    Console.WriteLine();
    Console.WriteLine("用法：");
    Console.WriteLine("  GovDocFormatter <输入.docx> <输出.docx> [选项]");
    Console.WriteLine();
    Console.WriteLine("选项：");
    Console.WriteLine("  --no-page-numbers      不添加单双页页码");
    Console.WriteLine("  --no-format-existing   不格式化现有正文和标题");
    Console.WriteLine("  --no-auto-title        不自动识别标题");
    Console.WriteLine("  --version, -v          显示版本号");
    Console.WriteLine("  --help, -h             显示帮助");
}
