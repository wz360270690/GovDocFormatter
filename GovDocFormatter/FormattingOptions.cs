namespace GovDocFormatter;

public sealed class FormattingOptions
{
    public string SerialNumber { get; set; } = "";
    public string SecretLevel { get; set; } = "";
    public string Urgency { get; set; } = "";
    public string DispatchMark { get; set; } = "";
    public string DispatchNumber { get; set; } = "";
    public string SignerName { get; set; } = "";
    public string Title { get; set; } = "";
    public string MainRecipient { get; set; } = "";
    public string Issuer { get; set; } = "";
    public string DocumentDate { get; set; } = "";
    public string CopyTo { get; set; } = "";
    public string PrintingOffice { get; set; } = "";
    public string PrintingDate { get; set; } = "";

    public bool InsertFrontMatter { get; set; } = true;
    public bool InsertEndingMatter { get; set; } = true;
    public bool ApplyPageNumbers { get; set; } = true;
    public bool FormatExistingContent { get; set; } = true;
    public bool AutoDetectTitle { get; set; } = true;
    public bool UpwardDocument { get; set; }

    public bool HasFrontMatter =>
        HasText(SerialNumber) ||
        HasText(SecretLevel) ||
        HasText(Urgency) ||
        HasText(DispatchMark) ||
        HasText(DispatchNumber) ||
        HasText(SignerName) ||
        HasText(Title) ||
        HasText(MainRecipient);

    public bool HasEndingMatter =>
        HasText(Issuer) ||
        HasText(DocumentDate) ||
        HasText(CopyTo) ||
        HasText(PrintingOffice) ||
        HasText(PrintingDate);

    private static bool HasText(string value) => !string.IsNullOrWhiteSpace(value);
}

public sealed class FormatResult
{
    public required string OutputPath { get; init; }
    public List<string> Messages { get; } = [];
}
