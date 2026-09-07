namespace ProcurementControl.Models;

/// <summary>
/// Строка документов сделки — перенос Get-PurchaseDocuments из
/// PurchaseStore.ps1 (строки 1942-1963).
/// </summary>
public sealed class PurchaseDocumentRow
{
    public long Id { get; set; }

    public string DocumentType { get; set; } = string.Empty;

    public string Supplier { get; set; } = string.Empty;

    public string OriginalName { get; set; } = string.Empty;

    public string StoredPath { get; set; } = string.Empty;

    public string FileHash { get; set; } = string.Empty;

    public string UploadedAt { get; set; } = string.Empty;
}

/// <summary>
/// Результат резолва файла документа — перенос объекта, который возвращает
/// Resolve-PurchaseDocumentFile (PurchaseStore.ps1, строки 1995-2070).
/// </summary>
public sealed record ResolvedDocumentFile(
    long Id,
    string Path,
    string Name,
    string FileHash,
    bool Found,
    bool Changed);
