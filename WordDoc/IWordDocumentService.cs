using WordDoc.Models;

namespace WordDoc;

/// <summary>
///     Interface for the Word document service
/// </summary>
public interface IWordDocumentService
{
    /// <summary>
    ///     Applies a list of changes to a Word document
    /// </summary>
    /// <param name="filePath">Path to the Word document</param>
    /// <param name="changes">List of changes to apply</param>
    void ApplyChangesToDocument(string filePath, List<Change> changes);
}
