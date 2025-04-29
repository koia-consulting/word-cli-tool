using System.Text.Json.Serialization;

namespace WordDoc.Models;

/// <summary>
/// Type of change to apply to a document
/// </summary>
[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ChangeType
{
    /// <summary>
    /// Add a comment to the document
    /// </summary>
    Comment,

    /// <summary>
    /// Add a suggested change (tracked change) to the document
    /// </summary>
    Suggestion
}
