using System.Text.Json.Serialization;

namespace WordDoc.Models;

/// <summary>
/// Represents a change to be applied to a Word document
/// </summary>
public record Change
{
    /// <summary>
    /// Position information for the change, specified by text search
    /// </summary>
    [JsonConverter(typeof(TextSearchPositionConverter))]
    public TextSearchPosition Position { get; init; }

    /// <summary>
    /// Text content of the comment or suggestion
    /// </summary>
    public string Text { get; init; }

    /// <summary>
    /// Type of change: Comment or Suggestion
    /// </summary>
    public ChangeType Type { get; init; }
}
