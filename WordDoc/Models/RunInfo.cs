using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordDoc.Models;

/// <summary>
/// Represents a text run and an offset within it
/// </summary>
public record RunInfo
{
    public required Run Run { get; init; }
    public required int Offset { get; init; } // Overall character offset within the paragraph up to the start of this run's relevant element
    public OpenXmlElement Element { get; init; } // The specific element (Text, DelText, Tab, Break) where the position falls
    public int ElementOffset { get; init; } // The character offset within the Element
}
