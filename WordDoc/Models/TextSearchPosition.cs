using System.Text.Json;
using System.Text.Json.Serialization;

namespace WordDoc.Models;

/// <summary>
/// Represents a position within a document located by text search
/// </summary>
public record TextSearchPosition
{
    /// <summary>
    /// The text to search for in the document
    /// </summary>
    public string SearchText { get; init; }

    /// <summary>
    /// The occurrence of the text to use (1-based index, defaults to first occurrence)
    /// </summary>
    public int Occurrence { get; init; } = 1;

    /// <summary>
    /// Controls whether the search is case-sensitive
    /// </summary>
    public bool CaseSensitive { get; init; } = false;

    /// <summary>
    /// Optional exact text to mark the end of the selection (if different from SearchText)
    /// If null, the entire SearchText will be marked
    /// </summary>
    public string EndSearchText { get; init; }

    /// <summary>
    /// The occurrence of the end text to use (1-based index, defaults to first occurrence)
    /// </summary>
    public int EndOccurrence { get; init; } = 1;
}

/// <summary>
/// Converter for serializing/deserializing TextSearchPosition
/// </summary>
public class TextSearchPositionConverter : JsonConverter<TextSearchPosition>
{
    public override TextSearchPosition Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType != JsonTokenType.StartObject)
        {
            throw new JsonException("Expected start of object");
        }

        string searchText = string.Empty;
        int occurrence = 1;
        bool caseSensitive = false;
        string endSearchText = null;
        int endOccurrence = 1;

        while (reader.Read() && reader.TokenType != JsonTokenType.EndObject)
        {
            if (reader.TokenType != JsonTokenType.PropertyName)
            {
                throw new JsonException("Expected property name");
            }

            string propertyName = reader.GetString();
            reader.Read(); // Move to property value

            switch (propertyName)
            {
                case "searchText":
                    searchText = reader.GetString() ?? string.Empty;
                    break;
                case "occurrence":
                    occurrence = reader.TokenType == JsonTokenType.Number ? reader.GetInt32() : 1;
                    break;
                case "caseSensitive":
                    caseSensitive = reader.TokenType == JsonTokenType.True;
                    break;
                case "endSearchText":
                    endSearchText = reader.GetString();
                    break;
                case "endOccurrence":
                    endOccurrence = reader.TokenType == JsonTokenType.Number ? reader.GetInt32() : 1;
                    break;
                default:
                    reader.Skip(); // Skip unknown properties
                    break;
            }
        }

        return new TextSearchPosition
        {
            SearchText = searchText,
            Occurrence = occurrence,
            CaseSensitive = caseSensitive,
            EndSearchText = endSearchText,
            EndOccurrence = endOccurrence
        };
    }

    public override void Write(Utf8JsonWriter writer, TextSearchPosition value, JsonSerializerOptions options)
    {
        writer.WriteStartObject();
        writer.WriteString("searchText", value.SearchText);
        writer.WriteNumber("occurrence", value.Occurrence);
        writer.WriteBoolean("caseSensitive", value.CaseSensitive);

        if (value.EndSearchText != null)
        {
            writer.WriteString("endSearchText", value.EndSearchText);
            writer.WriteNumber("endOccurrence", value.EndOccurrence);
        }

        writer.WriteEndObject();
    }
}
