using System.Text.Json;
using WordDoc.Models;

namespace WordDoc;

internal class Program
{
    private static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            ShowUsage();
            return;
        }

        var filePath = args[0];
        var jsonChanges = args[1];

        // Example changes using text search positioning
        var justComment = @"[{
            ""position"": {
                ""searchText"": ""himenaeos.

Lorem ipsum dolor
"",
                ""occurrence"": 1,
                ""caseSensitive"": false
            },
            ""text"": ""This is a comment 3"",
            ""type"": ""Comment""
        }]";

        var justSuggestion = @"[{
            ""position"": {
                ""searchText"": ""old text"",
                ""occurrence"": 1
            },
            ""text"": ""new text"",
            ""type"": ""Suggestion""
        }]";

        var both = @"[
            {
                ""position"": {
                    ""searchText"": ""First paragraph"",
                    ""occurrence"": 1
                },
                ""text"": ""This needs revision"",
                ""type"": ""Comment""
            },
            {
                ""position"": {
                    ""searchText"": ""Introduction"",
                    ""endSearchText"": ""Conclusion"",
                    ""occurrence"": 1,
                    ""endOccurrence"": 1
                },
                ""text"": ""This entire section is good"",
                ""type"": ""Comment""
            },
            {
                ""position"": {
                    ""searchText"": ""outdated information"",
                    ""occurrence"": 1
                },
                ""text"": ""current information"",
                ""type"": ""Suggestion""
            },
            {
                ""position"": {
                    ""searchText"": ""Chapter 3"",
                    ""endSearchText"": ""Chapter 4"",
                    ""occurrence"": 1,
                    ""endOccurrence"": 1
                },
                ""text"": ""Need to restructure this entire chapter"",
                ""type"": ""Comment""
            }
        ]";

        // Parse the JSON changes (using the provided JSON or an example)
        List<Change> changes;
        try
        {
            changes = ParseChanges(justComment);
        }
        catch
        {
            Console.WriteLine("Could not parse provided JSON. Using example changes instead.");
            changes = ParseChanges(both);
        }

        try
        {
            var documentService = new WordDocumentService();
            Console.WriteLine($"Applying changes to document: {filePath}");

            documentService.ApplyChangesToDocument(filePath, changes);
            Console.WriteLine("Document modified successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static List<Change> ParseChanges(string modificationsJson)
    {
        try
        {
            // Configure JsonSerializerOptions to handle our custom converters
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

            // Deserialize the JSON
            var changes = JsonSerializer.Deserialize<List<Change>>(modificationsJson, options);
            return changes ?? throw new ArgumentException("Invalid JSON format for modifications");
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid JSON format: {ex.Message}");
        }
    }

    private static void ShowUsage()
    {
        Console.WriteLine("Word Document Modifier - Add comments and suggested changes to Word documents");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  WordDocumentModifier.Cli <document_path> <modifications_json>");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  document_path     - Path to the Word document to modify");
        Console.WriteLine("  modifications_json - JSON array of modifications or path to a JSON file");
        Console.WriteLine();
        Console.WriteLine("JSON Format:");
        Console.WriteLine(@"  [
    {
        ""position"": {
            ""searchText"": ""Text to find"",
            ""occurrence"": 1,
            ""caseSensitive"": false,
            ""endSearchText"": ""Optional end text"",
            ""endOccurrence"": 1
        },
        ""text"": ""This is a comment"",
        ""type"": ""Comment""
    },
    {
        ""position"": {
            ""searchText"": ""Text to replace"",
            ""occurrence"": 1
        },
        ""text"": ""Replacement text"",
        ""type"": ""Suggestion""
    }
]");
        Console.WriteLine();
        Console.WriteLine("Position properties:");
        Console.WriteLine("  searchText     - Text to search for in the document");
        Console.WriteLine("  occurrence     - Which occurrence of the text to use (default: 1)");
        Console.WriteLine("  caseSensitive  - Whether the search is case-sensitive (default: false)");
        Console.WriteLine("  endSearchText  - Optional end text for range selection");
        Console.WriteLine("  endOccurrence  - Which occurrence of the end text to use (default: 1)");
        Console.WriteLine();
        Console.WriteLine("Types:");
        Console.WriteLine("  Comment    - Add a comment to the specified text");
        Console.WriteLine("  Suggestion - Add a suggested change to replace the specified text");
    }
}
