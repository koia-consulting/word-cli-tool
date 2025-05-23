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
            changes = ParseChanges(jsonChanges);
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
        Console.WriteLine("Word Document Modifier - Add comments and suggestions to Word documents by searching for text");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  WordDocumentModifier.Cli <document_path> <modifications_json_or_file>");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  document_path           - Path to the Word document to modify");
        Console.WriteLine("  modifications_json_or_file - JSON array of modifications or path to a JSON file");
        Console.WriteLine();
        Console.WriteLine("Example Command:");
        Console.WriteLine("  WordDocumentModifier.Cli document.docx \"[{\\\"position\\\":{\\\"searchText\\\":\\\"Hello world\\\"},\\\"text\\\":\\\"Nice greeting\\\",\\\"type\\\":\\\"Comment\\\"}]\"");
        Console.WriteLine("  WordDocumentModifier.Cli document.docx changes.json");
        Console.WriteLine();
        Console.WriteLine("JSON Format:");
        Console.WriteLine(@"[
    {
        ""position"": {
            ""searchText"": ""Text to find"",           // Required: The text to search for in the document
            ""occurrence"": 1,                        // Optional: Which occurrence to use (default: 1)
            ""caseSensitive"": false,                 // Optional: Match case (default: false)
            ""endSearchText"": ""Optional end text"",   // Optional: Text marking the end of a range
            ""endOccurrence"": 1                      // Optional: Which end text occurrence (default: 1)
        },
        ""text"": ""This is a comment"",                // The comment or replacement text
        ""type"": ""Comment""                           // Type: ""Comment"" or ""Suggestion""
    }
]");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine();
        Console.WriteLine("1. Add a comment to specific text:");
        Console.WriteLine(@"[{
    ""position"": {""searchText"": ""Lorem ipsum""},
    ""text"": ""This needs revision"",
    ""type"": ""Comment""
}]");
        Console.WriteLine();
        Console.WriteLine("2. Replace text with a suggestion:");
        Console.WriteLine(@"[{
    ""position"": {""searchText"": ""old text""},
    ""text"": ""new text"",
    ""type"": ""Suggestion""
}]");
        Console.WriteLine();
        Console.WriteLine("3. Comment on a range of text from start to end:");
        Console.WriteLine(@"[{
    ""position"": {
        ""searchText"": ""Chapter 1"",
        ""endSearchText"": ""Chapter 2"",
        ""occurrence"": 1,
        ""endOccurrence"": 1
    },
    ""text"": ""This entire chapter needs work"",
    ""type"": ""Comment""
}]");
        Console.WriteLine();
        Console.WriteLine("4. Find the second occurrence of text (when it appears multiple times):");
        Console.WriteLine(@"[{
    ""position"": {""searchText"": ""Lorem ipsum"", ""occurrence"": 2},
    ""text"": ""This is the second occurrence"",
    ""type"": ""Comment""
}]");
    }
}
