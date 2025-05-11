using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using WordDocLibrary;

partial class Program
{
    static void Main()
    {
        IWordDocumentService wordService = new WordDocumentService();

        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputFolder); // Ensure the folder exists

        string fileName = "SampleDocument.docx";
        string filePath = Path.Combine(outputFolder, fileName);

        // Create a new document
        wordService.CreateDocument(filePath, "This is a sample paragraph.");

        // Add a comment
        wordService.AddComment(filePath, "Filip", "This is a comment!");

        // (Optional) Fake a 'suggested edit'
        wordService.AddSuggestedEdit(filePath, " [suggested added text]");

        // Check if the file exists
        if (File.Exists(filePath))
        {
            Console.WriteLine($"Document created successfully at: {filePath}");
        }
        else
        {
            Console.WriteLine("Failed to create the document.");
        }
    }
}