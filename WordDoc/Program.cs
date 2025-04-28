using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordDoc;

internal static class Program
{
    private static void Main(string[] args)
    {
        for (var index = 0; index < args.Length; index++)
        {
            Console.WriteLine($"Id={index}: {args[index]}");
        }

        var outputFolder = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName, "output");
        Directory.CreateDirectory(outputFolder); // Ensure the folder exists

        var fileName = "SampleDocument.docx";
        var filePath = Path.Combine(outputFolder, fileName);

        // Create a new document
        using (var wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            // Add main document part
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var body = mainPart.Document.Body ?? new Body();

            // Add a paragraph
            var para = new Paragraph();
            var run = new Run();
            run.Append(new Text("This is a sample paragraph."));

            para.Append(run);
            body.Append(para);

            // Add a comment
            AddComment(mainPart, run, "Filip", "This is a comment!");

            // (Optional) Fake a 'suggested edit'
            AddSuggestedEdit(mainPart, para);

            mainPart.Document.Save();
        }

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

    static void AddComment(MainDocumentPart mainPart, Run run, string author, string commentText)
    {
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments();

        var comment = new Comment()
        {
            Id = "0",
            Author = author,
            Date = DateTime.Now
        };
        comment.AppendChild(new Paragraph(new Run(new Text(commentText))));
        commentsPart.Comments.Append(comment);
        commentsPart.Comments.Save();

        run.PrependChild(new CommentRangeStart() { Id = "0" });
        run.AppendChild(new CommentRangeEnd() { Id = "0" });
        run.AppendChild(new Run(new CommentReference() { Id = "0" }));
    }

    static void AddSuggestedEdit(MainDocumentPart mainPart, Paragraph para)
    {
        var run = new Run();
        var ins = new InsertedRun()
        {
            Id = "1",
            Author = "Filip",
            Date = DateTime.Now
        };
        ins.Append(new Text(" [suggested added text]"));

        run.Append(ins);
        para.Append(run);
    }
}
