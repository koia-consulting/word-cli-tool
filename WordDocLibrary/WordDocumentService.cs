using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WordDocLibrary;

public interface IWordDocumentService
{
    void CreateDocument(string filePath, string content);
    void AddComment(string filePath, string author, string commentText);
    void RespondToComment(string filePath, string commentId, string responseText);
    void AddSuggestedEdit(string filePath, string editText);
}

public class WordDocumentService : IWordDocumentService
{
    public void CreateDocument(string filePath, string content)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(content)))));
            mainPart.Document.Save();
        }
    }

    public void AddComment(string filePath, string author, string commentText)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
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

            var run = mainPart.Document.Body.Elements<Paragraph>().First().Elements<Run>().First();
            run.PrependChild(new CommentRangeStart() { Id = "0" });
            run.AppendChild(new CommentRangeEnd() { Id = "0" });
            run.AppendChild(new Run(new CommentReference() { Id = "0" }));

            mainPart.Document.Save();
        }
    }

    public void RespondToComment(string filePath, string commentId, string responseText)
    {
        // Implementation for responding to a comment
    }

    public void AddSuggestedEdit(string filePath, string editText)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
        {
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            var para = mainPart.Document.Body.Elements<Paragraph>().First();
            var run = new Run();
            var ins = new InsertedRun()
            {
                Id = "1",
                Author = "Filip",
                Date = DateTime.Now
            };
            ins.Append(new Text(editText));

            run.Append(ins);
            para.Append(run);

            mainPart.Document.Save();
        }
    }
}
