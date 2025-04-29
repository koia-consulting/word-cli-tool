using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using WordDoc.Models;

namespace WordDoc;

/// <summary>
///     Service for modifying Word documents with comments and suggestions
/// </summary>
public class WordDocumentService : IWordDocumentService
{
    private const string AuthorName = "WordDocumentModifier";
    private const string RevisionIdPrefix = "rev_";
    private static int _revisionCounter;

    /// <inheritdoc />
    public void ApplyChangesToDocument(string filePath, List<Change> changes)
    {
        using var doc = OpenDocument(filePath);

        foreach (var change in changes)
        {
            switch (change.Type)
            {
                case ChangeType.Comment:
                    AddCommentToDocument(doc, change.Position, change.Text);
                    break;
                case ChangeType.Suggestion:
                    AddSuggestionToDocument(doc, change.Position, change.Text);
                    break;
                default:
                    throw new ArgumentException($"Unknown change type: {change.Type}");
            }
        }

        SaveDocument(doc);
    }

    private static WordprocessingDocument OpenDocument(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
        {
            throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
        }

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("Document file not found.", filePath);
        }

        try
        {
            return WordprocessingDocument.Open(filePath, true);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to open document: {filePath}", ex);
        }
    }

    private static void SaveDocument(WordprocessingDocument doc)
    {
        if (doc == null)
        {
            throw new ArgumentNullException(nameof(doc), "Document cannot be null.");
        }

        try
        {
            doc.MainDocumentPart?.Document.Save();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to save document.", ex);
        }
    }

    private static SearchResult FindTextInDocument(WordprocessingDocument doc, TextSearchPosition position)
    {
        if (doc == null)
        {
            throw new ArgumentNullException(nameof(doc), "Document cannot be null.");
        }

        if (position == null)
        {
            throw new ArgumentNullException(nameof(position), "Position cannot be null.");
        }

        if (string.IsNullOrEmpty(position.SearchText))
        {
            throw new ArgumentException("Search text cannot be empty.", nameof(position));
        }

        var body = doc.MainDocumentPart?.Document.Body;
        if (body == null)
        {
            throw new InvalidOperationException("Document has no body.");
        }

        Console.WriteLine($"Searching for text: '{position.SearchText}'");
        Console.WriteLine($"Search options: CaseSensitive={position.CaseSensitive}, Occurrence={position.Occurrence}");

        var documentInfo = AnalyzeDocumentStructure(body);

        var (textContainers, combinedText, positionMap) = BuildTextMap(documentInfo);

        var searchResult = SearchForText(position, combinedText, positionMap);
        if (searchResult != null)
        {
            return searchResult;
        }

        searchResult = TryPartialTextSearch(position, combinedText, positionMap);
        if (searchResult != null)
        {
            return searchResult;
        }

        searchResult = TryFirstNonEmptyText(textContainers);
        if (searchResult != null)
        {
            return searchResult;
        }

        return CreateFallbackSearchResult(body, documentInfo.Paragraphs);
    }

    private static DocumentStructureInfo AnalyzeDocumentStructure(Body body)
    {
        var result = new DocumentStructureInfo
        {
            Paragraphs = body.Elements<Paragraph>().ToList(),
            Tables = body.Descendants<Table>().ToList(),
            Runs = body.Descendants<Run>().ToList(),
            TextElements = body.Descendants<Text>().ToList()
        };

        Console.WriteLine($"Paragraphs: {result.Paragraphs.Count}, Tables: {result.Tables.Count}");
        Console.WriteLine($"Runs: {result.Runs.Count}, Text elements: {result.TextElements.Count}");

        if (result.TextElements.Count > 0)
        {
            Console.WriteLine("Text content samples:");
            foreach (var text in result.TextElements.Take(3))
            {
                var preview = text.Text.Length > 30 ? text.Text.Substring(0, 30) + "..." : text.Text;
                Console.WriteLine($"- \"{preview}\"");
            }
        }

        return result;
    }

    private static (List<TextLocationInfo>, string, List<TextPositionMap>) BuildTextMap(DocumentStructureInfo docInfo)
    {
        var textContainers = new List<TextLocationInfo>();

        foreach (var text in docInfo.TextElements)
        {
            var textContent = text.Text;
            if (string.IsNullOrEmpty(textContent))
            {
                continue;
            }

            var run = text.Ancestors<Run>().FirstOrDefault();
            if (run == null)
            {
                continue;
            }

            var paragraph = run.Ancestors<Paragraph>().FirstOrDefault();
            var tableCell = run.Ancestors<TableCell>().FirstOrDefault();

            textContainers.Add(new TextLocationInfo
            {
                TextElement = text,
                Run = run,
                Paragraph = paragraph,
                TableCell = tableCell,
                Content = textContent,
                IsCellContent = tableCell != null
            });
        }

        var combinedText = new StringBuilder();
        var positionMap = new List<TextPositionMap>();

        var globalPosition = 0;
        foreach (var textInfo in textContainers)
        {
            var content = textInfo.Content;
            combinedText.Append(content);

            for (var i = 0; i < content.Length; i++)
            {
                positionMap.Add(new TextPositionMap
                {
                    GlobalPosition = globalPosition + i, TextInfo = textInfo, LocalPosition = i
                });
            }

            globalPosition += content.Length;
        }

        var fullText = combinedText.ToString();
        Console.WriteLine($"Combined text length: {fullText.Length} characters");
        if (fullText.Length > 0)
        {
            var preview = fullText.Length > 50 ? fullText.Substring(0, 50) + "..." : fullText;
            Console.WriteLine($"Text preview: '{preview}'");
        }

        return (textContainers, fullText, positionMap);
    }

    private static SearchResult SearchForText(
        TextSearchPosition position,
        string combinedText,
        List<TextPositionMap> positionMap)
    {
        var comparisonType = position.CaseSensitive
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;

        var searchIndex = 0;
        var currentOccurrence = 0;

        while ((searchIndex = combinedText.IndexOf(position.SearchText, searchIndex, comparisonType)) != -1)
        {
            currentOccurrence++;
            var endIndex = searchIndex + position.SearchText.Length;
            Console.WriteLine($"Found occurrence {currentOccurrence} at position {searchIndex}-{endIndex}");

            if (currentOccurrence == position.Occurrence)
            {
                if (searchIndex < positionMap.Count && endIndex <= positionMap.Count)
                {
                    return CreateSearchResultFromPositions(
                        searchIndex,
                        endIndex,
                        positionMap,
                        position.SearchText);
                }
            }

            searchIndex += position.SearchText.Length;
        }

        return null;
    }

    private static SearchResult TryPartialTextSearch(
        TextSearchPosition position,
        string combinedText,
        List<TextPositionMap> positionMap)
    {
        if (position.SearchText.Length <= 5)
        {
            return null;
        }

        var partialText = position.SearchText.Substring(0, 5);
        Console.WriteLine($"Trying with partial text: '{partialText}'");

        var comparisonType = position.CaseSensitive
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;

        var searchIndex = combinedText.IndexOf(partialText, 0, comparisonType);
        if (searchIndex == -1)
        {
            return null;
        }

        Console.WriteLine($"Found partial match at position {searchIndex}");
        var approximateEndIndex = Math.Min(searchIndex + position.SearchText.Length, positionMap.Count - 1);

        return CreateSearchResultFromPositions(
            searchIndex,
            approximateEndIndex,
            positionMap,
            "Partial match: " + partialText);
    }

    private static SearchResult TryFirstNonEmptyText(List<TextLocationInfo> textContainers)
    {
        foreach (var textInfo in textContainers)
        {
            if (string.IsNullOrWhiteSpace(textInfo.Content))
            {
                continue;
            }

            var container = textInfo.Paragraph ??
                            textInfo.TableCell?.Descendants<Paragraph>().FirstOrDefault();

            if (container == null)
            {
                continue;
            }

            var contentPreview = textInfo.Content.Substring(0, Math.Min(20, textInfo.Content.Length));
            Console.WriteLine($"Using first non-empty text: '{contentPreview}'");

            var startRunInfo = new RunInfo
            {
                Run = textInfo.Run, Offset = 0, Element = textInfo.TextElement, ElementOffset = 0
            };

            var endOffset = Math.Min(textInfo.Content.Length, 10);
            var endRunInfo = new RunInfo
            {
                Run = textInfo.Run, Offset = endOffset, Element = textInfo.TextElement, ElementOffset = endOffset
            };

            return new SearchResult
            {
                StartParagraph = container,
                StartRunInfo = startRunInfo,
                EndParagraph = container,
                EndRunInfo = endRunInfo,
                OriginalText = "First text content"
            };
        }

        return null;
    }

    private static SearchResult CreateFallbackSearchResult(Body body, List<Paragraph> paragraphs)
    {
        if (paragraphs.Count > 0)
        {
            var firstParagraph = paragraphs[0];
            var firstRun = firstParagraph.Descendants<Run>().FirstOrDefault();

            if (firstRun != null)
            {
                var textElement = firstRun.Descendants<Text>().FirstOrDefault();

                var runInfo = new RunInfo
                {
                    Run = firstRun,
                    Offset = 0,
                    Element = textElement ?? firstRun.Descendants().FirstOrDefault(),
                    ElementOffset = 0
                };

                Console.WriteLine("Using first paragraph as fallback");

                return new SearchResult
                {
                    StartParagraph = firstParagraph,
                    StartRunInfo = runInfo,
                    EndParagraph = firstParagraph,
                    EndRunInfo = runInfo,
                    OriginalText = "Fallback position"
                };
            }

            var newRun = new Run(new Text(""));
            firstParagraph.AppendChild(newRun);

            var newRunInfo = new RunInfo
            {
                Run = newRun, Offset = 0, Element = newRun.GetFirstChild<Text>(), ElementOffset = 0
            };

            return new SearchResult
            {
                StartParagraph = firstParagraph,
                StartRunInfo = newRunInfo,
                EndParagraph = firstParagraph,
                EndRunInfo = newRunInfo,
                OriginalText = "Created run fallback"
            };
        }

        var newParagraph = new Paragraph(new Run(new Text("")));
        body.AppendChild(newParagraph);

        var createdRunInfo = new RunInfo
        {
            Run = newParagraph.Descendants<Run>().First(),
            Offset = 0,
            Element = newParagraph.Descendants<Text>().First(),
            ElementOffset = 0
        };

        return new SearchResult
        {
            StartParagraph = newParagraph,
            StartRunInfo = createdRunInfo,
            EndParagraph = newParagraph,
            EndRunInfo = createdRunInfo,
            OriginalText = "Created paragraph fallback"
        };
    }

    private static SearchResult CreateSearchResultFromPositions(
        int startIndex,
        int endIndex,
        List<TextPositionMap> positionMap,
        string originalText)
    {
        var startPositionInfo = positionMap[startIndex];
        var endPositionInfo = positionMap[Math.Min(endIndex - 1, positionMap.Count - 1)];

        var startRunInfo = new RunInfo
        {
            Run = startPositionInfo.TextInfo.Run,
            Offset = startPositionInfo.LocalPosition,
            Element = startPositionInfo.TextInfo.TextElement,
            ElementOffset = startPositionInfo.LocalPosition
        };

        var endRunInfo = new RunInfo
        {
            Run = endPositionInfo.TextInfo.Run,
            Offset = endPositionInfo.LocalPosition + 1,
            Element = endPositionInfo.TextInfo.TextElement,
            ElementOffset = endPositionInfo.LocalPosition + 1
        };

        var startContainer = startPositionInfo.TextInfo.Paragraph ??
                             startPositionInfo.TextInfo.TableCell?.Descendants<Paragraph>().FirstOrDefault();

        var endContainer = endPositionInfo.TextInfo.Paragraph ??
                           endPositionInfo.TextInfo.TableCell?.Descendants<Paragraph>().FirstOrDefault();

        if (startContainer == null || endContainer == null)
        {
            return null;
        }

        return new SearchResult
        {
            StartParagraph = startContainer,
            StartRunInfo = startRunInfo,
            EndParagraph = endContainer,
            EndRunInfo = endRunInfo,
            OriginalText = originalText
        };
    }

    private static void AddCommentToDocument(WordprocessingDocument doc, TextSearchPosition position,
        string commentText)
    {
        if (doc == null)
        {
            throw new ArgumentNullException(nameof(doc));
        }

        if (position == null)
        {
            throw new ArgumentNullException(nameof(position));
        }

        if (string.IsNullOrEmpty(commentText))
        {
            throw new ArgumentException("Comment text cannot be null or empty.", nameof(commentText));
        }

        var commentsPart = EnsureCommentsPartExists(doc);
        var commentId = GenerateUniqueCommentId(commentsPart);

        var searchResult = FindTextInDocument(doc, position);

        var commentMarkers = CreateCommentMarkers(commentId);

        InsertCommentMarkersAtPosition(
            searchResult.StartParagraph,
            searchResult.EndParagraph,
            searchResult.StartRunInfo,
            searchResult.EndRunInfo,
            commentMarkers);

        AddCommentToCommentsPart(commentsPart, commentId, commentText);
    }

    private static WordprocessingCommentsPart EnsureCommentsPartExists(WordprocessingDocument doc)
    {
        if (doc?.MainDocumentPart == null)
        {
            throw new InvalidOperationException("Document or MainDocumentPart is null.");
        }

        var commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart;
        if (commentsPart == null)
        {
            commentsPart = doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments();
        }

        return commentsPart;
    }

    private static int GenerateUniqueCommentId(WordprocessingCommentsPart commentsPart)
    {
        if (commentsPart?.Comments == null)
        {
            return 1;
        }

        return commentsPart.Comments.Elements<Comment>()
            .Select(c => int.TryParse(c.Id?.Value, out var id) ? id : 0)
            .DefaultIfEmpty(0)
            .Max() + 1;
    }

    private static (CommentRangeStart, CommentRangeEnd, CommentReference) CreateCommentMarkers(int commentId)
    {
        var commentIdStr = commentId.ToString();
        return (
            new CommentRangeStart { Id = commentIdStr },
            new CommentRangeEnd { Id = commentIdStr },
            new CommentReference { Id = commentIdStr }
        );
    }

    private static void AddSuggestionToDocument(WordprocessingDocument doc, TextSearchPosition position, string newText)
    {
        if (doc == null)
        {
            throw new ArgumentNullException(nameof(doc));
        }

        if (position == null)
        {
            throw new ArgumentNullException(nameof(position));
        }

        if (newText == null)
        {
            throw new ArgumentNullException(nameof(newText));
        }

        EnsureTrackingIsEnabled(doc);

        var searchResult = FindTextInDocument(doc, position);

        InsertSuggestionChange(
            searchResult.StartParagraph,
            searchResult.StartRunInfo,
            searchResult.EndParagraph,
            searchResult.EndRunInfo,
            searchResult.OriginalText,
            newText,
            doc);
    }

    private static void EnsureTrackingIsEnabled(WordprocessingDocument doc)
    {
        if (doc?.MainDocumentPart == null)
        {
            throw new InvalidOperationException("Document or MainDocumentPart is null.");
        }

        var settingsPart = doc.MainDocumentPart.DocumentSettingsPart;
        if (settingsPart == null)
        {
            settingsPart = doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();
        }

        if (settingsPart.Settings.Elements<TrackRevisions>().FirstOrDefault() == null)
        {
            settingsPart.Settings.AppendChild(new TrackRevisions());
        }

        var rsids = settingsPart.Settings.Elements<Rsids>().FirstOrDefault();
        if (rsids == null)
        {
            rsids = new Rsids();
            settingsPart.Settings.AppendChild(rsids);
        }

        if (rsids.RsidRoot == null)
        {
            rsids.RsidRoot = new RsidRoot { Val = GenerateRsid() };
        }

        var revisionView = settingsPart.Settings.Elements<RevisionView>().FirstOrDefault();
        if (revisionView == null)
        {
            settingsPart.Settings.AppendChild(new RevisionView { Markup = true });
        }
        else
        {
            revisionView.Markup = true;
        }
    }

    private static HexBinaryValue GenerateRsid()
    {
        var bytes = new byte[4];
        Random.Shared.NextBytes(bytes);
        return new HexBinaryValue(BitConverter.ToString(bytes).Replace("-", ""));
    }

    /// <summary>
    ///     Inserts comment markers at the specified position.
    ///     Handles cases within a single run, across runs in the same paragraph,
    ///     and potentially across paragraphs. Also properly handles position-only comments
    ///     where start and end positions are the same.
    /// </summary>
    private static void InsertCommentMarkersAtPosition(
        Paragraph startParagraph,
        Paragraph endParagraph,
        RunInfo startRunInfo,
        RunInfo endRunInfo,
        (CommentRangeStart, CommentRangeEnd, CommentReference) commentMarkers)
    {
        var (commentStart, commentEnd, commentReference) = commentMarkers;
        var startRun = startRunInfo.Run;
        var startOffset = startRunInfo.Offset;
        var endRun = endRunInfo.Run;
        var endOffset = endRunInfo.Offset;

        var startMarker = (CommentRangeStart)commentStart.CloneNode(true);
        var endMarker = (CommentRangeEnd)commentEnd.CloneNode(true);
        var referenceMarker = (CommentReference)commentReference.CloneNode(true);

        var isPositionOnly = startRun == endRun &&
                             startParagraph == endParagraph &&
                             startOffset == endOffset;

        if (isPositionOnly)
        {
            var originalRunProperties = startRun.RunProperties?.CloneNode(true) as RunProperties;

            if (startOffset == 0)
            {
                startRun.InsertBeforeSelf(referenceMarker);
            }
            else
            {
                var originalTextElement = startRun.Elements<Text>().FirstOrDefault();
                var originalText = originalTextElement?.Text ?? "";

                Run runBefore = null;
                if (startOffset > 0 && startOffset <= originalText.Length)
                {
                    runBefore = new Run(
                        originalRunProperties?.CloneNode(true),
                        new Text(originalText[..startOffset]) { Space = SpaceProcessingModeValues.Preserve }
                    );
                }

                Run runAfter = null;
                if (startOffset < originalText.Length)
                {
                    runAfter = new Run(
                        originalRunProperties?.CloneNode(true),
                        new Text(originalText[startOffset..]) { Space = SpaceProcessingModeValues.Preserve }
                    );
                }

                if (runBefore != null)
                {
                    startRun.InsertBeforeSelf(runBefore);
                }

                startRun.InsertBeforeSelf(referenceMarker);

                if (runAfter != null)
                {
                    startRun.InsertBeforeSelf(runAfter);
                }

                startRun.Remove();
            }
        }
        else if (startRun == endRun && startParagraph == endParagraph)
        {
            var originalTextElement = startRun.Elements<Text>().FirstOrDefault();
            var originalText = originalTextElement?.Text ?? "";
            var originalRunProperties = startRun.RunProperties?.CloneNode(true) as RunProperties;

            Run runPartBefore = null;
            if (startOffset > 0 && startOffset <= originalText.Length)
            {
                runPartBefore = new Run(
                    originalRunProperties?.CloneNode(true),
                    new Text(originalText[..startOffset]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            Run runPartMiddle = null;
            if (endOffset > startOffset && startOffset < originalText.Length)
            {
                var middleLength = Math.Min(endOffset, originalText.Length) - startOffset;
                if (middleLength > 0)
                {
                    runPartMiddle = new Run(
                        originalRunProperties?.CloneNode(true),
                        new Text(originalText.Substring(startOffset, middleLength))
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        }
                    );
                }
            }

            Run runPartAfter = null;
            if (endOffset < originalText.Length)
            {
                runPartAfter = new Run(
                    originalRunProperties?.CloneNode(true),
                    new Text(originalText[endOffset..]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            OpenXmlElement referenceNode = startRun;
            if (runPartBefore != null)
            {
                referenceNode.InsertBeforeSelf(runPartBefore);
            }

            referenceNode.InsertBeforeSelf(startMarker);
            if (runPartMiddle != null)
            {
                referenceNode.InsertBeforeSelf(runPartMiddle);
            }

            referenceNode.InsertBeforeSelf(endMarker);
            referenceNode.InsertBeforeSelf(referenceMarker);
            if (runPartAfter != null)
            {
                referenceNode.InsertBeforeSelf(runPartAfter);
            }

            startRun.Remove();
        }
        else
        {
            var startTextElement = startRun.Elements<Text>().FirstOrDefault();
            var startText = startTextElement?.Text ?? "";
            var startRunProperties = startRun.RunProperties?.CloneNode(true) as RunProperties;
            Run startPartBefore = null;
            Run startPartInside = null;

            if (startOffset > 0 && startOffset <= startText.Length)
            {
                startPartBefore = new Run(
                    startRunProperties?.CloneNode(true),
                    new Text(startText[..startOffset]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            if (startOffset < startText.Length)
            {
                startPartInside = new Run(
                    startRunProperties?.CloneNode(true),
                    new Text(startText[startOffset..]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            OpenXmlElement startReferenceNode = startRun;
            if (startPartBefore != null)
            {
                startReferenceNode.InsertBeforeSelf(startPartBefore);
            }

            startReferenceNode.InsertBeforeSelf(startMarker);
            if (startPartInside != null)
            {
                startReferenceNode.InsertBeforeSelf(startPartInside);
            }

            startRun.Remove();

            var endTextElement = endRun.Elements<Text>().FirstOrDefault();
            var endText = endTextElement?.Text ?? "";
            var endRunProperties = endRun.RunProperties?.CloneNode(true) as RunProperties;
            Run endPartInside = null;
            Run endPartAfter = null;

            if (endOffset > 0 && endOffset <= endText.Length)
            {
                endPartInside = new Run(
                    endRunProperties?.CloneNode(true),
                    new Text(endText[..endOffset]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            if (endOffset < endText.Length)
            {
                endPartAfter = new Run(
                    endRunProperties?.CloneNode(true),
                    new Text(endText[endOffset..]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            OpenXmlElement endReferenceNode = endRun;
            if (endPartInside != null)
            {
                endReferenceNode.InsertBeforeSelf(endPartInside);
            }

            endReferenceNode.InsertBeforeSelf(endMarker);
            endReferenceNode.InsertBeforeSelf(referenceMarker);
            if (endPartAfter != null)
            {
                endReferenceNode.InsertBeforeSelf(endPartAfter);
            }

            endRun.Remove();

            if (startParagraph != endParagraph)
            {
                var currentElement = startParagraph.NextSibling() ??
                                     throw new InvalidOperationException("Cannot find element after start paragraph.");
                while (currentElement != null && currentElement != endParagraph)
                {
                    if (currentElement is Paragraph intermediateParagraph)
                    {
                        intermediateParagraph.PrependChild((CommentRangeStart)commentStart.CloneNode(true));
                        intermediateParagraph.AppendChild((CommentRangeEnd)commentEnd.CloneNode(true));
                    }

                    currentElement = currentElement.NextSibling();
                }

                if (currentElement == null && endParagraph != null)
                {
                    Console.WriteLine(
                        "Warning: Could not find end paragraph sequentially. Multi-paragraph comment might be incomplete.");
                }
            }
        }
    }

    private static void AddCommentToCommentsPart(WordprocessingCommentsPart commentsPart, int commentId,
        string commentText)
    {
        if (commentsPart == null)
        {
            throw new ArgumentNullException(nameof(commentsPart));
        }

        Comment comment = new()
        {
            Id = commentId.ToString(),
            Author = AuthorName,
            Date = DateTime.UtcNow,
            Initials = AuthorName.Substring(0, Math.Min(AuthorName.Length, 3)) // Example initials
        };

        Paragraph commentParagraph = new(new Run(new Text(commentText)));
        comment.AppendChild(commentParagraph);

        commentsPart.Comments.AppendChild(comment);
    }

    /// <summary>
    ///     Inserts suggestion (tracked change) markers and text.
    ///     Handles deletion of original text and insertion of new text.
    /// </summary>
    private static void InsertSuggestionChange(
        Paragraph startParagraph,
        RunInfo startRunInfo,
        Paragraph endParagraph,
        RunInfo endRunInfo,
        string originalText,
        string newText,
        WordprocessingDocument doc)
    {
        var revisionId = RevisionIdPrefix + Interlocked.Increment(ref _revisionCounter);

        if (startParagraph == endParagraph && startRunInfo.Run == endRunInfo.Run)
        {
            var originalRun = startRunInfo.Run;
            var startOffset = startRunInfo.Offset;
            var endOffset = endRunInfo.Offset;

            var originalTextElement = originalRun.Elements<Text>().FirstOrDefault();
            var originalRunFullText = originalTextElement?.Text ?? "";
            var originalRunProperties = originalRun.RunProperties?.CloneNode(true) as RunProperties;

            Run runPartBefore = null;
            if (startOffset > 0 && startOffset <= originalRunFullText.Length)
            {
                runPartBefore = new Run(
                    originalRunProperties?.CloneNode(true),
                    new Text(originalRunFullText[..startOffset]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            DeletedRun deletedRunStructure = null;
            if (!string.IsNullOrEmpty(originalText))
            {
                // Structure: <w:del w:author="..." w:date="..." w:id="...">
                //              <w:r><w:delText>...</w:delText></w:r>
                //            </w:del>
                deletedRunStructure = new DeletedRun(
                    new Run(
                        new RunProperties(
                            originalRunProperties?.CloneNode(true)),
                        new DeletedText { Text = originalText, Space = SpaceProcessingModeValues.Preserve }
                    )
                )
                {
                    Author = AuthorName, Date = DateTime.UtcNow, Id = revisionId
                };
            }

            InsertedRun insertedRunStructure = null;
            if (!string.IsNullOrEmpty(newText))
            {
                // Structure: <w:ins w:author="..." w:date="..." w:id="...">
                //              <w:r>
                //                <w:rPr>...</w:rPr> // Optional formatting
                //                <w:t>...</w:t>
                //              </w:r>
                //            </w:ins>
                insertedRunStructure = new InsertedRun(
                    new Run(
                        originalRunProperties?.CloneNode(true),
                        new Text { Text = newText, Space = SpaceProcessingModeValues.Preserve }
                    )
                )
                {
                    Author = AuthorName, Date = DateTime.UtcNow, Id = revisionId
                };
            }

            Run runPartAfter = null;
            if (endOffset >= 0 && endOffset < originalRunFullText.Length)
            {
                runPartAfter = new Run(
                    originalRunProperties?.CloneNode(true),
                    new Text(originalRunFullText[endOffset..]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            OpenXmlElement referenceNode = originalRun;
            if (runPartBefore != null)
            {
                referenceNode.InsertBeforeSelf(runPartBefore);
            }

            if (deletedRunStructure != null)
            {
                referenceNode.InsertBeforeSelf(deletedRunStructure);
            }

            if (insertedRunStructure != null)
            {
                referenceNode.InsertBeforeSelf(insertedRunStructure);
            }

            if (runPartAfter != null)
            {
                referenceNode.InsertBeforeSelf(runPartAfter);
            }

            originalRun.Remove();
        }
        else
        {
            var startRun = startRunInfo.Run;
            var startOffset = startRunInfo.Offset;
            var startRunProperties = startRun.RunProperties?.CloneNode(true);
            var startTextElement = startRun.Elements<Text>().FirstOrDefault();
            var startRunText = startTextElement?.Text ?? "";

            Run startPartBefore = null;
            if (startOffset > 0 && startOffset <= startRunText.Length)
            {
                startPartBefore = new Run(
                    startRunProperties?.CloneNode(true),
                    new Text(startRunText[..startOffset]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            var endRun = endRunInfo.Run;
            var endOffset = endRunInfo.Offset;
            var endRunProperties = endRun.RunProperties?.CloneNode(true);
            var endTextElement = endRun.Elements<Text>().FirstOrDefault();
            var endRunText = endTextElement?.Text ?? "";

            Run endPartAfter = null;
            if (endOffset >= 0 && endOffset < endRunText.Length)
            {
                endPartAfter = new Run(
                    endRunProperties?.CloneNode(true),
                    new Text(endRunText[endOffset..]) { Space = SpaceProcessingModeValues.Preserve }
                );
            }

            DeletedRun deletedRunStructure = null;
            if (!string.IsNullOrEmpty(originalText))
            {
                deletedRunStructure = new DeletedRun(
                    new Run(
                        new RunProperties(startRunProperties?.CloneNode(true)),
                        new DeletedText { Text = originalText, Space = SpaceProcessingModeValues.Preserve }
                    )
                ) { Author = AuthorName, Date = DateTime.UtcNow, Id = revisionId };
            }

            InsertedRun insertedRunStructure = null;
            if (!string.IsNullOrEmpty(newText))
            {
                insertedRunStructure = new InsertedRun(
                    new Run(
                        startRunProperties?.CloneNode(true),
                        new Text { Text = newText, Space = SpaceProcessingModeValues.Preserve }
                    )
                ) { Author = AuthorName, Date = DateTime.UtcNow, Id = revisionId };
            }

            OpenXmlElement startReferenceNode = startRun;
            if (startPartBefore != null)
            {
                startReferenceNode.InsertBeforeSelf(startPartBefore);
            }

            if (startParagraph == endParagraph)
            {
                var runs = startParagraph.Elements<Run>().ToList();

                var startIdx = runs.IndexOf(startRun);
                var endIdx = runs.IndexOf(endRun);

                if (startIdx >= 0 && endIdx >= 0 && startIdx < endIdx)
                {
                    for (var i = startIdx + 1; i < endIdx; i++)
                    {
                        runs[i].Remove();
                    }
                }
            }
            else
            {
                var startRuns = startParagraph.Elements<Run>().ToList();
                var startIdx = startRuns.IndexOf(startRun);
                if (startIdx >= 0)
                {
                    for (var i = startIdx + 1; i < startRuns.Count; i++)
                    {
                        startRuns[i].Remove();
                    }
                }

                var endRuns = endParagraph.Elements<Run>().ToList();
                var endIdx = endRuns.IndexOf(endRun);
                if (endIdx >= 0)
                {
                    for (var i = 0; i < endIdx; i++)
                    {
                        endRuns[i].Remove();
                    }
                }

                var body = doc.MainDocumentPart?.Document?.Body;
                var paragraphs = body.Elements<Paragraph>().ToList();
                var startParaIdx = paragraphs.IndexOf(startParagraph);
                var endParaIdx = paragraphs.IndexOf(endParagraph);

                if (startParaIdx >= 0 && endParaIdx >= 0 && startParaIdx < endParaIdx - 1)
                {
                    for (var i = startParaIdx + 1; i < endParaIdx; i++)
                    {
                        paragraphs[i].Remove();
                    }
                }
            }

            startReferenceNode.InsertBeforeSelf(deletedRunStructure);

            startReferenceNode.InsertBeforeSelf(insertedRunStructure);

            if (endPartAfter != null)
            {
                endRun.InsertAfterSelf(endPartAfter);
            }

            startRun.Remove();
            endRun.Remove();
        }
    }

    private class DocumentStructureInfo
    {
        public List<Paragraph> Paragraphs { get; set; }
        public List<Table> Tables { get; set; }
        public List<Run> Runs { get; set; }
        public List<Text> TextElements { get; set; }
    }

    private class TextLocationInfo
    {
        public Text TextElement { get; set; }
        public Run Run { get; set; }
        public Paragraph Paragraph { get; set; }
        public TableCell TableCell { get; set; }
        public string Content { get; set; }
        public bool IsCellContent { get; set; }
    }

    private class TextPositionMap
    {
        public int GlobalPosition { get; set; }
        public TextLocationInfo TextInfo { get; set; }
        public int LocalPosition { get; set; }
    }

    /// <summary>
    ///     Represents the result of a text search in the document
    /// </summary>
    private class SearchResult
    {
        public Paragraph StartParagraph { get; init; }
        public RunInfo StartRunInfo { get; init; }
        public Paragraph EndParagraph { get; init; }
        public RunInfo EndRunInfo { get; init; }
        public string OriginalText { get; init; }
    }
}
