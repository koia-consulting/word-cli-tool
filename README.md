```markdown
# WordDoc CLI Application

This application allows you to create, edit, and comment on Word documents using the command line.

## Prerequisites

- .NET 6.0 or later installed on your system.
- Ensure the `DocumentFormat.OpenXml` package is installed (already included in the project).

## Build and Run

1. Clone the repository and navigate to the project directory.
2. Build the project:
   ```bash
   dotnet build
   ```
3. Run the application using the CLI.

## Usage

The application supports the following modes:

### 1. Create a Document
Creates a new Word document with a specified filename and content.

**Command:**
```bash
dotnet run createDocument <fileName> <content>
```

**Example:**
```bash
dotnet run createDocument MyDocument.docx "This is the content of the document."
```

### 2. Add a Comment
Adds a comment to the first paragraph of an existing Word document.

**Command:**
```bash
dotnet run addComment <filePath> <author> <commentText>
```

**Example:**
```bash
dotnet run addComment ./output/MyDocument.docx Filip "This is a comment!"
```

### 3. Suggest an Edit
Adds a suggested edit to the first paragraph of an existing Word document.

**Command:**
```bash
dotnet run suggestEdit <filePath>
```

**Example:**
```bash
dotnet run suggestEdit ./output/MyDocument.docx
```

## Output
All generated or modified documents are saved in the `output` directory within the project folder.

## Notes
- Ensure the `output` directory exists or will be created automatically when running the commands.
- Use absolute or relative paths for file operations.
```