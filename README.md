# ReplaceTextInWordDocument

This C# project demonstrates how to use the OpenXML SDK library to open a Word document (.docx) and replace specific text with new text.

## Prerequisites

- **.NET SDK**: Ensure that you have the .NET SDK installed. Download it from [the official .NET website](https://dotnet.microsoft.com/download).
- **Open XML SDK**: This project uses the OpenXML SDK library to manipulate Word documents.

### Installing OpenXML SDK

Install the OpenXML SDK package via NuGet:

```bash
dotnet add package DocumentFormat.OpenXml
```

## Project Structure

This project contains a `Docx` class with a constructor that takes the path of the source Word file as a parameter.
There is then a `ReplaceText` method that takes the text to be replaced and the new text as parameters.
Finally, there is a `Save` method that takes the path of the target Word file as a parameter.

For now, the idea is simply to copy this class into a project (and evolve it as needed).
I have intentionally made this a console project to demonstrate its usage.

## Usage

### 1. Prepare the Project

Clone this repository and open it in your preferred C# editor to test, or simply copy the `Docx.cs` file into your own project.

### 2. Add a Word Document

Place a `.docx` Word document in the project folder or specify a path to an existing Word file. In the `example` folder, I created a Word file containing 2 placeholders to be replaced. The first is `#name`, which is written in a fragmented manner, and the second is `#year`, which appears in multiple locations including a text box.

### 3. Run the Code

The main code for text replacement is in the `ReplaceText` method. Here is an example in `Program.cs` showing how to use this code:

```csharp
using ReplaceTextInWordDocument;

// Load the Word document
var docx = new Docx("/Users/heliospeed/sources/ReplaceTextInWordDocument/example/example.docx");

// Perform successive replacements in the Word document (replacement is deliberately case-insensitive)
docx.ReplaceText("#name", "John Doe");
docx.ReplaceText("#year", DateTime.Now.Year.ToString());

// Save the new version of the document
docx.SaveFile("/Users/heliospeed/sources/ReplaceTextInWordDocument/example/exampleOut.docx");
```

### 4. Run the Program

In your terminal, navigate to the project folder and run the following command:

```bash
dotnet run
```

### 5. Verify the Result

The specified text in `example.docx` will have been replaced. You can open the Word document to verify the replacements.

## Notes

- **Case Sensitivity**: This code is case-insensitive. If you want to make it case-sensitive, adjust the logic to search for text with case sensitivity.

## Resources

- [OpenXML SDK Documentation](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk)
