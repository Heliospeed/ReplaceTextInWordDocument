using ReplaceTextInWordDocument;

using var docx = new Docx("../../../../example/example.docx");

docx.ReplaceText("#name", "John Doe");
docx.ReplaceText("#year", DateTime.Now.Year.ToString());

docx.Save("../../../../example/exampleOut.docx");