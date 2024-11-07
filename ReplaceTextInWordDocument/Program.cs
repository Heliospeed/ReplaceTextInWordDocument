// See https://aka.ms/new-console-template for more information
using ReplaceTextInWordDocument;

var docx = new Docx("../example/example.docx");

docx.ReplaceText("#name", "John Doe");
docx.ReplaceText("#year", DateTime.Now.Year.ToString());

docx.Save("../example/exampleOut.docx");