using System.Text;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ReplaceTextInWordDocument
{
    public class Docx
    {
        public MemoryStream WordDocumentIn { get; set; }       
        public MemoryStream WordDocumentOut { get; set; }
        public Docx(string inputFileName)
        {
            this.WordDocumentIn = new MemoryStream();
            this.WordDocumentOut = new MemoryStream();

            byte[] inputBytes = System.IO.File.ReadAllBytes(inputFileName);

            using (MemoryStream inputStream = new MemoryStream()){
                inputStream.Write(inputBytes, 0, inputBytes.Length);

                using (var document = WordprocessingDocument.Open(inputStream, false)){
                    document.Clone(this.WordDocumentIn);
                }
            }
        }

        public void ReplaceText(string oldValue, string newValue)
        {
            this.WordDocumentIn.Position = 0;
            using (var wordDoc = WordprocessingDocument.Open(this.WordDocumentIn, true))
            {
                var mainDocument = wordDoc.MainDocumentPart;

                if (mainDocument == null)
                {
                    return;
                }

                var body = mainDocument.Document.Body;
                if (body == null)
                {   
                    return;
                }

                var texts = body.Descendants<Text>().ToList();
                var fullText = new StringBuilder();
                var textElements = new List<Text>();
                var lastIndex = 0;

                // Parcourir les éléments Text pour reconstituer le texte complet.
                foreach (var text in texts)
                {
                    fullText.Append(text.Text);
                    textElements.Add(text);

                    // Vérification si le texte reconstitué contient le chaine à remplacer.
                    int index = fullText.ToString().IndexOf(oldValue, lastIndex, StringComparison.InvariantCultureIgnoreCase);
                    if (index > 0) {
                        ReplaceTextFragements(textElements, index, oldValue.Length, newValue);

                        var newText = Regex.Replace(fullText.ToString(), Regex.Escape(oldValue), new string(' ', newValue.Length), RegexOptions.IgnoreCase);
                        fullText.Clear().Append(newText);

                        lastIndex = index + 1;
                        //break;
                    }
                }

                SaveStream(wordDoc, textElements);
            }
        }

        private static void ReplaceTextFragements(List<Text> textElements, int startIndex, int length, string newText)
        {
            int currentIndex = 0;
            int remainingLength = length;

            var replacementText = new StringBuilder(newText);
            var replacedAtOnce = false;

            foreach (var text in textElements)  
            {
                var currentText = text.Text;

                if (startIndex >= currentIndex + currentText.Length){
                    currentIndex += currentText.Length;
                    continue;
                }

                var localStart = Math.Max(0, startIndex - currentIndex);
                var localEnd = Math.Min(currentText.Length, startIndex + remainingLength - currentIndex);

                if (localStart <= localEnd){
                    var replacementLength = localEnd - localStart;
                    var replacedChar = replacementLength;

                    if (remainingLength > 0){
                        if (!replacedAtOnce){
                            // Remplacement complet même si c'est un fragment
                            text.Text = currentText.Remove(localStart, replacementLength)
                                                   .Insert(localStart, replacementText.ToString());
                            replacedAtOnce = true;
                        }
                        else{
                            // Nettoyage des fragments restant
                            text.Text = currentText.Remove(localStart, replacementLength);
                        }
                    }

                    remainingLength -= replacedChar;
                }

                currentIndex += currentText.Length;
                startIndex += currentText.Length;

                if (remainingLength <= 0){
                    break;
                }
            }
        }

        private void SaveStream(WordprocessingDocument wordDoc, List<Text> textElements){
            if (wordDoc == null || wordDoc.MainDocumentPart == null)
            { 
                return;
            }

            // OpenXML 2
            // using  (var sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
            // {
            //     sw.Write(textElements);
            // }
            // wordDoc.Save();

            // OpenXML 3
            wordDoc.MainDocumentPart.Document.Save();
            
            this.WordDocumentOut = new MemoryStream();
            wordDoc.Clone(this.WordDocumentOut);
        }

        public bool Save(string outputFileName){
            var fileFullName = outputFileName;

            try {
                File.WriteAllBytes(fileFullName, this.WordDocumentOut.ToArray());
            }
            catch {
                return false;
            }

            return true;
        }
    }
}