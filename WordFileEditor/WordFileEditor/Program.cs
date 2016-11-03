using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text.RegularExpressions;

namespace WordFileEditor
{
    public class Tag
    {
        public Tag()
        {
        }

        public Tag(string name, string value)
        {
            this.Name = name;
            this.Value = value;
        }

        public string Name { get; set; }
        public string Value { get; set; }
    }
    class Program
    {

        static void Main(string[] args)
        {
            string strDoc = @"C:\Temp\DocExemplo.docx";

            Tag[] tags = new[] { new Tag("nome","João"), new Tag("Data","01/01/2001"), new Tag("objeto","Fusca") };
            //string strTxt = "Append text in body - OpenAndAddTextToWordDocument";

            List<Tag> foundedTags = FindAllTags(strDoc);
            //SearchAndReplace(strDoc, tags);
            //OpenAndAddTextToWordDocument(strDoc, strTxt);
        }

        public static List<Tag> FindAllTags(string document)
        {
            List<Tag> tags = new List<Tag>();

            Regex regex = new Regex(@"{{[\w,\d,\s,\.,\,\-]*}}");


            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                foreach (var item in regex.Matches(docText))
                {
                    tags.Add(new Tag() { Name = item.ToString().Trim('{', '}') });
                }

            }

            return tags;
        }


        public static void SearchAndReplace(string document, Tag[] tags)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }
                
                foreach (Tag tag in tags)
                {
                    string formatedExpression = "{{"+tag.Name+"}}";
                    docText = Regex.Replace(docText, formatedExpression, tag.Value,RegexOptions.IgnoreCase);
                    //docText = regexText.Replace(docText, "Hi Everyone!");
                }


                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        public static void OpenAndAddTextToWordDocument(string filepath, string txt)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(filepath, true);

            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));

            // Close the handle explicitly.
            wordprocessingDocument.Close();
        }

    }
}
