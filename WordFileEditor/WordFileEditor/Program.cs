using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

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

            Tag[] tags = new[] { new Tag("nome", "João"), new Tag("Data", "01/01/2001"), new Tag("objeto", "Fusca") };
            //string strTxt = "Append text in body - OpenAndAddTextToWordDocument";

            List<Tag> foundedTags = FindAllTags(strDoc);
            //SearchAndReplace(strDoc, tags);
            //OpenAndAddTextToWordDocument(strDoc, strTxt);
        }

        public static List<Tag> FindAllTags(string document)
        {
            List<Tag> tags = new List<Tag>();
            object missing = System.Reflection.Missing.Value;

            Regex regex = new Regex(@"{{[\w,\d,\s,\.,\,\-]*}}");


            Application app = new Application();
            app.Visible = false;
            Document doc = app.Documents.Open(@"C:\Temp\DOC TEMPLATE NAO MODIFIQUE - Copia.doc");

            Find findObject = app.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = "{{Responsavel}}";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = "Found";

            object replaceAll = WdReplace.wdReplaceAll;

            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                               ref missing, ref missing, ref missing, ref missing, ref missing,
                               ref replaceAll, ref missing, ref missing, ref missing, ref missing);

            string content = doc.Content.Text;


            doc.Save();

            doc.Close();

            app.Quit();
            return tags;
        }

    }
}
