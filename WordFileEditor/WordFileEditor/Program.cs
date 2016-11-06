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
        public static void SearchAndReplace(string docPath,List<Tag> tags)
        {
            Application app = new Application();
            app.Visible = false;
            Document doc = app.Documents.Open(docPath);

            foreach (Tag tag in tags)
            {
                Find findObject = app.Selection.Find;
                findObject.ClearFormatting();
                //Valor da tag(deve adicionar {{}}?)
                findObject.Text = tag.Name;
                findObject.Replacement.ClearFormatting();
                //Valor que o usuario digitou
                findObject.Replacement.Text = tag.Value;

                object replaceAll = WdReplace.wdReplaceAll;

                object missing = System.Reflection.Missing.Value;

                findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                                   ref missing, ref missing, ref missing, ref missing, ref missing,
                                   ref replaceAll, ref missing, ref missing, ref missing, ref missing);
            }

            doc.Save();

            doc.Close();

            app.Quit();
        }

        public static List<Tag> FindAllTags(string docPath)
        {
            List<Tag> tags = new List<Tag>();

            Regex regex = new Regex(@"{{[\w,\d,\s,\.,\,\-]*}}");

            Application app = new Application();
            app.Visible = false;
            Document doc = app.Documents.Open(docPath);

            MatchCollection matches = regex.Matches(doc.Content.Text);

            foreach (Match item in matches)
            {
                Tag tag = new Tag();
                tag.Name = item.Value.Trim(new[] { '{', '}' });
            }

            return tags;
        }

    }
}
