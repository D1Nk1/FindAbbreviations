using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;


namespace FindAbbreviations
{
    class Program
    {
        static void Main(string[] args)
        {
            Application application = new Application();
            Document document = application.Documents.Open("C:\\word.docx");
            
            int count = document.Words.Count;
            for (int i = 1; i <= count; i++)
            {
                string text = document.Words[i].Text;
                Console.WriteLine(Regex.Match(text, "[A-Z]{2,}"));
                //Console.WriteLine("Word {0} = {1}", i, text);
            }
            application.Quit();
            Console.ReadKey();
        }
    }
}
