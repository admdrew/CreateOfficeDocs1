using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using CommandLine;
using CommandLine.Text;

namespace CreateOfficeDocs1 {
    class Program {
        //private static Random random = new Random((int)DateTime.Now.Ticks);//thanks to McAden

        class Options {
            // doc type
            [Option('t', "type", Required = true, HelpText = "Document type.")]
            public string opDocumentType { get; set; }

            // number of pages/length?
            //[Option('

            // images yes/no
 
            // output file name prefix

            // help
            [HelpOption]
            public string GetUsage() {
                StringBuilder strUsage = new StringBuilder();
                strUsage.AppendLine("CreatOfficeDocs");
                strUsage.AppendLine("helpstuff");

                return strUsage.ToString();
            }
        }

        private static string RandomString(int size) {
            Random random = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++) {
                //ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                ch = (char)random.Next('a', 'z' + 1);
                builder.Append(ch);
            }

            return builder.ToString();
        }

        static void Main(string[] args) {
            //MakeWordDoc(100);
            //MakeWordDoc(1000);
            //MakeWordDoc(10000);

            MakeWordDoc();  // create a single 1-page file with no images

            // Take input from Options
            //

            // fire off MakeWordDoc with input
            //
        }

        /* MakeWordDoc([int numPages], [int numFiles], [bool includeImages], [String fileNamePrefix])
         * 
         */
        static int MakeWordDoc(
                int numPages = 1,                   // Number of output pages
                int numFiles = 1,                   // Number of output files
                bool includeImages = false,         // Include embedded images.
                String fileNamePrefix = "WordDoc"   // Prefix name of output file(s) 
            ) {
            // make word application
            int MakeWordDoc_success = 0;
            Application appWord = new Application();

            // make word doc
            Document docWord;
            string fileName = numPages.ToString() + "pages.docx";

            // necessary objects?
            // EDIT - noooope. thanks, .net 4.0
            //object objMiss = System.Reflection.Missing.Value;
            //object endofdoc = "\\endofdoc";

            //appWord.Visible = true;
            //docWord = appWord.Documents.Add(ref objMiss, ref objMiss, ref objMiss, ref objMiss);
            docWord = appWord.Documents.Add();

            // add 1 big paragraph per page
            for (int i = 0; i < numPages; i++) {
                Paragraph para;
                para = docWord.Content.Paragraphs.Add();
                para.Range.Text = "page " + i.ToString() + " of " + numPages.ToString() + " " + RandomString(3000);
                if (i % 10 == 0)
                    Console.Out.WriteLine(fileName + ": wrote page " + i.ToString() + " of " + numPages.ToString());
                para.Range.InsertParagraphAfter();
                docWord.Words.Last.InsertBreak(WdBreakType.wdPageBreak);
            }

            //string fileName = "C:\\Dropbox\\_dev\\visualstudio\\output\\test.docx";
            
            Console.Out.WriteLine("Writing to \"" + fileName + "\"");
            docWord.SaveAs(fileName);

            //clean up
            appWord.Quit();
            Marshal.FinalReleaseComObject(appWord);
            Marshal.FinalReleaseComObject(docWord);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            return MakeWordDoc_success;
        }
    }
}
