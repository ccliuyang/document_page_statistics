using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;
using iTextSharp.text.pdf;
using System.Reflection;


namespace PageCounterConsole
{
    class Program
    {
        static int totalPageCount = 0;

        static List<string> files = new List<string>();

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("请指定包含需要统计页码的文件或文件夹路径。");
                return;
            }

            string path = args[0];

            if (File.Exists(path))
            {
                files.Add(path);
            }
            else if (Directory.Exists(path))
            {
                string[] extensions = { "doc", "docx", "pdf" };
                foreach (string ext in extensions)
                {
                    string[] filesInFolder = Directory.GetFiles(path, "*." + ext);
                    files.AddRange(filesInFolder);
                }
            }
            else
            {
                Console.WriteLine("指定的路径无效。");
                return;
            }

            foreach (string file in files)
            {
                int pageCount = 0;

                string ext = Path.GetExtension(file);

                switch (ext)
                {
                    case ".doc":
                    case ".docx":
                        pageCount = GetWordPageCount(file);
                        break;
                    case ".pdf":
                        pageCount = GetPDFPageCount(file);
                        break;
                }

                totalPageCount += pageCount;

                Console.WriteLine(Path.GetFileName(file) + "\t" + pageCount + "页");
            }

            Console.WriteLine("总页码数：" + totalPageCount);
        }

        static int GetWordPageCount(string filePath)
        {
            Application application = new Application();
            Document document = application.Documents.Open(filePath);

            // int pageCount = document.Content.ComputeStatistics(WdStatistic.wdStatisticPages, false);
            int pageCount = document.ComputeStatistics(WdStatistic.wdStatisticPages, Missing.Value);

            document.Close();
            application.Quit();

            return pageCount;
        }

        static int GetPDFPageCount(string filePath)
        {
            using (PdfReader reader = new PdfReader(filePath))
            {
                return reader.NumberOfPages;
            }
        }
    }
}
