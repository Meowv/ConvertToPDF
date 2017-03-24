using System;
using System.IO;
using System.Text;
using Aspose.Words;

namespace PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("start...");
            bool ok = false;
            for (int i = 1; i <= 10; i++)
            {
                string sourcePath = @"D:\qix\doc\" + i + ".docx";
                string targetPath = @"D:\qix\pdf\" + i + "(docx).pdf";

                ok = OfficeToPdf.DOCConvertToPDF(sourcePath, targetPath);
                if (!ok)
                {
                    StreamReader reader = new StreamReader(sourcePath, Encoding.GetEncoding("gb2312"));
                    Document doc = new Document();
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    builder.Write(reader.ReadToEnd());
                    doc.Save(targetPath, SaveFormat.Pdf);
                    reader.Close();
                }
                Console.WriteLine("{0}.docx --> {1}.(docx).pdf ---- success", i, i);
            }
            Console.WriteLine("end...");
        }
    }
}
