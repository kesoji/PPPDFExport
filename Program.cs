using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Text.RegularExpressions;

namespace PPPDFExport
{
    class Program
    {
        static void Main(string[] args)
        {
            string filepath = args[0];
            string cwd = Directory.GetCurrentDirectory();

            try
            {
                filepath = Path.GetFullPath(filepath);

                PowerPoint.Application ppt = new PowerPoint.Application();
                PowerPoint.Presentation p = ppt.Presentations.Open(filepath);

                string dst = cwd + @"/" + getPDFName(Path.GetFileName(filepath));
                p.SaveCopyAs(dst, PowerPoint.PpSaveAsFileType.ppSaveAsPDF);

                p.Close();
                ppt.Quit();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
            }


        }

        private static string getPDFName(string filename)
        {
            return Regex.Replace(filename, @"\.[^.]+$", ".pdf");
        }
    }
}
