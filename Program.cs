using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace BSTReportPrep
{
    class Program
    {
        static void Main(string[] args)
        {
            string dirpathIN = @ConfigurationManager.AppSettings.Get("INdir");
            string dirpathOUT = @ConfigurationManager.AppSettings.Get("OUTdir");
            var dirIN = new DirectoryInfo(dirpathIN); // папка с файлами
            var dirOUT = new DirectoryInfo(dirpathOUT);
            DataTable CSVtable = new DataTable();
            string fileName = "";

            foreach (FileInfo file in dirIN.GetFiles())
            {
                fileName = Path.GetFileName(file.FullName);
                Console.WriteLine(fileName);
                CSVtable = CSVUtility.CSVUtility.GetDataTableFromCSVFile(file.FullName);
                fileName = dirpathOUT + "\\" + fileName;
                CSVUtility.CSVUtility.ToCSV(CSVtable, fileName);

                CSVtable = null;
                //GC.Collect(1, GCCollectionMode.Forced);
            }

        }

        static void FixFiles(string inDir, string outDir)
        {
            var dirIN = new DirectoryInfo(@inDir); // папка с входящими файлами 
            var dirOUT = new DirectoryInfo(@outDir); // папка с исходящими файлами             
            string fileName = "";

            foreach (FileInfo file in dirIN.GetFiles())
            {
                fileName = Path.GetFileName(file.FullName);
                Console.WriteLine(fileName);
                //FixReport(@file.FullName, @outDir + @"\" + fileName);
            }
        }
    }
}
