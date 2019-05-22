using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Extraction_list_Dossiers
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.GetEncoding(1252);
            // get application full path
            //string currentPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            if (args.Length == 0 || args.Length > 1)
            {
                Console.WriteLine("Utiliser 1 argument le path");
                return;
            }
            string path = args[0];
            if (Directory.Exists(path))
            {
                // get folder names list
                IEnumerable<string> dirs = Directory.GetDirectories(path, "scan*", SearchOption.AllDirectories)
                                            .Select(d => new DirectoryInfo(d).Name);
                if (dirs.Count() == 0)
                    Console.WriteLine("Aucun dossier trouve!!");
                // get just Dossier and index
                IEnumerable<string[]> doss_inds = dirs.Select(d => d.Split('_').Skip(1).ToArray());
                // Clear dirs
                dirs = Enumerable.Empty<string>();
                Console.WriteLine("Dossier\tIndice\n");


                /******** Extraire les données dans un fichiers *********/
                CreateExcelFile(doss_inds, path);
            }
            else
                Console.WriteLine("Path introuvabe!!");
            

            //Console.ReadKey();
        }
        private static void CreateExcelFile(IEnumerable<string[]> doss_inds, string path)
        {
            path += "\\list_dossier.xlsx";
            Application xlApp = new Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not installed in the system...");
                return;
            }

            object misValue = System.Reflection.Missing.Value;

            Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Dossier";
            xlWorkSheet.Cells[1, 2] = "Indice";
            int row = 2;
            Console.WriteLine("Creation du fichier Excel...");
            foreach (var dos_ind in doss_inds)
            {
                xlWorkSheet.Cells[row, 1] = dos_ind[0];
                xlWorkSheet.Cells[row, 2] = dos_ind[1];
                drawTextProgressBar(row - 1, doss_inds.Count());
                row++;
            }
            

            xlWorkBook.SaveAs(path, XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue,
                XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Excel file created successfully...");
            Console.WriteLine("Open file " + path);
            Console.ForegroundColor = ConsoleColor.White;
            Process.Start(path);
        }

        private static void drawTextProgressBar(int progress, int total)
        {
            //draw empty progress bar
            Console.CursorLeft = 0;
            Console.Write("["); //start
            Console.CursorLeft = 32;
            Console.Write("]"); //end
            Console.CursorLeft = 1;
            float onechunk = 30.0f / total;

            //draw filled part
            int position = 1;
            for (int i = 0; i < onechunk * progress; i++)
            {
                Console.BackgroundColor = ConsoleColor.Gray;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw unfilled part
            for (int i = position; i <= 31 ; i++)
            {
                Console.BackgroundColor = ConsoleColor.Green;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw totals
            Console.CursorLeft = 35;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write(progress.ToString() + " of " + total.ToString() + "    "); //blanks at the end remove any excess
        }
    }

    
}
