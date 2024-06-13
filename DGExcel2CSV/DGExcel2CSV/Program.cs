using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace DGExcel2CSV
{
    internal class Program
    {
        private static Application app;
        private static Workbook workbook;

        private static string targetFolder;
        
        public static void Main(string[] args)
        {
            Console.WriteLine("--------- Excel 2 CSV ---------");
            Console.WriteLine();
            
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("No arguments");
                return;
            }

            if (args.Length > 2)
            {
                Console.WriteLine("Too many arguments.");
                return;
            }
            Console.WriteLine("Entered Arguments ---");
            for (int i = 0; i < args.Length; i++)
            {
                Console.WriteLine($"\t[{i + 1}] {args[i]}");
            }

            bool isFile = File.Exists(args[0]);
            bool isDirectory = Directory.Exists(args[0]); 

            if (isFile == false && isDirectory == false)
            {
                Console.WriteLine($"Argument [0] is not a file or directory. {args[0]}");
                return;
            }
            
            // Output 경로가 주어지지 않았거나, null 일 경우
            if (args.Length == 1 || string.IsNullOrEmpty(args[1]))
            {
                // 파일의 상위폴더 혹은 직접 받은 폴더로 지정
                targetFolder = isFile ? Path.GetDirectoryName(args[0]) : args[0];
                // 새 폴더 생성
                targetFolder = Path.Combine(targetFolder, "_CSV");
                Console.WriteLine($"Make Temp CSV output folder : {targetFolder}");
            }
            else
            {
                targetFolder = args[1];
            }
            
            Console.WriteLine();
            Console.WriteLine($"Target folder : {targetFolder}");
            Console.WriteLine();
            
            if (Directory.Exists(targetFolder) == false)
            {
                Directory.CreateDirectory(targetFolder);
            }

            if (isFile)
            {
                Convert(args[0]);
            }
            else if (isDirectory)
            {
                var fileList = Directory.GetFiles(args[0], "*.xlsx");
                foreach (var file in fileList)
                {
                    Convert(file);
                }
            }
            Console.WriteLine();
            Console.WriteLine("Finished.");
            return;
        }

        private static void Convert(string file)
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            workbook = app.Workbooks.Open(file);
            var csvPath = Path.ChangeExtension(file, ".csv");
            var csvFileName = Path.GetFileName(csvPath);
            var savePath = Path.Combine(targetFolder, csvFileName); 
            Console.Write($"Converting... [{file}]");
            
            workbook.SaveAs(savePath, XlFileFormat.xlCSV);
            Console.WriteLine(" Saved!");
            
            workbook.Close(false);
            app.Quit();

            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(app);
        }
    }
}
