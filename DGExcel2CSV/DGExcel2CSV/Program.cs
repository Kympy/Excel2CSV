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
            if (args == null || args.Length == 0) return;

            if (args.Length > 2) return;

            bool isFile = File.Exists(args[0]);
            bool isDirectory = Directory.Exists(args[0]); 

            if (isFile == false && isDirectory == false)
            {
                return;
            }
            
            // Output 경로가 주어지지 않았거나, null 일 경우
            if (args.Length == 1 || string.IsNullOrEmpty(args[1]))
            {
                // 파일의 상위폴더 혹은 직접 받은 폴더로 지정
                targetFolder = isFile ? Path.GetDirectoryName(args[0]) : args[0];
                // 새 폴더 생성
                targetFolder = Path.Combine(targetFolder, "_CSV");
            }

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

            return;
        }

        private static void Convert(string file)
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            workbook = app.Workbooks.Open(file);

            var csvPath = Path.ChangeExtension(file, ".csv");
            
            workbook.SaveAs(csvPath, XlFileFormat.xlCSV);
            
            workbook.Close(false);
            app.Quit();

            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(app);
        }
    }
}
