using OfficeOpenXml;

namespace MyApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var acceptedExtensions = args.AsQueryable().Skip(1).FirstOrDefault();
            var directory = args.AsQueryable().FirstOrDefault();

            Directory:
            if (directory == null)
            {
                Console.WriteLine("Directório:");
                directory = Console.ReadLine()!.Trim();
            }

            if (!Directory.Exists(directory))
            {
                Console.WriteLine("Pasta não encontrada.");
                directory = null;
                goto Directory;
            }

            if(acceptedExtensions == null)
            {
                Console.WriteLine("Extensões que pretende procurar: (Separadas por \",\")");
                acceptedExtensions = Console.ReadLine()?.Trim();
            }

            var acceptedExtensionsList = acceptedExtensions?.Split(',').ToList() ?? new List<string>();
            GenerateExcel(GetAllFiles(directory, acceptedExtensionsList));
        }

        public static List<string> GetAllFiles(string directory, List<string> acceptedExtensions)
        {
            var files = new List<string>();

            if (!acceptedExtensions.Any(e => !string.IsNullOrEmpty(e))) return Directory.GetFiles(directory, "*.*", SearchOption.AllDirectories).ToList();

            foreach (var extension in acceptedExtensions)
            {
                files.AddRange(Directory.GetFiles(directory, $"*.{extension.Trim()}", SearchOption.AllDirectories).ToList());
            }
            return files;
        }

        public static void GenerateExcel(List<string> files)
        {
            Console.WriteLine($"Foram encontrados {files.Count} ficheiro(s) para exportar.");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Files");
                var row = 1;
                foreach (var file in files)
                {
                    worksheet.Cells[row, 1].Value = file.Split("\\").LastOrDefault();
                    row++;
                }

                var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                package.SaveAs($"{desktopPath}/teste.xlsx");
                Console.WriteLine("Ficheiro criado com sucesso!");
            }
        }
    }
}