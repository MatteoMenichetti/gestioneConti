using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Text;

namespace FormatoFogliSpesa;

public class Formato
{
    private static string _samplePath = "Sample.xlsx";
    public static void AggiuntaFoglio(string path, int ? year = null)
    {
        var b = Path.Exists(path);
        var b1 = Path.Exists(_samplePath);
        
        if (Path.Exists(path) && Path.Exists(_samplePath))
        {
            using var workbook = new XLWorkbook(path);
            using var sampleWB = new XLWorkbook(_samplePath);
            
            
            year = year.HasValue ? year.Value : DateTime.Now.Year;

            try
            {
                ModificaFormule(sampleWB, year.Value);
                var name = Path.GetFileName(path);
                var nameWithoutExtension = Path.GetFileNameWithoutExtension(path);
                var extension = Path.GetExtension(path);
                var fullPath = Path.GetFullPath(path);
                var directoryPath = fullPath.Substring(0, fullPath.Length - name.Length);
                var sb = new StringBuilder(directoryPath).Append(nameWithoutExtension+"New").Append(extension);
                var finalPath = sb.ToString();
                
                workbook.SaveAs(finalPath);
            }
            catch(Exception ex)
            {
                if(ex is ArgumentException && ex.Message.Contains("A worksheet with the same name"))
                    Console.WriteLine(ex.Message);
            }
        }
    }

    private static void ModificaFormule(XLWorkbook workbook, int year)
    {
        var worksheet = workbook.Worksheet("Anno").CopyTo(workbook, $"{year}");
        for (int i = 0; i < 12; i++)
        {
            worksheet.Cell(4, 21).FormulaA1 =
                $"=SE(OGGI()>DATA({year}; 1; INDIRETTO(\"'Storico Telefono'!$B$4\")); INDIRETTO(\"'Storico Telefono'!$A$4\"); )";
            
            
        }
    }
}