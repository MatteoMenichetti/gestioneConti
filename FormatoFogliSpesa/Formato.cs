using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq.Expressions;
using System.Text;

namespace FormatoFogliSpesa;

public class Formato
{
    private static readonly string _samplePath = "Sample.xlsx";
    public static void AggiuntaAnno(string path, int ? year = null)
    {
        var b = Path.Exists(path);
        var b1 = Path.Exists(_samplePath);
        
        if (Path.Exists(path) && Path.Exists(_samplePath))
        {
            using var workbook = new XLWorkbook(path);
            using var sampleWB = new XLWorkbook(_samplePath);
            year = year.HasValue ? year.Value : DateTime.Now.Year;

            AggiuntaFoglioAnno(workbook, year.Value);

            AggiornaFogliStorico(sampleWB, workbook, year.Value);
            
            ModificaFormule(sampleWB, workbook, year.Value);
            
            var name = Path.GetFileName(path);
            var nameWithoutExtension = Path.GetFileNameWithoutExtension(path);
            var extension = Path.GetExtension(path);
            var fullPath = Path.GetFullPath(path);
            var directoryPath = fullPath.Substring(0, fullPath.Length - name.Length);
            var sb = new StringBuilder(directoryPath).Append(nameWithoutExtension+"New").Append(extension);
            var finalPath = sb.ToString();
                
            workbook.SaveAs(finalPath);
            
        }
    }

    private static void AggiornaFogliStorico(XLWorkbook sampleWb, XLWorkbook workbook, int year)
    {
        var storicoTelefonoSampleSheet = sampleWb.Worksheets.Worksheet("Storico Telefono");
        var storicoTelefonoSheet = workbook.Worksheets.Worksheet("Storico Telefono");

        var storicoStipendioSampleSheet = sampleWb.Worksheets.Worksheet("Storico Stipendio");

        Dictionary<IXLWorksheet, int> worksheets = new Dictionary<IXLWorksheet, int>();
        
        worksheets.Add(workbook.Worksheets.Worksheet("Storico Telefono"), 3);
        worksheets.Add(workbook.Worksheets.Worksheet("Storico Affitto"), 4);
        worksheets.Add(workbook.Worksheets.Worksheet("Storico Stipendio"), 4);
        
        List<string> lastValues = new List<string>();

        foreach (var worksheet in worksheets)
        {
            var ws = worksheet.Key;
            var storicoSampleSheet = sampleWb.Worksheets.Worksheet(ws.Name);
            
            int newRightLimit = worksheet.Value;
            
            var lastCUsed = worksheet.Key.Row(2).LastCellUsed().Address.ColumnNumber;

            var columnToStart = lastCUsed + 1;

            //configurazione intestazione anno
            ws.Cell(1, columnToStart).Style.Font.Bold = true;
            ws.Range(1, columnToStart, 1, columnToStart + newRightLimit).Merge().Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;
            ws.Cell(1, columnToStart).Value = year.ToString();

            //copio contenuto celle descrizione del contenuto e applico l'ultimo aggiornamento dei prezzi

            for (int i = 0; i < newRightLimit; i++)
            {
                var tmp = storicoSampleSheet.Cell(2, i + 1).Value;
                ws.Cell(2, columnToStart + i).Value = storicoSampleSheet.Cell(2, i + 1).Value;
                var lastValue = ws.Column(lastCUsed - newRightLimit + i + 1).LastCellUsed().Value;
                
                ws.Cell(3, columnToStart + i).Value = lastValue;
            }
        }

    var storicoAffittoSampleSheet = sampleWb.Worksheets.Worksheet("Storico Affitto");
        
        
    }

    private static void ComposizioneFormulaModificaTelefono()
    {
        
    }

    private static void AggiuntaFoglioAnno(XLWorkbook workbook, int year)
    {
        try
        {
            workbook.Worksheets.Add(year);
        }
        catch (Exception ex)
        {
            if(ex is ArgumentException && ex.Message.Contains("A worksheet with the same name"))
                Console.WriteLine(ex.Message);
        }
    }
    private static void ModificaFormule(XLWorkbook sample, XLWorkbook workbook, int year)
    {
        
        
    }

    public static void AggiornaFormulaStipendio(string path, int? year = null)
    {
        year = year.HasValue ? year.Value : DateTime.Now.Year;
        XLWorkbook workbook = new XLWorkbook(path);
        
        IXLWorksheet yearSheet = workbook.Worksheet(year.Value.ToString());
        IXLWorksheet storicoStipendioSheet = workbook.Worksheet("Storico Stipendio");

        int[] ultimoStipendio = [storicoStipendioSheet.LastColumnUsed().LastCellUsed().Address.RowNumber, storicoStipendioSheet.LastColumnUsed().LastCellUsed().Address.ColumnNumber - 3];
        
        var rowProp = storicoStipendioSheet.Range(ultimoStipendio[0], ultimoStipendio[1], ultimoStipendio[0], ultimoStipendio[1]+3)
            .CellsUsed().Select(c => c.Value.ToString()).ToList();

        var range = storicoStipendioSheet.Range(ultimoStipendio[0], ultimoStipendio[1], ultimoStipendio[0], ultimoStipendio[1] + 1).RangeAddress;

        var salaryCell = range.FirstAddress.ColumnLetter + range.FirstAddress.RowNumber;
        var dayCell = range.LastAddress.ColumnLetter + range.LastAddress.RowNumber;
        var cell = yearSheet.Search("Base").ToList();

        int monthToStartUpdate = Int32.Parse(rowProp[3]);
        cell.RemoveRange(0, monthToStartUpdate);
        
        foreach (var formulaStipendio in cell)
        {
            var str = $"=IF(TODAY()>=DATA({year.Value},{monthToStartUpdate},'Storico Stipendio'!{dayCell}),'Storico Stipendio'!{salaryCell},)";
            var str20 = str[20];
            var res = yearSheet.Cell(formulaStipendio.Address.RowNumber, formulaStipendio.Address.ColumnNumber + 1).FormulaA1 =
                $"=IF(TODAY()>=DATE({year.Value},{++monthToStartUpdate},'Storico Stipendio'!{dayCell}),'Storico Stipendio'!{salaryCell},)";;
        }
        
        workbook.Save();
    }

    public static void AggiornaFormuleTelefono(string path, int? year = null)
    {
        year = year.HasValue ? year.Value : DateTime.Now.Year;
        XLWorkbook workbook = new XLWorkbook(path);
        
        IXLWorksheet yearSheet = workbook.Worksheet(year.Value.ToString());
        IXLWorksheet storicoTelefonoSheet = workbook.Worksheet("Storico Telefono");
        
        int[] ultimoCosto = [storicoTelefonoSheet.LastColumnUsed().LastCellUsed().Address.RowNumber, storicoTelefonoSheet.LastColumnUsed().LastCellUsed().Address.ColumnNumber - 2];
        
        var rowProp = storicoTelefonoSheet.Range(ultimoCosto[0], ultimoCosto[1], ultimoCosto[0], ultimoCosto[1]+2)
            .CellsUsed().Select(c => c.Value.ToString()).ToList();

        var range = storicoTelefonoSheet.Range(ultimoCosto[0], ultimoCosto[1], ultimoCosto[0], ultimoCosto[1] + 2).RangeAddress;

        var costCell = range.FirstAddress.ColumnLetter + range.FirstAddress.RowNumber;
        var dayCell = "B" + range.FirstAddress.RowNumber;
        var celleDaAggiornare = yearSheet.Search("Telefono").ToList();

        int monthToStartUpdate = Int32.Parse(rowProp[2]);
        celleDaAggiornare.RemoveRange(0, monthToStartUpdate);
        
        foreach (var formulaTelefono in celleDaAggiornare)
        {
            var str = $"=IF(TODAY()>=DATA({year.Value},{monthToStartUpdate},'Storico Telefono'!{dayCell}),'Storico Telefono'!{costCell},)";
            var str20 = str[20];
            var res = yearSheet.Cell(formulaTelefono.Address.RowNumber, formulaTelefono.Address.ColumnNumber + 2).FormulaA1 =
                $"=IF(TODAY()>=DATE({year.Value},{++monthToStartUpdate},'Storico Telefono'!{dayCell}),'Storico Telefono'!{costCell},)";;
        }
        
        workbook.Save();
    }
    
    public static void AggiornaFormuleAffitto(string path, int? year = null)
    {
        year = year.HasValue ? year.Value : DateTime.Now.Year;
        XLWorkbook workbook = new XLWorkbook(path);
        
        IXLWorksheet yearSheet = workbook.Worksheet(year.Value.ToString());
        IXLWorksheet storicoAffittoSheet = workbook.Worksheet("Storico Affitto");
        
        int[] ultimoCosto = [storicoAffittoSheet.LastColumnUsed().LastCellUsed().Address.RowNumber, storicoAffittoSheet.LastColumnUsed().LastCellUsed().Address.ColumnNumber - 3];
        
        var rowProp = storicoAffittoSheet.Range(ultimoCosto[0], ultimoCosto[1], ultimoCosto[0], ultimoCosto[1]+3)
            .CellsUsed().Select(c => c.Value.ToString()).ToList();

        var range = storicoAffittoSheet.Range(ultimoCosto[0], ultimoCosto[1], ultimoCosto[0], ultimoCosto[1] + 3).RangeAddress;

        var costCell = range.FirstAddress.ColumnLetter + range.FirstAddress.RowNumber;
        var dayCell = "C" + range.FirstAddress.RowNumber;
        var condominioCell = "B"+range.FirstAddress.RowNumber;
        var celleAffittoDaAggiornare = yearSheet.Search("Affitto").ToList();
        var celleSpeseCondomialiDaAggiornare = yearSheet.Search("Spese condiminio").ToList();

        int monthToStartUpdate = Int32.Parse(rowProp[3]);
        celleAffittoDaAggiornare.RemoveRange(0, monthToStartUpdate);
        celleSpeseCondomialiDaAggiornare.RemoveRange(0, monthToStartUpdate);
        
        foreach (var formulaAffitto in celleAffittoDaAggiornare)
        {
            var str = $"=IF(TODAY()>=DATA({year.Value},{monthToStartUpdate},'Storico Affitto'!{dayCell}),'Storico Affitto'!{costCell},)";
            
            var res = yearSheet.Cell(formulaAffitto.Address.RowNumber, formulaAffitto.Address.ColumnNumber + 2).FormulaA1 =
                $"=IF(TODAY()>=DATE({year.Value},{++monthToStartUpdate},'Storico Affitto'!{dayCell}),'Storico Affitto'!{costCell},)";;
        }

        foreach (var formulaSpesaCondomiale in celleSpeseCondomialiDaAggiornare)
        {
            var res = yearSheet.Cell(formulaSpesaCondomiale.Address.RowNumber, formulaSpesaCondomiale.Address.ColumnNumber + 2).FormulaA1 =
                $"=IF(TODAY()>=DATE({year.Value},{++monthToStartUpdate},'Storico Affitto'!{condominioCell}),'Storico Affitto'!{condominioCell},)";;

        }
        
        workbook.Save();
    }
}