using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Input;
using Avalonia.Platform.Storage;
using Avalonia.Controls;
using ClosedXML.Excel;

namespace CostsManager.ViewModels;

public class MainViewModel : INotifyPropertyChanged
{
    // ── State ───────────────────────────────────────────────────────
    private string _filePath = string.Empty;
    private string _yearInput = string.Empty;
    private string _fileStatusMessage = string.Empty;
    private string _yearStatusMessage = string.Empty;
    private string _rentStatusMessage = string.Empty;
    private string _phoneStatusMessage = string.Empty;
    private bool _isFileLoaded;

    public ObservableCollection<string> LogMessages { get; } = new();

    // ── Properties ──────────────────────────────────────────────────
    public string FilePath
    {
        get => _filePath;
        set { _filePath = value; OnPropertyChanged(); }
    }

    public string YearInput
    {
        get => _yearInput;
        set { _yearInput = value; OnPropertyChanged(); }
    }

    public bool IsFileLoaded
    {
        get => _isFileLoaded;
        private set { _isFileLoaded = value; OnPropertyChanged(); }
    }

    public string FileStatusMessage
    {
        get => _fileStatusMessage;
        private set { _fileStatusMessage = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasFileStatus)); }
    }
    public bool HasFileStatus => !string.IsNullOrEmpty(_fileStatusMessage);

    public string YearStatusMessage
    {
        get => _yearStatusMessage;
        private set { _yearStatusMessage = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasYearStatus)); }
    }
    public bool HasYearStatus => !string.IsNullOrEmpty(_yearStatusMessage);

    public string RentStatusMessage
    {
        get => _rentStatusMessage;
        private set { _rentStatusMessage = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasRentStatus)); }
    }
    public bool HasRentStatus => !string.IsNullOrEmpty(_rentStatusMessage);

    public string PhoneStatusMessage
    {
        get => _phoneStatusMessage;
        private set { _phoneStatusMessage = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasPhoneStatus)); }
    }
    public bool HasPhoneStatus => !string.IsNullOrEmpty(_phoneStatusMessage);

    public bool HasGlobalStatus => LogMessages.Count > 0;

    // ── Commands ────────────────────────────────────────────────────
    public ICommand BrowseFileCommand { get; }
    public ICommand AddYearSheetCommand { get; }
    public ICommand UpdateRentFormulasCommand { get; }
    public ICommand UpdatePhoneFormulasCommand { get; }

    public MainViewModel()
    {
        BrowseFileCommand        = new AsyncRelayCommand(BrowseFileAsync);
        AddYearSheetCommand      = new AsyncRelayCommand(AddYearSheetAsync,  () => IsFileLoaded);
        UpdateRentFormulasCommand  = new AsyncRelayCommand(UpdateRentAsync,   () => IsFileLoaded);
        UpdatePhoneFormulasCommand = new AsyncRelayCommand(UpdatePhoneAsync,  () => IsFileLoaded);
    }

    // ── Browse ──────────────────────────────────────────────────────
    private async Task BrowseFileAsync()
    {
        var topLevel = GetTopLevel();
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Seleziona file Excel",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Excel") { Patterns = new[] { "*.xlsx", "*.xlsm" } }
            }
        });

        if (files.Count == 0) return;

        FilePath = files[0].TryGetLocalPath() ?? string.Empty;
        IsFileLoaded = File.Exists(FilePath);
        FileStatusMessage = IsFileLoaded
            ? $"✓ File caricato: {Path.GetFileName(FilePath)}"
            : "✗ Percorso non valido.";

        AddLog(FileStatusMessage);
        ((AsyncRelayCommand)AddYearSheetCommand).RaiseCanExecuteChanged();
        ((AsyncRelayCommand)UpdateRentFormulasCommand).RaiseCanExecuteChanged();
        ((AsyncRelayCommand)UpdatePhoneFormulasCommand).RaiseCanExecuteChanged();
    }

    // ── Add Year Sheet ───────────────────────────────────────────────
    private async Task AddYearSheetAsync()
    {
        YearStatusMessage = string.Empty;

        if (!int.TryParse(YearInput.Trim(), out int year) || year < 1900 || year > 2200)
        {
            YearStatusMessage = "✗ Anno non valido. Inserisci un numero tra 1900 e 2200.";
            return;
        }

        await Task.Run(() =>
        {
            using var wb = new XLWorkbook(FilePath);
            string sheetName = year.ToString();

            if (wb.TryGetWorksheet(sheetName, out _))
            {
                YearStatusMessage = $"✗ Il foglio '{sheetName}' esiste già.";
                return;
            }

            // Create the new year sheet with a minimal template
            var ws = wb.AddWorksheet(sheetName);
            ws.Cell("A1").Value = $"Anno {year}";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 14;
            ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromArgb(0x0D, 0x1B, 0x3E);
            ws.Cell("A1").Style.Font.FontColor = XLColor.White;

            // Header row
            string[] headers = { "Mese", "Affitto", "Telefono", "Totale" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cell(2, i + 1);
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            }

            // 12 months rows
            string[] months = { "Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
                                 "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre" };
            for (int m = 0; m < 12; m++)
            {
                int row = m + 3;
                ws.Cell(row, 1).Value = months[m];
                ws.Cell(row, 2).FormulaA1 = "0";   // placeholder – rent
                ws.Cell(row, 3).FormulaA1 = "0";   // placeholder – phone
                ws.Cell(row, 4).FormulaA1 = $"B{row}+C{row}";
            }

            ws.Columns().AdjustToContents();
            wb.Save();
        });

        YearStatusMessage = $"✓ Foglio '{year}' aggiunto con successo.";
        AddLog(YearStatusMessage);
        YearInput = string.Empty;
    }

    // ── Update Rent Formulas ─────────────────────────────────────────
    private async Task UpdateRentAsync()
    {
        RentStatusMessage = "⏳ Aggiornamento in corso…";

        await Task.Run(() =>
        {
            using var wb = new XLWorkbook(FilePath);

            // Find all sheets that look like years and update rent column (B)
            int updated = 0;
            foreach (var ws in wb.Worksheets)
            {
                if (!int.TryParse(ws.Name, out _)) continue;

                // Example: recalculate/reset rent formulas in column B rows 3-14
                for (int row = 3; row <= 14; row++)
                {
                    var cell = ws.Cell(row, 2);
                    if (string.IsNullOrWhiteSpace(cell.FormulaA1))
                        cell.FormulaA1 = "0";
                    // In a real scenario you would apply your specific formula here
                    // e.g. cell.FormulaA1 = "RentBase * InflationRate";
                }

                updated++;
            }

            wb.Save();
            RentStatusMessage = $"✓ Formule affitto aggiornate in {updated} foglio/i.";
        });

        AddLog(RentStatusMessage);
    }

    // ── Update Phone Formulas ────────────────────────────────────────
    private async Task UpdatePhoneAsync()
    {
        PhoneStatusMessage = "⏳ Aggiornamento in corso…";

        await Task.Run(() =>
        {
            using var wb = new XLWorkbook(FilePath);

            int updated = 0;
            foreach (var ws in wb.Worksheets)
            {
                if (!int.TryParse(ws.Name, out _)) continue;

                // Example: recalculate/reset phone formulas in column C rows 3-14
                for (int row = 3; row <= 14; row++)
                {
                    var cell = ws.Cell(row, 3);
                    if (string.IsNullOrWhiteSpace(cell.FormulaA1))
                        cell.FormulaA1 = "0";
                    // In a real scenario: cell.FormulaA1 = "PhoneBase + ExtraMinutes * RatePerMin";
                }

                updated++;
            }

            wb.Save();
            PhoneStatusMessage = $"✓ Formule telefono aggiornate in {updated} foglio/i.";
        });

        AddLog(PhoneStatusMessage);
    }

    // ── Helpers ──────────────────────────────────────────────────────
    private void AddLog(string message)
    {
        LogMessages.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
        while (LogMessages.Count > 50) LogMessages.RemoveAt(0);
        OnPropertyChanged(nameof(HasGlobalStatus));
    }

    private static TopLevel? GetTopLevel()
    {
        if (Avalonia.Application.Current?.ApplicationLifetime is
            Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
            return desktop.MainWindow;
        return null;
    }

    // ── INotifyPropertyChanged ───────────────────────────────────────
    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

// ── Minimal async command ────────────────────────────────────────────
public class AsyncRelayCommand : ICommand
{
    private readonly Func<Task> _execute;
    private readonly Func<bool>? _canExecute;
    private bool _isRunning;

    public AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null)
    {
        _execute = execute;
        _canExecute = canExecute;
    }

    public bool CanExecute(object? parameter) => !_isRunning && (_canExecute?.Invoke() ?? true);

    public async void Execute(object? parameter)
    {
        if (!CanExecute(parameter)) return;
        _isRunning = true;
        RaiseCanExecuteChanged();
        try   { await _execute(); }
        finally
        {
            _isRunning = false;
            RaiseCanExecuteChanged();
        }
    }

    public event EventHandler? CanExecuteChanged;
    public void RaiseCanExecuteChanged() =>
        CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}
