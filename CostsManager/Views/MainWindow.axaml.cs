using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace CostsManager.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    // Drag the frameless window by the title bar
    private void TitleBar_PointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            BeginMoveDrag(e);
    }

    private void BtnClose_Click(object? sender, RoutedEventArgs e) => Close();

    private void BtnMinimize_Click(object? sender, RoutedEventArgs e) =>
        WindowState = WindowState.Minimized;

    // Confirm on Enter in the year TextBox
    private void YearInput_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key == Key.Return)
        {
            if (DataContext is ViewModels.MainViewModel vm &&
                vm.AddYearSheetCommand.CanExecute(null))
                vm.AddYearSheetCommand.Execute(null);
        }
    }
}
