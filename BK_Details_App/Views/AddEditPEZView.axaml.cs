using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.VisualTree;
using BK_Details_App.ViewModels;

namespace BK_Details_App;

public partial class AddEditPEZView : UserControl
{
    public AddEditPEZView()
    {
        InitializeComponent();
        DataContext = new AddEditPEZVM();
    }

    public AddEditPEZView(int id, string filePath)
    {
        InitializeComponent();
        DataContext = new AddEditPEZVM(id, filePath);
    }

    private void TitleBar_PointerPressed(object? sender, PointerPressedEventArgs e)
    {
        // Получаем окно, в котором находится UserControl
        var window = this.GetVisualRoot() as Window;
        if (e.GetCurrentPoint(window).Properties.IsLeftButtonPressed)
            window?.BeginMoveDrag(e);
    }

    private void Minimize_Click(object? sender, RoutedEventArgs e)
    {
        (this.GetVisualRoot() as Window)!.WindowState = WindowState.Minimized;
    }
}