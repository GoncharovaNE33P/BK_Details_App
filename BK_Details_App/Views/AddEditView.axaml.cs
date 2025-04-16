using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.VisualTree;
using BK_Details_App.Models;
using BK_Details_App.ViewModels;

namespace BK_Details_App;

public partial class AddEditView : UserControl
{
    public AddEditView()
    {
        InitializeComponent();
        DataContext = new AddEditVM();
    }

    public AddEditView(Category category, Groups group)
    {
        InitializeComponent();
        DataContext = new AddEditVM(category, group);
    }

    public AddEditView(Materials material)
    {
        InitializeComponent();
        DataContext = new AddEditVM(material);
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