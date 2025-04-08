using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Markup.Xaml;
using BK_Details_App.ViewModels;

namespace BK_Details_App;

public partial class DetailsView : UserControl
{
    public DetailsView()
    {
        InitializeComponent();
        DataContext = new DetailsVM();
    }

    public DetailsView(bool s)
    {
        InitializeComponent();
        DataContext = new DetailsVM(s);
    }

    private void OnPointerWheelChanged(object? sender, PointerWheelEventArgs e)
    {
        if (sender is ScrollViewer scrollViewer)
        {
            scrollViewer.Offset = scrollViewer.Offset.WithX(scrollViewer.Offset.X - e.Delta.Y * 30);
            e.Handled = true;
        }
    }
}