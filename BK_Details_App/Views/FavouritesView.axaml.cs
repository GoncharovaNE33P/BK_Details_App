using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Markup.Xaml;
using BK_Details_App.ViewModels;

namespace BK_Details_App;

public partial class FavouritesView : UserControl
{
    public FavouritesView(bool a)
    {
        InitializeComponent();
        DataContext = new FavouritesVM(a);
    }
    public FavouritesView()
    {
        InitializeComponent();
        DataContext = new FavouritesVM();
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