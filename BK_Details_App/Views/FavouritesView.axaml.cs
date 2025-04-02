using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using BK_Details_App.ViewModels;

namespace BK_Details_App;

public partial class FavouritesView : UserControl
{
    public FavouritesView()
    {
        InitializeComponent();
        DataContext = new FavouritesVM();
    }
}