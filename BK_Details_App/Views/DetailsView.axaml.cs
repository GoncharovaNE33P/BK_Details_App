using Avalonia;
using Avalonia.Controls;
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
}