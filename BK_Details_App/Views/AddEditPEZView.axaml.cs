using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using BK_Details_App.ViewModels;

namespace BK_Details_App;

public partial class AddEditPEZView : UserControl
{
    public AddEditPEZView()
    {
        InitializeComponent();
        DataContext = new AddEditPEZVM();
    }

    public AddEditPEZView(string filePath)
    {
        InitializeComponent();
        DataContext = new AddEditPEZVM(filePath);
    }

    public AddEditPEZView(int id, string filePath)
    {
        InitializeComponent();
        DataContext = new AddEditPEZVM(id, filePath);
    }
}