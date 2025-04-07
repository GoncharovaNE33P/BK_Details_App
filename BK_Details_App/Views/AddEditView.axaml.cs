using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
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

    public AddEditView(Category category, Groups group, Materials material)
    {
        InitializeComponent();
        DataContext = new AddEditVM(category, group, material);
    }
}