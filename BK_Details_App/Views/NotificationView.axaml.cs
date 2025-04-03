using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace BK_Details_App;

public partial class NotificationView : UserControl
{
    public NotificationView()
    {
        InitializeComponent();
    }

    public async Task ShowMessage(string message, int duration = 3000)
    {
        MessageText.Text = message;
        Opacity = 1;

        await Task.Delay(duration);
    }
}