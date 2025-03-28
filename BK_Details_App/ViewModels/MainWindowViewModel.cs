using Avalonia.Controls;
using ReactiveUI;

namespace BK_Details_App.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        #region Properties
        public static MainWindowViewModel Instance; // создаем объект для обращения к другим объектам данного класса
        public MainWindowViewModel()
        {
            Instance = this;
        }

        UserControl _us = new DetailsView();

        public UserControl Us //UserControl для организации навигации по страницам
        {
            get => _us;
            set => this.RaiseAndSetIfChanged(ref _us, value);
        }
        #endregion
    }
}
