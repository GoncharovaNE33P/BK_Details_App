using System.Collections.Generic;
using Avalonia.Controls;
using BK_Details_App.Models;
using ReactiveUI;
using BK_Details_App.Models;

namespace BK_Details_App.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {        
        #region Properties

        public static MainWindowViewModel Instance; // создаем объект для обращения к другим объектам данного класса

        private static List<Materials> _allMaterials = new();
        internal static List<Materials> AllMaterials { get => _allMaterials; set => _allMaterials = value; }
        public MainWindowViewModel()
        {
            Instance = this;
        }

        static List<PEZ> _baseListPEZs = new();
        internal static List<PEZ> BaseListPEZs { get => _baseListPEZs; set => _baseListPEZs = value; }

        static string _filePath;

        public static string FilePath { get => _filePath; set => _filePath = value; }

        UserControl _us = new DetailsView();

        public UserControl Us //UserControl для организации навигации по страницам
        {
            get => _us;
            set => this.RaiseAndSetIfChanged(ref _us, value);
        }        

        #endregion
    }
}
