using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Subjects;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using BK_Details_App.Models;
using MsBox.Avalonia;
using ReactiveUI;

namespace BK_Details_App.ViewModels
{
    internal class FavouritesVM : ViewModelBase
    {
        List<Materials> _favsList = new();
        public List<Materials> FavsList
        {
            get => _favsList;
            set => this.RaiseAndSetIfChanged(ref _favsList, value);
        }

        List<Materials> _filteredFavs = new();
        public List<Materials> FilteredFavs { get => _filteredFavs; set => this.RaiseAndSetIfChanged(ref _filteredFavs, value); }

        string _searchFavs;
        public string SearchFavs { get { return _searchFavs; } set { _searchFavs = value; FilterFavs(); } }

        bool _isAscFavs = false;
        public bool IsAscFavs { get => _isAscFavs; set { this.RaiseAndSetIfChanged(ref _isAscFavs, value); FilterFavs(); } }

        Groups _selectedGrFavs = new();
        public Groups SelectedGrFavs 
        { 
            get 
            { 
                if (_selectedGrFavs.Name is null) 
                    return _groupsList[0]; 
                else return _selectedGrFavs; 
            } 
            set { this.RaiseAndSetIfChanged(ref _selectedGrFavs, value); FilterFavs(); } }

        int _countItemsFavs = 0;
        public int CountItemsFavs { get => _countItemsFavs; set => this.RaiseAndSetIfChanged(ref _countItemsFavs, value); }

        int _countItemsFileFavs = 0;
        public int CountItemsFileFavs { get => _countItemsFileFavs; set => this.RaiseAndSetIfChanged(ref _countItemsFileFavs, value); }

        bool _nothingFound = false;
        public bool NothingFound { get => _nothingFound; set => this.RaiseAndSetIfChanged(ref _nothingFound, value); }

        List<Groups> _groupsList = new();
        public List<Groups> GroupsList { get => _groupsList; set => this.RaiseAndSetIfChanged(ref _groupsList, value); }

        public DetailsVM DetailsVMObj => new DetailsVM();

        public FavouritesVM()
        {
            FilteredFavs = GetMaterials();
            GroupsList = [new Groups() { Name = "Все группы" }, .. DetailsVMObj.GroupsList];
            CountItemsFavs = FilteredFavs.Count;
            CountItemsFileFavs = FilteredFavs.Count;
            FilterFavs();
        }

        List<Materials> GetMaterials()
        {
            List<string> buf = [.. DetailsVMObj.ReadFavorites()];
            FavsList = DetailsVMObj.MaterialsList.Where(x => buf.Contains(x.Name)).ToList();
            return FavsList;
        }

        void FilterFavs()
        {
            try
            {
                FilteredFavs = FavsList.ToList();
                if (_selectedGrFavs.Name == "Все категории")
                {
                    FilteredFavs = FavsList.ToList();
                }
                


                if (!string.IsNullOrWhiteSpace(SearchFavs))
                {
                    FilteredFavs = FilteredFavs.Where(x => x.Name.ToLower().Contains(_searchFavs.ToLower())).ToList();
                }

                if (_selectedGrFavs.Name != null && _selectedGrFavs.Name != "Все группы")
                {
                    FilteredFavs = FilteredFavs.Where(x => x.GroupNavigation.Name == SelectedGrFavs.Name).ToList();
                }

                if (!_isAscFavs)
                {
                    FilteredFavs = new(
                        FilteredFavs.OrderBy(x => x.Name)
                    );
                }
                else
                {
                    FilteredFavs = new(
                        FilteredFavs.OrderByDescending(x => x.Name)
                    );
                }

                CountItemsFavs = FilteredFavs.Count();
                if (CountItemsFavs == 0)
                    NothingFound = true;
                else
                    NothingFound = false;
            }
            catch (Exception ex)
            {
                MessageBoxManager.GetMessageBoxStandard("Ошибка", ex.Message, MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
            }
        }

        public void RemoveFromFavorite(string _material)
        {
            string _filePath = "test.xlsx";

            Workbook _workbook = new Workbook(_filePath);
            string _sheetName = "Избранное";
            Worksheet _sheet = _workbook.Worksheets[_sheetName];
            bool _foundFavForDel = false;
            foreach (Cell _currentCell in _sheet.Cells)
            {
                if (_currentCell.StringValue == _material)
                {
                    _sheet.Cells.DeleteRow(_currentCell.Row);
                    _foundFavForDel = true;
                }
            }

            if (_foundFavForDel)
            {
                _workbook.Save(_filePath);
            }
            else
            {
                DetailsVMObj.ShowError("Внимание", "Материал для удаления не найден");
            }

            FavsList = GetMaterials();
            CountItemsFileFavs = FavsList.Count;
            FilterFavs();
        }

        public void ToDetailsView()
        {
            MainWindowViewModel.Instance.Us = new DetailsView();
        }
    }
}
