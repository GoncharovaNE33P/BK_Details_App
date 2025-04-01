using Aspose.Cells;
using Avalonia.Styling;
using BK_Details_App.Models;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using MsBox.Avalonia;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reactive.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BK_Details_App.ViewModels
{
    internal class DetailsVM : ViewModelBase
    {
        #region properties
        List<Materials> _materialsList = new();
        public List<Materials> MaterialsList 
        { 
            get => _materialsList;
            set => this.RaiseAndSetIfChanged(ref _materialsList, value);
        }
        List<Materials> _filteredMaterials = new();
        public List<Materials> FilteredMaterials { get => _filteredMaterials; set => this.RaiseAndSetIfChanged(ref _filteredMaterials, value); }

        List<Groups> _groupsList = new();
        public List<Groups> GroupsList { get => _groupsList; set => this.RaiseAndSetIfChanged(ref _groupsList, value); }
        Groups _selectedGroup = new();
        public Groups SelectedGroup { get => _selectedGroup; set { this.RaiseAndSetIfChanged(ref _selectedGroup, value); FilterMaterials(); FilterCategories(); } }

        List<Category> _categoriesList = new();
        public List<Category> CategoriesList { get => _categoriesList; set => this.RaiseAndSetIfChanged(ref _categoriesList, value); }
        List<Category> _filteredCategories = new();
        public List<Category> FilteredCategories { get => _filteredCategories; set => this.RaiseAndSetIfChanged(ref _filteredCategories, value); }
        Category _selectedCategory = new();
        public Category SelectedCategory { get => _selectedCategory; set { this.RaiseAndSetIfChanged(ref _selectedCategory, value); FilterMaterials(); } }

        string _searchMaterials;
        public string SearchMaterials { get { return _searchMaterials; } set { _searchMaterials = value; FilterMaterials(); } }

        bool _isAscending = false;
        public bool IsAscending
        {
            get => _isAscending;
            set
            {
                this.RaiseAndSetIfChanged(ref _isAscending, value);
                FilterMaterials();
            }
        }

        List<string> _favs = new();
        public List<string> Favs { get => _favs; set => this.RaiseAndSetIfChanged(ref _favs, value); }
        #endregion

        public DetailsVM()
        {
            try
            {
                ReadFromExcelFile();
                SelectedGroup = _groupsList[0];
                SelectedCategory = _categoriesList.Where(x => x.GroupNavigation == SelectedGroup).FirstOrDefault();
                ReadFavorites();
            }
            catch(Exception ex)
            {
                MessageBoxManager.GetMessageBoxStandard("Ошибка", ex.Message, MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
            }       
        }

        #region filters
        void FilterMaterials()
        {
            try
            {
                FilteredMaterials = MaterialsList.ToList();

                if (!string.IsNullOrWhiteSpace(_searchMaterials))
                {
                    FilteredMaterials = FilteredMaterials.Where(x => x.Name.ToLower().Contains(_searchMaterials.ToLower())).ToList();
                }

                if (_selectedGroup != null)
                {
                    FilteredMaterials = FilteredMaterials.Where(x => x.GroupNavigation == SelectedGroup).ToList();
                }

                if (_selectedCategory != null)
                {
                    FilteredMaterials = FilteredMaterials.Where(x => x.CategoryNavigation == SelectedCategory).ToList();
                }

                if (!_isAscending)
                {
                    FilteredMaterials = new(
                        FilteredMaterials.OrderBy(x => x.Name)
                    );
                }
                else
                {
                    FilteredMaterials = new(
                        FilteredMaterials.OrderByDescending(x => x.Name)
                    );
                }
            }
            catch(Exception ex)
            {
                MessageBoxManager.GetMessageBoxStandard("Ошибка", ex.Message, MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
            }
        }

        void FilterCategories()
        {
            try
            {
                if (SelectedGroup != null)
                {
                    FilteredCategories = new(CategoriesList.Where(x => x.GroupNavigation == SelectedGroup));
                }
                else
                {
                    FilteredCategories = new(CategoriesList);
                }

                SelectedCategory = FilteredCategories.FirstOrDefault();
            }
            catch (Exception ex)
            {
                MessageBoxManager.GetMessageBoxStandard("Ошибка", ex.Message, MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
            }
        }
        #endregion

        void ReadFromExcelFile()
        {
            // Загрузить файл Excel
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Materials", "materials.xlsx");
            Workbook wb = new Workbook(filePath);

            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;
            Random random = new Random();
            // Перебрать все рабочие листы
            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {
                // Получить рабочий лист, используя его индекс
                Worksheet worksheet = collection[worksheetIndex];
                if (worksheet.Name == "Крепёж" || worksheet.Name == "Электромонт") continue;
                Groups groups = new Groups() 
                { 
                    GroupIdNumber = random.Next(1, 1000), 
                    Name = worksheet.Name 
                };
                _groupsList.Add(groups);

                // Получить количество строк и столбцов
                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;
                int lastColoredRow = -1; // Запоминаем индекс последней добавленной ячейки с жирным шрифтом

                // Цикл по строкам
                for (int i = 1; i < rows; i++)
                {
                    // Перебрать каждый столбец в выбранной строке
                    for (int j = 1; j < cols; j++) //начинаем со столбца с наименованием
                    {
                        Cell cell = worksheet.Cells[i, j];
                        Aspose.Cells.Style style = cell.GetStyle();

                        if (cell.StringValue == "") break;

                        if (style.Font.IsBold is true) 
                        {
                            // Если предыдущая ячейка в той же колонке тоже была с жирным шрифтом, удаляем её
                            if (lastColoredRow == cell.Row - 1)
                            {
                                _categoriesList.RemoveAt(_categoriesList.Count - 1);
                            }

                            Category category = new Category() 
                            {
                                CategoryId = random.Next(1, 1000),
                                Name = cell.StringValue,
                                GroupNavigation = _groupsList[_groupsList.Count - 1],
                                Group = _groupsList[_groupsList.Count - 1].GroupIdNumber
                            };
                            _categoriesList.Add(category);

                            lastColoredRow = cell.Row; // Запоминаем текущую строку
                            break;
                        }
                        else
                        {
                            Materials materials = new Materials()
                            {
                                IdNumber = worksheet.Cells[i, j - 1].StringValue == "" ? random.Next(1, 1000) : worksheet.Cells[i, j - 1].IntValue, //в предыдущей ячейке содержится номер
                                Name = cell.StringValue,
                                Measurement = worksheet.Cells[i, j + 1].StringValue,
                                Analogs = string.IsNullOrWhiteSpace(worksheet.Cells[i, j + 2].StringValue) ? "Аналогов нет" : worksheet.Cells[i, j + 2].StringValue,
                                Note = string.IsNullOrWhiteSpace(worksheet.Cells[i, j + 3].StringValue) ? "Примечание отсутствует" : worksheet.Cells[i, j + 3].StringValue,
                                GroupNavigation = _groupsList[_groupsList.Count - 1],
                                Group = _groupsList[_groupsList.Count - 1].GroupIdNumber,
                                CategoryNavigation = _categoriesList[_categoriesList.Count - 1],
                                Category = _categoriesList[_categoriesList.Count - 1].CategoryId
                            };
                            _materialsList.Add(materials);
                            break;
                        }
                    }
                    // переходим на следующую строку
                }
                // переходим на следующий лист
            }

        }

        public void ReadFavorites()
        {
            string filePath = "fav.bin";

            if (File.Exists(filePath))
            {
                using (BinaryReader reader = new BinaryReader(File.Open(filePath, FileMode.Open)))
                {
                    while (reader.BaseStream.Position < reader.BaseStream.Length)
                    {
                        string str = reader.ReadString();
                        if (!Favs.Contains(str)) Favs.Add(str);
                    }
                }
            }
            else
            {
                Console.WriteLine("Файл не найден.");
            }
        }

        public void AddToFavorite(string _material)
        {
            string filePath = "fav.bin";

            if (Favs.Any(x => x == _material))
            {
                return; //!!!!!!!!!!!!!!!!!!!!!!!!!!сказать что уже в избранном
            }
            else
            {
                using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    fs.Seek(0, SeekOrigin.End); // Переход в конец файла перед записью

                    using (BinaryWriter writer = new BinaryWriter(fs))
                    {
                        writer.Write(_material);
                    }
                }
            }

            ReadFavorites();
            MainWindowViewModel.Instance.Us = new();
        }
    }
}
