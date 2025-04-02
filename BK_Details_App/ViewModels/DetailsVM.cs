using Aspose.Cells;
using BK_Details_App.Models;
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
using ExcelDataReader;
using Avalonia.Controls;
using System.Data;
using DynamicData;

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

        ObservableCollection<PEZ> _collectionPEZs = new();
        public ObservableCollection<PEZ> CollectionPEZs { get => _collectionPEZs; set => this.RaiseAndSetIfChanged(ref _collectionPEZs, value); }

        List<PEZ> _listPEZs = new();
        public List<PEZ> ListPEZs { get => _listPEZs; set => this.RaiseAndSetIfChanged(ref _listPEZs, value); }

        private string filePath;
        public string FilePath { get => filePath; set => this.RaiseAndSetIfChanged(ref filePath, value); }

        string _searchPEZs;
        public string SearchPEZs { get { return _searchPEZs; } set { _searchPEZs = value; FiltersPEZs(); } }

        bool _isAscendingPEZs = false;
        public bool IsAscendingPEZs
        {
            get => _isAscendingPEZs;
            set
            {
                this.RaiseAndSetIfChanged(ref _isAscendingPEZs, value);
                FiltersPEZs();
            }
        }

        #endregion

        public DetailsVM()
        {
            try
            {
                ReadFromExcelFile();
                SelectedGroup = _groupsList[0];
                SelectedCategory = _categoriesList.Where(x => x.GroupNavigation == SelectedGroup).FirstOrDefault();
                //ReadFavorites();
                //Favs = ReadFavorites();
            }
            catch (Exception ex)
            {
                if (!Design.IsDesignMode)
                {
                    MessageBoxManager.GetMessageBoxStandard("Ошибка", ex.Message, MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
                }
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

        public void FiltersPEZs()
        {
            if (FilePath.EndsWith(".csv")) LoadCsv(FilePath);
            else LoadExcel(FilePath);

            if (!string.IsNullOrWhiteSpace(_searchPEZs))
            {
                ListPEZs = ListPEZs.Where(x => x.Name.ToLower().Contains(_searchPEZs.ToLower())).ToList();
                CollectionPEZs.AddRange(ListPEZs);
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

        public async Task OpenFileAsync()
        {
            OpenFileDialog? dialog = new OpenFileDialog
            {
                Title = "Выберите файл Excel",
                Filters = {
            new FileDialogFilter { Name = "Excel Files", Extensions = { "xls", "xlsx" } },
            new FileDialogFilter { Name = "CSV Files", Extensions = { "csv" } }
        }
            };

            string[]? files = await dialog.ShowAsync(new Window());
            if (files == null || files.Length == 0) return;

            FilePath = files[0];
            if (FilePath.EndsWith(".csv")) LoadCsv(FilePath);
            else LoadExcel(FilePath);
        }

        private void LoadExcel(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using FileStream? stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using IExcelDataReader? reader = ExcelReaderFactory.CreateReader(stream);

            var result = reader.AsDataSet();
            var table = result.Tables[0];

            CollectionPEZs.Clear();
            foreach (DataRow row in table.Rows.Cast<DataRow>().Skip(1))
            {
                if (row[2].ToString() == "метка" || row[2].ToString() == "заземление") continue;
                else
                {
                    CollectionPEZs.Add(
                        new PEZ
                        {
                            IdNumber = int.TryParse(row[0]?.ToString(), out int id) ? id : 0,
                            Mark = row[1]?.ToString(),
                            Name = row[2]?.ToString(),
                            Quantity = int.TryParse(row[3]?.ToString(), out int quantity) ? quantity : 0
                        }
                    );
                }
            }
            ListPEZs.Clear();
            ListPEZs = CollectionPEZs.ToList();
        }

        private void LoadCsv(string filePath)
        {
            Encoding? encoding = Encoding.UTF8;

            byte[]? bytes = File.ReadAllBytes(filePath);
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) encoding = Encoding.UTF8;
            else encoding = Encoding.GetEncoding("Windows-1251");

            string[]? lines = File.ReadAllLines(filePath, encoding);
            CollectionPEZs.Clear();

            foreach (string? line in lines.Skip(1))
            {
                string[]? parts = line.Split(';');
                if (parts.Length < 4) continue;

                if (parts[2] == "метка" || parts[2] == "заземление") continue;
                else
                {
                    CollectionPEZs.Add(
                        new PEZ
                        {
                            IdNumber = int.TryParse(parts[0]?.ToString(), out int id) ? id : 0,
                            Mark = parts[1]?.ToString(),
                            Name = parts[2]?.ToString(),
                            Quantity = int.TryParse(parts[3]?.ToString(), out int quantity) ? quantity : 0
                        }
                    );
                }
            }
            ListPEZs.Clear();
            ListPEZs = CollectionPEZs.ToList();
        }


        /*public void ReadFavorites()
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
        }*/

        public List<string> ReadFavorites()
        {
            List<string> values = new List<string>();
            string filePath = "test.xlsx";

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Файл не найден", filePath);

            Workbook workbook = new Workbook(filePath);
            Worksheet sheet = workbook.Worksheets["Избранное"];

            if (sheet == null)
                throw new ArgumentException("Лист не найден");

            int rowCount = sheet.Cells.MaxDataRow;

            for (int i = 0; i <= rowCount; i++)
            {
                string cellValue = sheet.Cells[i, 0].StringValue; // Читаем первую колонку
                if (!string.IsNullOrEmpty(cellValue))
                    values.Add(cellValue);
            }

            return values;
        }

        public void AddToFavorite(string _material)
        {
            string filePath = "test.xlsx";

            if (Favs.Any(x => x == _material))
            {
                MessageBoxManager.GetMessageBoxStandard("Внимание", _material + " уже в избранном", MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
                return; //!!!!!!!!!!!!!!!!!!!!!!!!!!сказать что уже в избранном
            }
            else
            {
                /*using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    fs.Seek(0, SeekOrigin.End); // Переход в конец файла перед записью

                    using (BinaryWriter writer = new BinaryWriter(fs))
                    {
                        writer.Write(_material);
                    }
                }*/

                //Проверяем, существует ли файл
                Workbook workbook;
                if (File.Exists(filePath))
                {
                    workbook = new Workbook(filePath);
                }
                else
                {
                    workbook = new Workbook();
                }

                string sheetName = "Избранное";
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    int sheetIndex = workbook.Worksheets.Add();
                    sheet = workbook.Worksheets[sheetIndex];
                    sheet.Name = sheetName;
                }

                //Определяем первую пустую строку
                int lastRow = sheet.Cells.MaxDataRow + 1;
                sheet.Cells[lastRow, 0].PutValue(_material);

                //Сохраняем файл
                workbook.Save(filePath);
            }

            //ReadFavorites();
            Favs = ReadFavorites();
            MainWindowViewModel.Instance.Us = new FavouriteGroups();
        }
    }
}
