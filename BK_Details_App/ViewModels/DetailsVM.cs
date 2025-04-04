using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reactive.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using Avalonia.Controls;
using Avalonia.Platform;
using BK_Details_App.Models;
using DynamicData;
using ExcelDataReader;
using MsBox.Avalonia;
using MsBox.Avalonia.Enums;
using ReactiveUI;
using Tmds.DBus.Protocol;
using FuzzySharp;
using FuzzySharp.PreProcess;
using FuzzySharp.SimilarityRatio.Scorer.Composite;
using FuzzySharp.SimilarityRatio;
using ClosedXML.Excel;
using Avalonia.Media;
using DocumentFormat.OpenXml.Spreadsheet;

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

        List<BK_Details_App.Models.Groups> _groupsList = new();
        public List<BK_Details_App.Models.Groups> GroupsList { get => _groupsList; set => this.RaiseAndSetIfChanged(ref _groupsList, value); }
        BK_Details_App.Models.Groups _selectedGroup = new();
        public BK_Details_App.Models.Groups SelectedGroup { get => _selectedGroup; set { this.RaiseAndSetIfChanged(ref _selectedGroup, value); FilterMaterials(); FilterCategories(); } }

        List<Category> _categoriesList = new();
        public List<Category> CategoriesList { get => _categoriesList; set => this.RaiseAndSetIfChanged(ref _categoriesList, value); }
        List<Category> _filteredCategories = new();
        public List<Category> FilteredCategories { get => _filteredCategories; set => this.RaiseAndSetIfChanged(ref _filteredCategories, value); }
        Category _selectedCategory = new();
        public Category SelectedCategory { get => _selectedCategory; set { this.RaiseAndSetIfChanged(ref _selectedCategory, value); FilterMaterials(); } }

        string _searchMaterials;
        public string SearchMaterials { get { return _searchMaterials; } set { _searchMaterials = value; FilterMaterials(); } }

        bool _isAscending = false;
        public bool IsAscending { get => _isAscending; set { this.RaiseAndSetIfChanged(ref _isAscending, value); FilterMaterials(); } }

        List<string> _favs = new();
        public List<string> Favs { get => _favs; set => this.RaiseAndSetIfChanged(ref _favs, value); }

        ObservableCollection<PEZ> _collectionPEZs = new();
        public ObservableCollection<PEZ> CollectionPEZs { get => _collectionPEZs; set => this.RaiseAndSetIfChanged(ref _collectionPEZs, value); }

        List<PEZ> _listPEZs = new();
        public List<PEZ> ListPEZs { get => _listPEZs; set => this.RaiseAndSetIfChanged(ref _listPEZs, value); }

        private string _filePath;
        public string FilePath { get => _filePath; set => this.RaiseAndSetIfChanged(ref _filePath, value); }

        private string _nameFile;
        public string NameFile { get => _nameFile; set => this.RaiseAndSetIfChanged(ref _nameFile, value); }

        string _searchPEZs;
        public string SearchPEZs { get { return _searchPEZs; } set { _searchPEZs = value; FiltersPEZs(); } }

        bool _isAscendingPEZs = false;
        public bool IsAscendingPEZs { get => _isAscendingPEZs; set { this.RaiseAndSetIfChanged(ref _isAscendingPEZs, value); FiltersPEZs(); } }

        int _countItemsMaterials = 0;
        public int CountItemsMaterials { get => _countItemsMaterials; set => this.RaiseAndSetIfChanged(ref _countItemsMaterials, value); }

        int _countItemsFileMaterials = 0;
        public int CountItemsFileMaterials { get => _countItemsFileMaterials; set => this.RaiseAndSetIfChanged(ref _countItemsFileMaterials, value); }

        int _countItemsPEZs = 0;
        public int CountItemsPEZs { get => _countItemsPEZs; set => this.RaiseAndSetIfChanged(ref _countItemsPEZs, value); }

        int _countItemsFilePEZ = 0;
        public int CountItemsFilePEZ { get => _countItemsFilePEZ; set => this.RaiseAndSetIfChanged(ref _countItemsFilePEZ, value); }

        int _selectedFilter = 0;
        public int SelectedFilter { get => _selectedFilter; set { _selectedFilter = value; FiltersPEZs(); } }

        private Window? _addPEZWindow;
        private Window? _addWindow;

        #endregion

        public DetailsVM()
        {
            try
            {
                NameFile = "Тестовое ПЭ3";
                ReadFromExcelFile();
                SelectedGroup = _groupsList[0];
                SelectedCategory = _categoriesList.Where(x => x.GroupNavigation == SelectedGroup).FirstOrDefault();
                FilterMaterials();
                Favs = ReadFavorites();
                if (MainWindowViewModel.BaseListPEZs.Count > 0) CollectionPEZs.AddRange(MainWindowViewModel.BaseListPEZs);
                if (MainWindowViewModel.FilePath != null)
                {
                    FilePath = MainWindowViewModel.FilePath;
                    NameFile = FilePath.Split('\\')[FilePath.Split('\\').Length - 1];
                }
            }
            catch (Exception ex)
            {
                if (!Design.IsDesignMode)
                {
                    MessageBoxManager.GetMessageBoxStandard("DetailsVM: Ошибка", ex.Message + "\n" + ex.StackTrace, MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
                }
            }
        }

        #region Методы для вывода оконных сообщений
        public void ShowError(string title, string message)
        {
            MessageBoxManager.GetMessageBoxStandard(new MsBox.Avalonia.Dto.MessageBoxStandardParams
            {
                ContentTitle = title,
                ContentMessage = message,
                Icon = MsBox.Avalonia.Enums.Icon.Error,
                WindowIcon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.png"))),
                WindowStartupLocation = WindowStartupLocation.CenterOwner

            }).ShowAsync();
        }

        public void ShowSuccess(string title, string message)
        {
            MessageBoxManager.GetMessageBoxStandard(new MsBox.Avalonia.Dto.MessageBoxStandardParams
            {
                ContentTitle = title,
                ContentMessage = message,
                Icon = MsBox.Avalonia.Enums.Icon.Success,
                WindowIcon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.png"))),
                WindowStartupLocation = WindowStartupLocation.CenterOwner

            }).ShowAsync();
        }

        #endregion

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

                CountItemsMaterials = FilteredMaterials.Count();
            }
            catch (Exception ex)
            {
                MessageBoxManager.GetMessageBoxStandard("FilterMaterials: Ошибка", ex.Message, MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
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
                MessageBoxManager.GetMessageBoxStandard("FilterCategories: Ошибка", ex.Message, MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
            }
        }

        public void FiltersPEZs()
        {
            try
            {
                if (FilePath == null)
                {
                    ShowError("Ошибка!", "Сначала необходимо выбрать файл!");
                    return;
                }

                ListPEZs = MainWindowViewModel.BaseListPEZs;

                if (!string.IsNullOrWhiteSpace(_searchPEZs))
                {
                    ListPEZs = ListPEZs.Where(x => x.Name.ToLower().Contains(_searchPEZs.ToLower())).ToList();
                }

                if (!_isAscendingPEZs)
                {
                    ListPEZs = ListPEZs.OrderBy(x => x.Name).ToList();
                }
                else
                {
                    ListPEZs = ListPEZs.OrderByDescending(x => x.Name).ToList();
                }

                int count1 = 10;
                int count2 = 11;
                int count3 = 20;
                int count4 = 21;

                switch (_selectedFilter)
                {
                    case 0:
                        ListPEZs = ListPEZs.ToList();
                        break;

                    case 1:
                        ListPEZs = ListPEZs.Where(x => x.Quantity <= count1).ToList();
                        break;

                    case 2:
                        ListPEZs = ListPEZs.Where(x => x.Quantity >= count2 && x.Quantity < count3).ToList();
                        break;

                    case 3:
                        ListPEZs = ListPEZs.Where(x => x.Quantity >= count4).ToList();
                        break;
                }

                CollectionPEZs.Clear();
                CollectionPEZs.AddRange(ListPEZs);
                CountItemsPEZs = ListPEZs.Count();
            }
            catch (Exception ex)
            {
                ShowError("FilterPEZs: Ошибка!", ex.ToString());
            }
        }
        #endregion

        #region Методы считывания данных из Excel

        public void ReadFromExcelFile()
        {
            // Загрузить файл Excel
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Materials", "materials.xlsx");
            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(filePath);

            // Получить все рабочие листы
            WorksheetCollection collection = wb.Worksheets;
            Random random = new Random();
            // Перебрать все рабочие листы
            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {
                // Получить рабочий лист, используя его индекс
                Aspose.Cells.Worksheet worksheet = collection[worksheetIndex];

                BK_Details_App.Models.Groups groups = new BK_Details_App.Models.Groups()
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
                        Aspose.Cells.Cell cell = worksheet.Cells[i, j];
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

            CountItemsMaterials = MaterialsList.Count();
            CountItemsFileMaterials = MaterialsList.Count();
            MainWindowViewModel.AllMaterials = MaterialsList;
        }

        public async Task OpenFileAsync()
        {
            try
            {
                OpenFileDialog? dialog = new OpenFileDialog
                {
                    Title = "Выберите файл Excel",
                    Filters = {
                        new FileDialogFilter { Name = "CSV Files", Extensions = { "csv" } },
                        new FileDialogFilter { Name = "Excel Files", Extensions = { "xls", "xlsx" } }
                    }
                };

                string[]? files = await dialog.ShowAsync(new Window());
                if (files == null || files.Length == 0) return;

                MainWindowViewModel.FilePath = files[0];                
                FilePath = MainWindowViewModel.FilePath;
                NameFile = FilePath.Split('\\')[FilePath.Split('\\').Length - 1];
                if (FilePath.EndsWith(".csv")) LoadCsv(FilePath);
                else LoadExcel(FilePath);
            }
            catch (Exception ex)
            {
                ShowError("OpenFileAsync: Ошибка!", ex.ToString());
            }
        }

        private void LoadExcel(string filePath)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using FileStream? stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                using IExcelDataReader? reader = ExcelReaderFactory.CreateReader(stream);

                var result = reader.AsDataSet();
                var table = result.Tables[0];

                
                foreach (DataRow row in table.Rows.Cast<DataRow>().Skip(1))
                {
                    if (row[2].ToString() == "метка" || row[2].ToString() == "заземление") continue;
                    else
                    {
                        MainWindowViewModel.BaseListPEZs.Add(
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

                CollectionPEZs.Clear();
                ListPEZs.Clear();
                ListPEZs = MainWindowViewModel.BaseListPEZs;
                CollectionPEZs.AddRange(ListPEZs);

                CountItemsFilePEZ = ListPEZs.Count();
                CountItemsPEZs = ListPEZs.Count();

                MatchPEZMaterials();
            }
            catch (Exception ex)
            {
                ShowError("LoadExcel: Ошибка!", ex.ToString());
            }
        }

        private void LoadCsv(string filePath)
        {
            try
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
                        MainWindowViewModel.BaseListPEZs.Add(
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

                CollectionPEZs.Clear();
                ListPEZs.Clear();
                ListPEZs = MainWindowViewModel.BaseListPEZs;
                CollectionPEZs.AddRange(ListPEZs);

                CountItemsFilePEZ = ListPEZs.Count();
                CountItemsPEZs = ListPEZs.Count();

                MatchPEZMaterials();
            }
            catch (Exception ex)
            {
                ShowError("LoadCsv: Ошибка!", ex.ToString());
            }
        }

        public void ExportToExcel()
        {
            if (FilePath == null)
            {
                ShowError("Ошибка!", "Сначала необходимо выбрать файл!");
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderPath = Path.Combine(desktopPath, "REPORT_Comparison");
            Directory.CreateDirectory(folderPath);

            string filePath = Path.Combine(folderPath, "REPORT.xlsx");

            XLWorkbook workbook;
            workbook = new XLWorkbook();

            string sheetName = "REPORT";
            var sheet = workbook.Worksheets.Contains(sheetName)
            ? workbook.Worksheet(sheetName)
            : workbook.AddWorksheet(sheetName);

            //Определяем первую пустую строку
            sheet.Cell(1, 1).Value = NameFile;
            sheet.Cell(1, 1).Style.Font.Bold = true;
            sheet.Cell(1, 1).Style.Font.FontSize = 14;
            sheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Cell(1, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell(1, 1).Style.Border.OutsideBorderColor = XLColor.Black;
            sheet.Cell(1, 2).Value = "Материалы";
            sheet.Cell(1, 2).Style.Font.Bold = true;
            sheet.Cell(1, 2).Style.Font.FontSize = 14;
            sheet.Cell(1, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Cell(1, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            sheet.Cell(1, 2).Style.Border.OutsideBorderColor = XLColor.Black;
            int lastRow = sheet.LastRowUsed()?.RowNumber() + 1 ?? 1;
            
            foreach (PEZ pez in MainWindowViewModel.BaseListPEZs)
            {
                sheet.Cell(lastRow, 1).Value = pez.Name;
                sheet.Cell(lastRow, 2).Value = pez.Matched;
                sheet.Cell(lastRow, 1).Style.Fill.BackgroundColor = XLColor.FromHtml(pez.Color);
                sheet.Cell(lastRow, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                sheet.Cell(lastRow, 1).Style.Border.OutsideBorderColor = XLColor.Black;
                sheet.Cell(lastRow, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                sheet.Cell(lastRow, 2).Style.Border.OutsideBorderColor = XLColor.Black;
                lastRow++;
            }            

            //Сохраняем файл
            workbook.SaveAs(filePath);

            ShowSuccess("Успех!", "Отчёт выгружен в файл Excel!");
        }

        #endregion

        #region Методы связанные с избранными материалами

        public List<string> ReadFavorites()
        {
            List<string> values = new List<string>();
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Materials", "test.xlsx");

            if (!File.Exists(filePath))
                //throw new FileNotFoundException("Файл не найден в ReadFavorites", filePath);
                return values;
            else
            {
                Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filePath);
                Aspose.Cells.Worksheet sheet = workbook.Worksheets["Избранное"];

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
        }

        public async Task AddToFavorite(string _material)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Materials", "test.xlsx");

            if (Favs.Any(x => x == _material))
            {
                MessageBoxManager.GetMessageBoxStandard("Внимание", _material + " уже в избранном", MsBox.Avalonia.Enums.ButtonEnum.Ok, MsBox.Avalonia.Enums.Icon.Error).ShowAsync();
                return;
            }
            else
            {
                XLWorkbook workbook;
                if (File.Exists(filePath))
                {
                    workbook = new XLWorkbook(filePath);
                }
                else
                {
                    workbook = new XLWorkbook();
                }

                string sheetName = "Избранное";
                var sheet = workbook.Worksheets.Contains(sheetName)
                ? workbook.Worksheet(sheetName)
                : workbook.AddWorksheet(sheetName);

                //Определяем первую пустую строку
                int lastRow = sheet.LastRowUsed()?.RowNumber() + 1 ?? 1;
                sheet.Cell(lastRow, 1).Value = _material;
                
                //Сохраняем файл
                workbook.SaveAs(filePath);
            }

            //ReadFavorites();
            Favs = ReadFavorites();
            ShowSuccess("Успех", $"{_material} добавлен в избранное");
        }

        public void ToFavouritesView()
        {
            MainWindowViewModel.Instance.Us = new FavouritesView();
        }

        #endregion

        #region Методы открытия окна добавления/редактирования ПЭЗ и метод удаления ПЭЗ

        public void ShowAddPEZ()
        {
            try
            {
                int id = 0;
                if (CollectionPEZs.Count == 0)
                {
                    ShowError("Ошибка!", "Сначала необходимо выбрать файл!");
                    return;
                }

                if (_addPEZWindow != null) return;

                var viewModel = new AddEditPEZVM(id, FilePath);

                _addPEZWindow = new Window
                {
                    MinHeight = 600,
                    MinWidth = 1500,
                    Content = new AddEditPEZView { DataContext = viewModel },
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    CanResize = false,
                    Title = "BK_Details_App",
                    Icon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.png")))
                };

                viewModel.CloseAction = () =>
                {
                    if (viewModel.CloseAction != null)
                    {
                        _addPEZWindow.Closing -= PreventClosing;
                        _addPEZWindow.Close();
                        _addPEZWindow = null;
                    }
                };

                _addPEZWindow.Closing += PreventClosing;

                _addPEZWindow.Show();
            }
            catch (Exception ex)
            {
                ShowError("ShowAddPEZ: Ошибка!", ex.ToString());
            }
        }

        public void ShowEditPEZ(int id)
        {
            try
            {
                if (CollectionPEZs.Count == 0)
                {
                    ShowError("Ошибка!", "Сначала необходимо выбрать файл!");
                    return;
                }

                if (_addPEZWindow != null) return;

                var viewModel = new AddEditPEZVM(id, FilePath);

                _addPEZWindow = new Window
                {
                    MinHeight = 600,
                    MinWidth = 1500,
                    Content = new AddEditPEZView { DataContext = viewModel },
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    CanResize = false,
                    Title = "BK_Details_App",
                    Icon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.png")))
                };

                viewModel.CloseAction = () =>
                {
                    if (viewModel.CloseAction != null)
                    {
                        _addPEZWindow.Closing -= PreventClosing;
                        _addPEZWindow.Close();
                        _addPEZWindow = null;
                    }
                };

                _addPEZWindow.Closing += PreventClosing;

                _addPEZWindow.Show();
            }
            catch (Exception ex)
            {
                ShowError("ShowEditPEZ: Ошибка!", ex.ToString());
            }
        }

        private void PreventClosing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }

        public async void DeletePEZ(int id)
        {
            try
            {
                string Messege = "Вы уверенны, что хотите удалить данное ПЭЗ?";
                ButtonResult result = await MessageBoxManager.GetMessageBoxStandard("Сообщение с уведомлением об удалении!", Messege, ButtonEnum.YesNo).ShowAsync();
                PEZ? PEZRemove = CollectionPEZs.FirstOrDefault(p => p.IdNumber == id);                

                switch (result)
                {
                    case ButtonResult.Yes:
                        {
                            if (PEZRemove != null)
                            {
                                CollectionPEZs.Remove(PEZRemove);
                                ListPEZs = CollectionPEZs.ToList();
                                CountItemsFilePEZ = ListPEZs.Count();
                                CountItemsPEZs = ListPEZs.Count();
                            }
                            Messege = "ПЭЗ удалён!";
                            ShowSuccess("Сообщение с уведомлением об удалении!", Messege);
                            break;
                        }
                    case ButtonResult.No:
                        {
                            Messege = "Удаление отменено!";
                            ShowSuccess("Сообщение с уведомлением об удалении!", Messege);
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                ShowError("DeletePEZ: Ошибка!", ex.ToString());
            }
        }

        #endregion

        public void MatchPEZMaterials()
        {
            string[]? materialNames = MaterialsList.Select(x => x.Name).ToArray(); // список имен в массив, потому что вроде как ExtractOne лучше/в принципе работает с массивами
            FuzzySharp.SimilarityRatio.Scorer.IRatioScorer? scorer = ScorerCache.Get<WeightedRatioScorer>();
            
            Parallel.ForEach(CollectionPEZs, _element =>
            {
                FuzzySharp.Extractor.ExtractedResult<string>? match = FuzzySharp.Process.ExtractOne(_element.Name, materialNames, s => s, scorer);

                _element.Matched = materialNames[match.Index];

                _element.Color = match.Score switch
                {
                    100 => "#66C190",
                    > 70 => "#FFE666",
                    _ => "#FF9166"
                };
            });
        }

        public void AddMaterial(Materials material)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Materials", "materials.xlsx");
            XLWorkbook wb = new XLWorkbook(filePath);

            foreach (var ws in wb.Worksheets)
            {
                foreach (var cell in ws.CellsUsed(c => c.HasFormula))
                {
                    if (string.IsNullOrWhiteSpace(cell.FormulaA1))
                        cell.Clear(); // Или cell.Value = null;
                }
            }

            var currentWorksheet = wb.Worksheet(SelectedGroup.Name);

            int neededRow = -1;
            int lastRow = currentWorksheet.LastRowUsed()?.RowNumber() ?? 0;
            for (int i = 1; i < lastRow; i++)
            {
                if (currentWorksheet.Cell(i, 2).GetString() == SelectedCategory.Name)
                {
                    neededRow = i + MaterialsList.Count(x => x.CategoryNavigation.Name == SelectedCategory.Name) + 1;
                    currentWorksheet.Row(neededRow).InsertRowsAbove(1);
                }

                if (i == neededRow)
                {
                    currentWorksheet.Cell(i, 1).Value = MaterialsList.Where(x => x.GroupNavigation.Name == SelectedGroup.Name).Max(x => x.IdNumber) + 1;
                    currentWorksheet.Cell(i, 2).Value = material.Name;
                    currentWorksheet.Cell(i, 3).Value = material.Measurement;
                    currentWorksheet.Cell(i, 4).Value = material.Analogs;
                    currentWorksheet.Cell(i, 5).Value = material.Note;
                    break;
                }
            }
            
            wb.SaveAs(filePath);
        }

        public void ShowAddMaterials()
        {
            try
            {
                int id = 0;

                if (_addWindow != null) return;

                var viewModel = new AddEditVM();

                _addWindow = new Window
                {
                    MinHeight = 600,
                    MinWidth = 1500,
                    Content = new AddEditView { DataContext = viewModel },
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    CanResize = false,
                    Title = "BK_Details_App",
                    Icon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.png")))
                };

                viewModel.CloseAction = () =>
                {
                    if (viewModel.CloseAction != null)
                    {
                        _addWindow.Closing -= PreventClosing;
                        _addWindow.Close();
                        _addWindow = null;
                    }
                };

                _addWindow.Closing += PreventClosing;

                _addWindow.Show();
            }
            catch (Exception ex)
            {
                ShowError("ShowAdd: Ошибка!", ex.ToString());
            }
        }
    }
}
