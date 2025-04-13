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
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using Microsoft.Extensions.Logging;

namespace BK_Details_App.ViewModels
{
    public class DetailsVM : ViewModelBase
    {
        #region Properties

        List<Materials> _materialsList = new();
        public List<Materials> MaterialsList { get => _materialsList; set => this.RaiseAndSetIfChanged(ref _materialsList, value); }

        List<Materials> _filteredMaterials = new();
        public List<Materials> FilteredMaterials { get => _filteredMaterials; set => this.RaiseAndSetIfChanged(ref _filteredMaterials, value); }

        List<Models.Groups> _groupsList = new();
        public List<Models.Groups> GroupsList { get => _groupsList; set => this.RaiseAndSetIfChanged(ref _groupsList, value); }
        Models.Groups _selectedGroup = new();
        public Models.Groups SelectedGroup { get => _selectedGroup; set { this.RaiseAndSetIfChanged(ref _selectedGroup, value); FilterMaterials(); FilterCategories(); } }

        List<Models.Category> _categoriesList = new();
        public List<Models.Category> CategoriesList { get => _categoriesList; set => this.RaiseAndSetIfChanged(ref _categoriesList, value); }
        List<Models.Category> _filteredCategories = new();
        public List<Models.Category> FilteredCategories { get => _filteredCategories; set => this.RaiseAndSetIfChanged(ref _filteredCategories, value); }
        Models.Category _selectedCategory = new();
        public Models.Category SelectedCategory { get => _selectedCategory; set { this.RaiseAndSetIfChanged(ref _selectedCategory, value); FilterMaterials(); } }

        string _searchMaterials = "";
        public string SearchMaterials { get { return _searchMaterials; } set { _searchMaterials = value; FilterMaterials(); } }

        bool _isAscending = false;
        public bool IsAscending { get => _isAscending; set { this.RaiseAndSetIfChanged(ref _isAscending, value); FilterMaterials(); } }

        List<string> _favs = new();
        public List<string> Favs { get => _favs; set => this.RaiseAndSetIfChanged(ref _favs, value); }

        ObservableCollection<PEZ> _collectionPEZs = new();
        public ObservableCollection<PEZ> CollectionPEZs { get => _collectionPEZs; set => this.RaiseAndSetIfChanged(ref _collectionPEZs, value); }

        List<PEZ> _listPEZs = new();
        public List<PEZ> ListPEZs { get => _listPEZs; set => this.RaiseAndSetIfChanged(ref _listPEZs, value); }

        private string _filePath = "";
        public string FilePath { get => _filePath; set => this.RaiseAndSetIfChanged(ref _filePath, value); }

        private string _nameFile = "";
        public string NameFile { get => _nameFile; set => this.RaiseAndSetIfChanged(ref _nameFile, value); }

        string _searchPEZs = "";
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

        private void PreventClosing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }

        #endregion
        public string path = Path.Combine(Directory.GetCurrentDirectory(), "Materials", "test.xlsx");
        public DetailsVM(bool skipInit = false) 
        {            
            if (skipInit)
            {
                return;
            } 
        }
        public DetailsVM()
        {

            try
            {
                NameFile = "Тестовое ПЭ3";
                ReadFromExcelFile();
                SelectedGroup = _groupsList[0];
                SelectedCategory = _categoriesList.FirstOrDefault(x => x.GroupNavigation == SelectedGroup);
                FilterMaterials();

                Favs = ReadFavorites(path);
                if (MainWindowViewModel.BaseListPEZs.Count > 0) CollectionPEZs.AddRange(MainWindowViewModel.BaseListPEZs);
                if (MainWindowViewModel.FilePath != null)
                {
                    FilePath = MainWindowViewModel.FilePath;
                    NameFile = FilePath.Split('\\')[FilePath.Split('\\').Length - 1];
                    CountItemsFilePEZ = MainWindowViewModel.BaseListPEZs.Count();
                    CountItemsPEZs = MainWindowViewModel.BaseListPEZs.Count();
                    FiltersPEZs();
                }
            }
            catch (Exception ex)
            {
                if (!Design.IsDesignMode)
                {
                    ShowError("DetailsVM: Ошибка!", ex.ToString());
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
                WindowIcon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.ico"))),
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
                WindowIcon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.ico"))),
                WindowStartupLocation = WindowStartupLocation.CenterOwner

            }).ShowAsync();
        }

        #endregion

        #region Методы поиска, сортировки и фильтрации

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
                ShowError("FilterMaterials: Ошибка!", ex.ToString());
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
                ShowError("FilterCategories: Ошибка!", ex.ToString());
            }
        }

        public void FiltersPEZs()
        {
            try
            {
                if (FilePath == "")
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

        #region Методы импорта и экспорта данных Excel

        public void ReadFromExcelFile()
        {
            try
            {
                string filePath = Path.Combine(AppContext.BaseDirectory.Substring(0, AppContext.BaseDirectory.IndexOf("bin") - 1), "Materials", "materials.xlsx");
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook(filePath);
                WorksheetCollection collection = wb.Worksheets;
                Random random = new Random();

                _groupsList.Clear();
                _categoriesList.Clear();
                _materialsList.Clear();

                for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
                {
                    Aspose.Cells.Worksheet worksheet = collection[worksheetIndex];

                    var group = new BK_Details_App.Models.Groups()
                    {
                        GroupIdNumber = random.Next(1, 1000),
                        Name = worksheet.Name
                    };
                    _groupsList.Add(group);

                    Models.Category? lastCategory = null;
                    bool lastWasBold = false;
                    int rows = worksheet.Cells.MaxDataRow;
                    int cols = worksheet.Cells.MaxDataColumn;

                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= cols; j++)
                        {
                            var cell = worksheet.Cells[i, j];
                            var style = cell.GetStyle();
                            var value = cell.StringValue;

                            if (string.IsNullOrWhiteSpace(value))
                                break;

                            if (style.Font.IsBold)
                            {
                                if (lastWasBold && _categoriesList.Count > 0)
                                {
                                    // Удаляем предыдущую категорию, если была две жирные строки подряд
                                    _categoriesList.RemoveAt(_categoriesList.Count - 1);
                                }

                                var category = new Models.Category()
                                {
                                    CategoryId = random.Next(1, 1000),
                                    Name = value,
                                    GroupNavigation = group,
                                    Group = group.GroupIdNumber
                                };
                                _categoriesList.Add(category);
                                lastCategory = category;
                                lastWasBold = true;
                                break;
                            }
                            else if (lastCategory != null)
                            {
                                var material = new Materials()
                                {
                                    IdNumber = worksheet.Cells[i, j - 1].StringValue == "" ? random.Next(1, 1000) : worksheet.Cells[i, j - 1].IntValue,
                                    Name = cell.StringValue,
                                    Measurement = worksheet.Cells[i, j + 1].StringValue,
                                    Analogs = string.IsNullOrWhiteSpace(worksheet.Cells[i, j + 2].StringValue) ? "Аналогов нет" : worksheet.Cells[i, j + 2].StringValue,
                                    Note = string.IsNullOrWhiteSpace(worksheet.Cells[i, j + 3].StringValue) ? "Примечание отсутствует" : worksheet.Cells[i, j + 3].StringValue,
                                    GroupNavigation = group,
                                    Group = group.GroupIdNumber,
                                    CategoryNavigation = lastCategory,
                                    Category = lastCategory.CategoryId
                                };
                                _materialsList.Add(material);
                                lastWasBold = false;
                                break;
                            }
                        }
                    }
                }

                CountItemsMaterials = MaterialsList.Count();
                CountItemsFileMaterials = MaterialsList.Count();
                MainWindowViewModel.AllMaterials = MaterialsList;
            }

            catch (Exception ex)
            {
                //ShowError("ReadFromExcelFile: Ошибка!", ex.ToString());
                using ILoggerFactory factory = LoggerFactory.Create(builder => builder.AddConsole());
                Microsoft.Extensions.Logging.ILogger logger = factory.CreateLogger<Program>();
                logger.LogInformation($":::::EXCEPTION:::::::::::::::EXCEPTION:::::::::::::::EXCEPTION::::::::{ex.ToString()}.", "what");
            }
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
                FiltersPEZs();
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

                MainWindowViewModel.BaseListPEZs.Clear();
                foreach (DataRow row in table.Rows.Cast<DataRow>().Skip(1))
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
                
                ListPEZs = MainWindowViewModel.BaseListPEZs.ToList();
                CollectionPEZs.Clear();
                CollectionPEZs.AddRange(ListPEZs);

                CountItemsFilePEZ = ListPEZs.Count();
                CountItemsPEZs = ListPEZs.Count();
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

                MainWindowViewModel.BaseListPEZs.Clear();
                foreach (string? line in lines.Skip(1))
                {
                    string[]? parts = line.Split(';');
                    if (parts.Length < 4) continue;

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
                               
                ListPEZs = MainWindowViewModel.BaseListPEZs.ToList();
                CollectionPEZs.Clear();
                CollectionPEZs.AddRange(ListPEZs);

                CountItemsFilePEZ = ListPEZs.Count();
                CountItemsPEZs = ListPEZs.Count();
            }
            catch (Exception ex)
            {
                ShowError("LoadCsv: Ошибка!", ex.ToString());
            }
        }

        public void ExportToExcel()
        {
            try
            {
                if (FilePath == "")
                {
                    ShowError("Ошибка!", "Сначала необходимо выбрать файл!");
                    return;
                }
                if (!MainWindowViewModel.BaseListPEZs.Any(x => x.Color != null))
                {
                    ShowError("Ошибка!", "Сначала необходимо выполнить сравнение!");
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
            catch (Exception ex)
            {
                ShowError("ExportToExcel: Ошибка!", ex.ToString());
            }
        }

        #endregion

        #region Методы связанные с избранными материалами

        public List<string> ReadFavorites(string filePath)
        {
            try
            {
                List<string> values = new List<string>();                

                if (!File.Exists(filePath))                    
                    return values;
                else
                {
                    Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(filePath);
                    Aspose.Cells.Worksheet sheet = workbook.Worksheets["Избранное"];

                    if (sheet == null)
                    {
                        WorksheetCollection worksheets = workbook.Worksheets;
                        Aspose.Cells.Worksheet worksheet = worksheets.Add("Избранное");
                        workbook.Save(filePath);
                    }

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
            catch (Exception ex)
            {
                ShowError("ReadFavorites: Ошибка!", ex.ToString());
                return new List<string>();
            }
        }

        public void AddToFavorite(string _material)
        {
            try
            {
                string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Materials", "test.xlsx");

                if (Favs.Any(x => x == _material))
                {
                    ShowError("Внимание!", _material + " уже в избранном");
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
                Favs = ReadFavorites(path);
                ShowSuccess("Успех", $"{_material} добавлен в избранное");
            }
            catch (Exception ex)
            {
                ShowError("AddToFavorite: Ошибка!", ex.ToString());
            }
        }

        public void ToFavouritesView()
        {
            MainWindowViewModel.Instance.Us = new FavouritesView();
        }

        #endregion

        #region Методы работы с ПЭЗ

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
                    Icon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.ico")))
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
                    Icon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.ico")))
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

        public async void DeletePEZ(int id)
        {
            try
            {
                string Messege = "Вы уверенны, что хотите удалить данное ПЭЗ?";
                ButtonResult result = await MessageBoxManager.GetMessageBoxStandard("Сообщение с уведомлением об удалении!", Messege, ButtonEnum.YesNo).ShowAsync();
                PEZ? PEZRemove = MainWindowViewModel.BaseListPEZs.FirstOrDefault(p => p.IdNumber == id);                

                switch (result)
                {
                    case ButtonResult.Yes:
                        {
                            if (PEZRemove != null)
                            {
                                MainWindowViewModel.BaseListPEZs.Remove(PEZRemove);
                                CollectionPEZs.Clear();
                                CollectionPEZs.AddRange(MainWindowViewModel.BaseListPEZs);
                                ListPEZs = CollectionPEZs.ToList();
                                CountItemsFilePEZ = ListPEZs.Count();
                                CountItemsPEZs = ListPEZs.Count();

                                await SaveToFileAsync(FilePath);
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

        private async Task SaveToFileAsync(string filePath)
        {
            try
            {
                if (filePath.EndsWith(".csv"))
                {
                    await SaveToCsvAsync(filePath);
                }
                else if (filePath.EndsWith(".xls") || filePath.EndsWith(".xlsx"))
                {
                    SaveToExcel(filePath);
                }
            }
            catch (Exception ex)
            {
                ShowError("SaveToFileAsync: Ошибка!", ex.ToString());
            }
        }

        private async Task SaveToCsvAsync(string filePath)
        {
            try 
            {
                List<string>? lines = new List<string>
                {
                    "#;Метка;Имя;Количество"
                };

                foreach (PEZ? pez in MainWindowViewModel.BaseListPEZs)
                {
                    lines.Add($"{pez.IdNumber};{pez.Mark};{pez.Name};{pez.Quantity}");
                }

                await File.WriteAllLinesAsync(filePath, lines, Encoding.GetEncoding("Windows-1251"));
            }
            catch (Exception ex)
            {
                ShowError("SaveToCsvAsync: Ошибка!", ex.ToString());
            }
        }

        private void SaveToExcel(string filePath)
        {
            try
            {
                using XLWorkbook workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add(NameFile);

                // Заголовки
                worksheet.Cell(1, 1).Value = "#";
                worksheet.Cell(1, 2).Value = "Метка";
                worksheet.Cell(1, 3).Value = "Имя";
                worksheet.Cell(1, 4).Value = "Количество";

                int row = 2;
                foreach (PEZ? pez in MainWindowViewModel.BaseListPEZs)
                {
                    worksheet.Cell(row, 1).Value = pez.IdNumber;
                    worksheet.Cell(row, 2).Value = pez.Mark;
                    worksheet.Cell(row, 3).Value = pez.Name;
                    worksheet.Cell(row, 4).Value = pez.Quantity;
                    row++;
                }

                workbook.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                ShowError("SaveToExcel: Ошибка!", ex.ToString());
            }
        }

        public void MatchPEZMaterials()
        {
            try
            {
                if (FilePath == "")
                {
                    ShowError("Ошибка!", "Сначала необходимо выбрать файл!");
                    return;
                }

                string[]? materialNames = MaterialsList.Select(x => x.Name).ToArray(); // список имен в массив, потому что вроде как ExtractOne лучше/в принципе работает с массивами
                FuzzySharp.SimilarityRatio.Scorer.IRatioScorer? scorer = ScorerCache.Get<WeightedRatioScorer>();
                               
                Parallel.ForEach(CollectionPEZs, _element =>
                {
                    MatchColors(_element, materialNames, scorer);
                }); 
                   
                MainWindowViewModel.Instance.Us = new DetailsView();               
            }
            catch (Exception ex)
            {
                ShowError("Ошибка!", ex.ToString());
            }
        }

        void MatchColors(PEZ _element, string[]? materialNames, FuzzySharp.SimilarityRatio.Scorer.IRatioScorer? scorer)
        {
            FuzzySharp.Extractor.ExtractedResult<string>? match = FuzzySharp.Process.ExtractOne(_element.Name, materialNames, s => s, scorer);

            _element.Matched = materialNames[match.Index];

            _element.Color = match.Score switch
            {
                100 => "#66C190",
                > 70 => "#FFE666",
                _ => "#FF9166"
            };

            _element.Color = match.Score switch
            {
                100 => "#66C190",
                > 70 => "#FFE666",
                _ => "#FF9166"
            };

            if (_element.Name == "метка" || _element.Name == "заземление")
            {
                _element.Color = "#FFFFFF";
            }
        }

        #endregion        

        #region Методы работы с Материалами

        public void AddMaterial(Materials material)
        {
            string filePath = Path.Combine(AppContext.BaseDirectory.Substring(0, AppContext.BaseDirectory.IndexOf("bin") - 1), "Materials", "materials.xlsx");
            XLWorkbook wb = new XLWorkbook(filePath);
            var ws = wb.Worksheet(material.GroupNavigation.Name);

            // Найти последнюю строку нужной категории
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            int insertAtRow = -1;
            bool inCategory = false;

            for (int i = 1; i <= lastRow; i++)
            {
                var cell = ws.Cell(i, 2);
                var fontBold = cell.Style.Font.Bold;

                if (fontBold && cell.GetString() == material.CategoryNavigation.Name)
                {
                    inCategory = true;
                    continue;
                }

                if (inCategory)
                {
                    var nextCell = ws.Cell(i, 2);
                    if (nextCell.Style.Font.Bold)
                    {
                        // Встретили следующую категорию – вставлять перед ней
                        insertAtRow = i;
                        break;
                    }

                    if (i == lastRow)
                    {
                        insertAtRow = lastRow + 1;
                        break;
                    }
                }
            }

            if (insertAtRow == -1)
            {
                // Если категорию не нашли – просто добавляем в конец
                insertAtRow = lastRow + 1;
            }

            ws.Row(insertAtRow).InsertRowsAbove(1);
            var row = ws.Row(insertAtRow);

            row.Cell(1).Value = MaterialsList.Where(x => x.GroupNavigation.Name == material.GroupNavigation.Name).Max(x => x.IdNumber) + 1;
            row.Cell(2).Value = material.Name;
            row.Cell(3).Value = material.Measurement;
            row.Cell(4).Value = material.Analogs;
            row.Cell(5).Value = material.Note;

            wb.SaveAs(filePath);
        }

        public void ShowAddMaterials(Materials? material)
        {
            try
            {
                int id = 0;

                if (_addWindow != null) return;
                var viewModel = new AddEditVM(SelectedCategory, SelectedGroup);
                if (material != null) 
                    viewModel = new AddEditVM(material);

                _addWindow = new Window
                {
                    MinHeight = 600,
                    MinWidth = 1500,
                    Content = new AddEditView { DataContext = viewModel },
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    CanResize = false,
                    Title = "BK_Details_App",
                    Icon = new WindowIcon(AssetLoader.Open(new Uri("avares://BK_Details_App/Assets/logobk.ico")))
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

        public void DeleteMaterial(Materials material)
        {
            string filePath = Path.Combine(AppContext.BaseDirectory.Substring(0, AppContext.BaseDirectory.IndexOf("bin") - 1), "Materials", "materials.xlsx");
            XLWorkbook wb = new XLWorkbook(filePath);
            var ws = wb.Worksheet(material.GroupNavigation.Name);

            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            for (int i = 1; i <= lastRow; i++)
            {
                var cell = ws.Cell(i, 2);
                if (cell.GetString() == material.Name)
                {
                    ws.Row(i).Delete();
                    wb.SaveAs(filePath);
                    break;
                }
            }

            string fp = Path.Combine(Directory.GetCurrentDirectory(), "Materials", "test.xlsx");

            
            XLWorkbook workbook;
            if (File.Exists(fp))
            {
                workbook = new XLWorkbook(fp);
            }
            else
            {
                workbook = new XLWorkbook();
            }

            string sheetName = "Избранное";
            var sheetFavs = workbook.Worksheets.Contains(sheetName)
            ? workbook.Worksheet(sheetName)
            : workbook.AddWorksheet(sheetName);

            //Определяем первую пустую строку
            int lastRowFavs = sheetFavs.LastRowUsed()?.RowNumber() + 1 ?? 1;

            //Сохраняем файл
            workbook.SaveAs(fp);

            var sheet = wb.Worksheet(material.GroupNavigation.Name);

            if (sheet == null)
                throw new ArgumentException("Лист не найден");

            int rowCount = sheet.LastRowUsed()?.RowNumber() ?? 0;

            for (int i = 1; i <= rowCount; i++)
            {
                string cellValue = sheet.Cell(i, 1).GetString();
                if (!string.IsNullOrEmpty(cellValue) && cellValue == material.Name) sheet.Row(i).Delete();
            }
            workbook.SaveAs(fp);

            if (Favs.Count > 0)
            {
                new FavouritesVM().RemoveFromFavorite(material.Name);
            }
            MainWindowViewModel.Instance.Us = new DetailsView();
            ShowSuccess("Успех!", "Материал удален");
        }

        #endregion               
    }
}
