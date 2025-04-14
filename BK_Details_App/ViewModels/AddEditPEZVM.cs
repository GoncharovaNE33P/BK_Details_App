using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reactive;
using System.Text;
using System.Threading.Tasks;
using BK_Details_App.Models;
using DynamicData;
using ExcelDataReader;
using ReactiveUI;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using Aspose.Cells;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace BK_Details_App.ViewModels
{
    internal class AddEditPEZVM : ViewModelBase
    {
        #region Properties

        private string _headPage = "";
        public string HeadPage { get => _headPage; set => this.RaiseAndSetIfChanged(ref _headPage, value); }

        private string _buttonName = "";
        public string ButtonName { get => _buttonName; set => this.RaiseAndSetIfChanged(ref _buttonName, value); }

        private string _quantityPEZ = "";
        public string QuantityPEZ { get => _quantityPEZ; set => this.RaiseAndSetIfChanged(ref _quantityPEZ, value); }

        private string _filePath;
        public string FilePath { get => _filePath; set => this.RaiseAndSetIfChanged(ref _filePath, value); }

        private PEZ _newPEZ;
        internal PEZ NewPEZ { get => _newPEZ; set => this.RaiseAndSetIfChanged(ref _newPEZ, value); }

        public ReactiveCommand<Unit,Unit> ToBackCommand { get; }
        public Action? CloseAction { get; set; }

        DetailsVM DetailsVMObj => new DetailsVM();

        private string _oldName = "";
        public string OldName { get => _oldName; set => this.RaiseAndSetIfChanged(ref _oldName, value); }

        #endregion

        #region Конструкторы

        public AddEditPEZVM()
        {
            try
            {
                _headPage = "Добавление ПЭЗ";
                _buttonName = "Добавить ПЭЗ";

                _newPEZ = new PEZ();

                ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
            }
            catch (Exception ex)
            {
                DetailsVMObj.ShowError("AddEditPEZVM1: Ошибка!", ex.ToString());
            }
        }

        public AddEditPEZVM(int id, string filePath)
        {
            try
            {
                LoadData(id);

                if (id == 0)
                {
                    _headPage = "Добавление ПЭЗ";
                    _buttonName = "Добавить ПЭЗ";
                    _oldName = "";
                }
                else
                {
                    _headPage = "Редактирование ПЭЗ";
                    _buttonName = "Сохранить изменения ПЭЗ";
                    _oldName = _newPEZ.Name;
                }

                FilePath = filePath;

                ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
            }
            catch (Exception ex)
            {
                DetailsVMObj.ShowError("AddEditPEZVM2: Ошибка!", ex.ToString());
            }
        }

        #endregion

        private void LoadData(int id)
        {
            try
            {
                PEZ? original = MainWindowViewModel.BaseListPEZs.FirstOrDefault(x => x.IdNumber == id);

                if (original != null)
                {
                    NewPEZ = original.Clone();
                    QuantityPEZ = original.Quantity.ToString();
                }
                else
                {
                    NewPEZ = new PEZ();
                    QuantityPEZ = "";
                }
            }
            catch (Exception ex)
            {
                DetailsVMObj.ShowError("LoadData: Ошибка!", ex.ToString());
            }
        }


        public void AddEditPEZ()
        {
            try
            {
                if (string.IsNullOrEmpty(NewPEZ.Name) || string.IsNullOrEmpty(NewPEZ.Mark) || string.IsNullOrEmpty(QuantityPEZ))
                {
                    DetailsVMObj.ShowError("Ошибка!", "Заполните все поля!");
                    return;
                }

                if (!int.TryParse(QuantityPEZ, out int result) || result == 0)
                {
                    DetailsVMObj.ShowError("Ошибка!", "Введено некорректное значение в поле \"Количество\"!");
                    return;
                }

                if (NewPEZ != null) NewPEZ.Quantity = result;

                if (FilePath.EndsWith(".csv"))
                {
                    ProcessCsv(FilePath);
                    return;
                }

                else if (FilePath.EndsWith(".xlsx") || FilePath.EndsWith(".xls"))
                {
                    ProcessExcel(FilePath);
                }

                MainWindowViewModel.Instance.Us = new DetailsView();
                    
            }
            catch (Exception ex)
            {
                DetailsVMObj.ShowError("AddEditPEZ: Ошибка!", ex.ToString());
            }
        }

        #region Сохранить изменения в файле

        public void ProcessCsv(string filePath)
        {
            try
            {
                if (NewPEZ.IdNumber == 0)
                {
                    if (MainWindowViewModel.BaseListPEZs.Any(x => x.Name.Trim().ToLower() == NewPEZ.Name.Trim().ToLower()))
                    {
                        DetailsVMObj.ShowError("Внимание!", NewPEZ.Name + " уже существует!");
                        return;
                    }

                    NewPEZ.IdNumber = MainWindowViewModel.BaseListPEZs.Count > 0
                        ? MainWindowViewModel.BaseListPEZs.Max(p => p.IdNumber) + 1
                        : 1;

                    MainWindowViewModel.BaseListPEZs.Add(NewPEZ);

                    using (var stream = new FileStream(filePath, FileMode.Append, FileAccess.Write))
                    using (var writer = new StreamWriter(stream, Encoding.GetEncoding("Windows-1251")))
                    {
                        string line = $"{NewPEZ.IdNumber};{NewPEZ.Mark};{NewPEZ.Name};{NewPEZ.Quantity}";
                        writer.WriteLine(line);
                    }

                    DetailsVMObj.CollectionPEZs.Clear();
                    DetailsVMObj.CollectionPEZs.AddRange(MainWindowViewModel.BaseListPEZs);

                    CloseAction?.Invoke();

                    MainWindowViewModel.Instance.Us = new DetailsView();

                    DetailsVMObj.ShowSuccess("Успех!", $"{NewPEZ.Name} добавлен в файл {DetailsVMObj.NameFile}");
                }
                else
                {
                    if (MainWindowViewModel.BaseListPEZs.Any(x => x.Name.Trim().ToLower() == NewPEZ.Name.Trim().ToLower()) &&
                    NewPEZ.Name.Trim().ToLower() != OldName.Trim().ToLower())
                    {
                        DetailsVMObj.ShowError("Внимание!", NewPEZ.Name + " уже существует!");
                        return;
                    }

                    List<PEZ> pezList = MainWindowViewModel.BaseListPEZs;

                    int index = pezList.FindIndex(p => p.IdNumber == NewPEZ.IdNumber);
                    if (index != -1) pezList[index] = NewPEZ;

                    MainWindowViewModel.BaseListPEZs = pezList;

                    using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.GetEncoding("Windows-1251")))
                    {
                        writer.WriteLine("#;Метка;Имя;Количество");

                        foreach (PEZ? p in pezList)
                        {
                            string line = $"{p.IdNumber};{p.Mark};{p.Name};{p.Quantity}";
                            writer.WriteLine(line);
                        }
                    }

                    DetailsVMObj.CollectionPEZs.Clear();
                    DetailsVMObj.CollectionPEZs.AddRange(pezList);

                    CloseAction?.Invoke();

                    MainWindowViewModel.Instance.Us = new DetailsView();

                    DetailsVMObj.ShowSuccess("Успех!", $"{NewPEZ.Name} изменён в файле {DetailsVMObj.NameFile}");
                }
            }
            catch (Exception ex)
            {
                DetailsVMObj.ShowError("ProcessCsv: Ошибка!", ex.ToString());
            }
        }

        public void ProcessExcel(string filePath)
        {
            try
            {
                XLWorkbook workbook;
                IXLWorksheet worksheet;

                // Загружаем существующий файл или создаём новый
                if (File.Exists(filePath))
                {
                    workbook = new XLWorkbook(filePath);
                    worksheet = workbook.Worksheets.First();
                }
                else
                {
                    workbook = new XLWorkbook();
                    worksheet = workbook.AddWorksheet(DetailsVMObj.NameFile);

                    // Заголовки
                    worksheet.Cell(1, 1).Value = "#";
                    worksheet.Cell(1, 2).Value = "Метка";
                    worksheet.Cell(1, 3).Value = "Имя";
                    worksheet.Cell(1, 4).Value = "Количество";
                }

                if (NewPEZ.IdNumber == 0)
                {
                    if (MainWindowViewModel.BaseListPEZs.Any(x => x.Name.Trim().ToLower() == NewPEZ.Name.Trim().ToLower()))
                    {
                        DetailsVMObj.ShowError("Внимание!", NewPEZ.Name + " уже существует!");
                        return;
                    }

                    NewPEZ.IdNumber = MainWindowViewModel.BaseListPEZs.Count > 0
                        ? MainWindowViewModel.BaseListPEZs.Max(p => p.IdNumber) + 1
                        : 1;

                    MainWindowViewModel.BaseListPEZs.Add(NewPEZ);

                    int newRow = worksheet.LastRowUsed().RowNumber() + 1;

                    worksheet.Cell(newRow, 1).Value = NewPEZ.IdNumber;
                    worksheet.Cell(newRow, 2).Value = NewPEZ.Mark;
                    worksheet.Cell(newRow, 3).Value = NewPEZ.Name;
                    worksheet.Cell(newRow, 4).Value = NewPEZ.Quantity;

                    workbook.SaveAs(filePath);

                    DetailsVMObj.CollectionPEZs.Clear();
                    DetailsVMObj.CollectionPEZs.AddRange(MainWindowViewModel.BaseListPEZs);

                    CloseAction?.Invoke();

                    DetailsVMObj.ShowSuccess("Успех!", $"{NewPEZ.Name} добавлен в файл {DetailsVMObj.NameFile}");                    
                }
                else
                {
                    if (MainWindowViewModel.BaseListPEZs.Any(x => x.Name.Trim().ToLower() == NewPEZ.Name.Trim().ToLower()) &&
                    NewPEZ.Name.Trim().ToLower() != OldName.Trim().ToLower())
                    {
                        DetailsVMObj.ShowError("Внимание!", NewPEZ.Name + " уже существует!");
                        return;
                    }

                    List<PEZ> pezList = MainWindowViewModel.BaseListPEZs;
                    int index = pezList.FindIndex(p => p.IdNumber == NewPEZ.IdNumber);
                    if (index != -1) pezList[index] = NewPEZ;

                    MainWindowViewModel.BaseListPEZs = pezList;

                    // Полная перезапись
                    worksheet.Clear(); // полностью очищаем
                    worksheet.Cell(1, 1).Value = "#";
                    worksheet.Cell(1, 2).Value = "Метка";
                    worksheet.Cell(1, 3).Value = "Имя";
                    worksheet.Cell(1, 4).Value = "Количество";

                    for (int i = 0; i < pezList.Count; i++)
                    {
                        PEZ? p = pezList[i];
                        worksheet.Cell(i + 2, 1).Value = p.IdNumber;
                        worksheet.Cell(i + 2, 2).Value = p.Mark;
                        worksheet.Cell(i + 2, 3).Value = p.Name;
                        worksheet.Cell(i + 2, 4).Value = p.Quantity;
                    }

                    workbook.SaveAs(filePath);

                    DetailsVMObj.CollectionPEZs.Clear();
                    DetailsVMObj.CollectionPEZs.AddRange(pezList);

                    CloseAction?.Invoke();

                    DetailsVMObj.ShowSuccess("Успех!", $"{NewPEZ.Name} изменён в файле {DetailsVMObj.NameFile}");                    
                }
            }
            catch (Exception ex)
            {
                DetailsVMObj.ShowError("ProcessExcel: Ошибка!", ex.ToString());
            }
        }

        #endregion
    }
}
