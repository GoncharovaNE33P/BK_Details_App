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

namespace BK_Details_App.ViewModels
{
    internal class AddEditPEZVM : ViewModelBase
    {
        #region properties

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
        public ReactiveCommand<Unit,Unit> Command { get; }
        public Action? CloseAction { get; set; }

        DetailsVM DetailsVMObject => new DetailsVM();

        #endregion

        public AddEditPEZVM()
        {
            _headPage = "Добавление ПЭЗ";
            _buttonName = "Добавить ПЭЗ";

            _newPEZ = new PEZ();            

            ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
            
        }

        public AddEditPEZVM(int id, string filePath)
        {
            if (id == 0)
            {
                _headPage = "Добавление ПЭЗ";
                _buttonName = "Добавить ПЭЗ";
            }
            else
            {
                _headPage = "Редактирование ПЭЗ";
                _buttonName = "Сохранить изменения ПЭЗ";                
            }

            FilePath = filePath;
            
            LoadData(id);

            ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
        }

        private void LoadData(int id)
        {
            List<PEZ> ListPEZ = MainWindowViewModel.BaseListPEZs;

            NewPEZ = ListPEZ.FirstOrDefault(x => x.IdNumber == id) ?? new PEZ();

            QuantityPEZ = NewPEZ.Quantity.ToString();
        }

        public void AddEditPEZ()
        {
            try
            {
                if (string.IsNullOrEmpty(NewPEZ.Name) || string.IsNullOrEmpty(NewPEZ.Mark) || string.IsNullOrEmpty(QuantityPEZ))
                {
                    DetailsVMObject.ShowError("Ошибка!", "Заполните все поля!");
                    return;
                }

                if (!int.TryParse(QuantityPEZ, out int result) || result == 0)
                {
                    DetailsVMObject.ShowError("Ошибка!", "Введено некорректное значение в поле \"Количество\"!");
                    return;
                }

                if (NewPEZ != null) NewPEZ.Quantity = result;

                if (FilePath.EndsWith(".csv"))
                    ProcessCsv(FilePath);
                else if (FilePath.EndsWith(".xlsx") || FilePath.EndsWith(".xls"))
                    ProcessExcel(FilePath);
            }
            catch (Exception ex)
            {
                DetailsVMObject.ShowError("Ошибка!", ex.ToString());
            }
        }

        public async void ProcessCsv(string filePath)
        {
            try
            {
                using(FileStream stream = new FileStream(filePath, FileMode.Append, FileAccess.Write))
                using (StreamWriter writer = new StreamWriter(stream, Encoding.GetEncoding("Windows-1251")))
                using (CsvWriter csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Delimiter = ";",
                    Encoding = Encoding.GetEncoding("Windows-1251"),
                    HasHeaderRecord= !File.Exists(filePath) || new FileInfo(filePath).Length == 0
                }))
                {
                    if (NewPEZ.IdNumber == 0)
                    {
                        NewPEZ.IdNumber = MainWindowViewModel.BaseListPEZs.Count() > 0 ? MainWindowViewModel.BaseListPEZs.Max(p => p.IdNumber) + 1 : 1;
                        MainWindowViewModel.BaseListPEZs.Add(NewPEZ);

                        csv.WriteRecord(NewPEZ);

                        DetailsVMObject.CollectionPEZs.Clear();
                        DetailsVMObject.CollectionPEZs.AddRange(MainWindowViewModel.BaseListPEZs);
                        DetailsVMObject.CountItemsFilePEZ = MainWindowViewModel.BaseListPEZs.Count();

                        MainWindowViewModel.Instance.Us = new DetailsView();

                        DetailsVMObject.ShowSuccess("Успех!", $"{NewPEZ.Name} добавлен в файл {DetailsVMObject.NameFile}");
                    }
                    else
                    {
                        List<PEZ> PezList = MainWindowViewModel.BaseListPEZs;
                        PezList.RemoveAll(p => p.IdNumber == NewPEZ.IdNumber);
                        PezList.Add(NewPEZ);
                        writer.BaseStream.SetLength(0);
                        csv.WriteRecords(PezList);

                        MainWindowViewModel.BaseListPEZs = PezList;
                        DetailsVMObject.CollectionPEZs.Clear();
                        DetailsVMObject.CollectionPEZs.AddRange(PezList);
                        DetailsVMObject.CountItemsFilePEZ = MainWindowViewModel.BaseListPEZs.Count();

                        DetailsVMObject.ShowSuccess("Успех!", $"{NewPEZ.Name} изменён в файле {DetailsVMObject.NameFile}");
                    }
                }                    
            }
            catch (Exception ex)
            {
                DetailsVMObject.ShowError("Ошибка!", ex.ToString());
            }
        }

        public async void ProcessExcel(string filePath)
        {

        }
    }
}
