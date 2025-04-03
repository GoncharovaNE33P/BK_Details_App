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

namespace BK_Details_App.ViewModels
{
    internal class AddEditPEZVM : ViewModelBase
    {
        #region properties

        private string _headPage = "";

        public string HeadPage { get => _headPage; set => this.RaiseAndSetIfChanged(ref _headPage, value); }

        private string _buttonName = "";

        public string ButtonName { get => _buttonName; set => this.RaiseAndSetIfChanged(ref _buttonName, value); }

        private PEZ _newPEZ;
        internal PEZ NewPEZ { get => _newPEZ; set => this.RaiseAndSetIfChanged(ref _newPEZ, value); }

        public ReactiveCommand<Unit,Unit> ToBackCommand { get; }
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
            _headPage = "Редактирование ПЭЗ";
            _buttonName = "Сохранить изменения ПЭЗ";
            
            LoadData(id, filePath);

            ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
        }

        private void LoadData(int id, string filePath)
        {
            List<PEZ> ListPEZ = new List<PEZ>();

            if (filePath.EndsWith(".xlsx") || filePath.EndsWith(".xls")) ListPEZ = LoadExcel(filePath);
            else if (filePath.EndsWith(".csv")) ListPEZ = LoadCsv(filePath);

            _newPEZ = ListPEZ.FirstOrDefault(x => x.IdNumber == id) ?? new PEZ();
        }

        private List<PEZ> LoadExcel(string filePath)
        {
            List<PEZ> ListPEZ = new List<PEZ>();

            try
            {               
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using FileStream? stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                using IExcelDataReader? reader = ExcelReaderFactory.CreateReader(stream);

                var result = reader.AsDataSet();
                var table = result.Tables[0];

                foreach (DataRow row in table.Rows.Cast<DataRow>().Skip(1))
                {
                    ListPEZ.Add(
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
            catch (Exception ex)
            {
                DetailsVMObject.ShowError("Ошибка!", ex.ToString());
            }

            return ListPEZ;
        }

        private List<PEZ> LoadCsv(string filePath)
        {
            List<PEZ> ListPEZ = new List<PEZ>();

            try
            {
                Encoding? encoding = Encoding.UTF8;

                byte[]? bytes = File.ReadAllBytes(filePath);
                if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) encoding = Encoding.UTF8;
                else encoding = Encoding.GetEncoding("Windows-1251");

                string[]? lines = File.ReadAllLines(filePath, encoding);

                foreach (string? line in lines.Skip(1))
                {
                    string[]? parts = line.Split(';');
                    if (parts.Length < 4) continue;
                    else
                    {
                        ListPEZ.Add(
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
            }
            catch (Exception ex)
            {
                DetailsVMObject.ShowError("Ошибка!", ex.ToString());
            }

            return ListPEZ;
        }
    }
}
