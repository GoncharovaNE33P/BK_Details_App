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


        private string _quantityPEZ = "";
        public string QuantityPEZ { get => _quantityPEZ; set => this.RaiseAndSetIfChanged(ref _quantityPEZ, value); }


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
            
            LoadData(id, filePath);

            ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
        }

        private void LoadData(int id, string filePath)
        {
            List<PEZ> ListPEZ = MainWindowViewModel.BaseListPEZs;

            _newPEZ = ListPEZ.FirstOrDefault(x => x.IdNumber == id) ?? new PEZ();

            QuantityPEZ = NewPEZ.Quantity.ToString();
        }

        public void AddEditPEZ(PEZ myData)
        {      
            if (string.IsNullOrEmpty(NewPEZ.Name) || string.IsNullOrEmpty(NewPEZ.Mark) || string.IsNullOrEmpty(QuantityPEZ))
            {
                DetailsVMObject.ShowError("Ошибка!", "Заполните все поля!");
                return;
            }
            else
            {                

                if (int.TryParse(QuantityPEZ, out int result)) NewPEZ.Quantity = result;
                else
                {
                    DetailsVMObject.ShowError("Ошибка!", "Введено некорректное значение в поле \"Количество\"!");
                    return;
                }

                if (NewPEZ.IdNumber == 0)
                {
                    NewPEZ.IdNumber = MainWindowViewModel.BaseListPEZs.Max().IdNumber + 1;
                    MainWindowViewModel.BaseListPEZs.Add(NewPEZ);
                }
                else
                {
                    if (NewPEZ != null)
                    {
                        NewPEZ.Name = myData.Name;
                        NewPEZ.Mark = myData.Mark;
                        NewPEZ.Quantity = myData.Quantity;
                    }
                }
            }            
        }

        public async void SaveInFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;
        }        
    }
}
