using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive;
using System.Text;
using System.Threading.Tasks;
using BK_Details_App.Models;
using DynamicData;
using ReactiveUI;

namespace BK_Details_App.ViewModels
{
    internal class AddEditVM : ViewModelBase
    {
        #region properties

        private string _header = "";
        public string Header { get => _header; set => this.RaiseAndSetIfChanged(ref _header, value); }


        private string _button = "";
        public string Button { get => _button; set => this.RaiseAndSetIfChanged(ref _button, value); }

        private Materials _newMaterial;
        internal Materials NewMaterial { get => _newMaterial; set => this.RaiseAndSetIfChanged(ref _newMaterial, value); }

        public ReactiveCommand<Unit, Unit> ToBackCommand { get; }
        public ReactiveCommand<Unit, Unit> Command { get; }
        public Action? CloseAction { get; set; }

        DetailsVM DetailsVMObj => new DetailsVM();

        #endregion

        public AddEditVM()
        {
            _header = "Добавление материала";
            _button = "Добавить материал";

            _newMaterial = new Materials();

            ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
        }

        public void AddEdit()
        {
            if (string.IsNullOrEmpty(NewMaterial.Name))
            {
                DetailsVMObj.ShowError("Ошибка!", "Поле Имя обязательно для заполнения");
                return;
            }
            else
            {
                if (NewMaterial.IdNumber == 0)
                {
                    MainWindowViewModel.AllMaterials.Add(NewMaterial);
                    NewMaterial.CategoryNavigation = DetailsVMObj.SelectedCategory;
                    NewMaterial.GroupNavigation = DetailsVMObj.SelectedGroup;
                    DetailsVMObj.AddMaterial(NewMaterial);
                    CloseAction?.Invoke();
                    MainWindowViewModel.Instance.Us = new DetailsView();
                    DetailsVMObj.ShowSuccess("Успех", "Материал добавлен");
                }
                else
                {
                    if (NewMaterial != null)
                    {
                        return;//::::::::::::::::::::::::::::
                    }
                }
            }
        }
    }
}
