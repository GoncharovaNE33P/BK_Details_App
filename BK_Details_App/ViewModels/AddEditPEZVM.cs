using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive;
using System.Text;
using System.Threading.Tasks;
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

        public ReactiveCommand<Unit,Unit> ToBackCommand { get; }
        public Action? CloseAction { get; set; }

        #endregion

        public AddEditPEZVM()
        {
            _headPage = "Добавление ПЭЗ";
            _buttonName = "Добавить ПЭЗ";

            ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
        }

        public AddEditPEZVM(string filePath)
        {
            _headPage = "Добавление ПЭЗ";
            _buttonName = "Добавить ПЭЗ";

            ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
        }

        public AddEditPEZVM(int id, string filePath)
        {
            _headPage = "Редактирование ПЭЗ";
            _buttonName = "Сохранить изменения ПЭЗ";

            ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
        }
                
    }
}
