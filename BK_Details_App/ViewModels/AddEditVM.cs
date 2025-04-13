using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reactive;
using System.Text;
using System.Threading.Tasks;
using BK_Details_App.Models;
using ClosedXML.Excel;
using DynamicData;
using ReactiveUI;

namespace BK_Details_App.ViewModels
{
    internal class AddEditVM : ViewModelBase
    {
        #region Properties

        private string _header = "";
        public string Header { get => _header; set => this.RaiseAndSetIfChanged(ref _header, value); }

        private string _oldName;
        public string OldName { get => _oldName; set => this.RaiseAndSetIfChanged(ref _oldName, value); }


        private string _button = "";
        public string Button { get => _button; set => this.RaiseAndSetIfChanged(ref _button, value); }

        private Materials _newMaterial;
        internal Materials NewMaterial { get => _newMaterial; set => this.RaiseAndSetIfChanged(ref _newMaterial, value); }

        public ReactiveCommand<Unit, Unit> ToBackCommand { get; }
        public Action? CloseAction { get; set; }

        DetailsVM DetailsVMObj => new DetailsVM();

        #endregion

        #region Конструкторы

        public AddEditVM()
        {
        }

        public AddEditVM(Category category, Groups group)
        {
            try
            {
                _header = "Добавление материала";
                _button = "Добавить материал";
                _oldName = "";

                _newMaterial = new Materials();
                _newMaterial.CategoryNavigation = category;
                _newMaterial.GroupNavigation = group;

                ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
            }
            catch (Exception ex)
            {
                DetailsVMObj.ShowError("AddEditVM1: Ошибка!", ex.ToString());
            }
        }

        public AddEditVM(Materials material)
        {
            try
            {
                _header = "Редактирование материала";
                _button = "Редактировать материал";
                _oldName = material.Name;

                _newMaterial = material;

                ToBackCommand = ReactiveCommand.Create(() => CloseAction?.Invoke());
            }
            catch (Exception ex)
            {
                DetailsVMObj.ShowError("AddEditVM2: Ошибка!", ex.ToString());
            }
        }

        #endregion

        public void AddEdit()
        {
            try
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
                        if (DetailsVMObj.MaterialsList.Any(x => x.Name == NewMaterial.Name))
                        {
                            DetailsVMObj.ShowError("Внимание!", NewMaterial.Name + " уже существует!");
                            return;
                        }

                        MainWindowViewModel.AllMaterials.Add(NewMaterial);
                        DetailsVMObj.AddMaterial(NewMaterial);
                        CloseAction?.Invoke();
                        MainWindowViewModel.Instance.Us = new DetailsView();
                        DetailsVMObj.ShowSuccess("Успех", "Материал добавлен");
                    }
                    else
                    {
                        if (NewMaterial != null)
                        {
                            if (DetailsVMObj.MaterialsList.Any(x => x.Name == NewMaterial.Name) && NewMaterial.Name != OldName)
                            {
                                DetailsVMObj.ShowError("Внимание!", NewMaterial.Name + " уже существует!");
                                return;
                            }

                            List<string> favs = DetailsVMObj.ReadFavorites(DetailsVMObj.path);
                            if (favs.Contains(OldName))
                            {
                                string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Materials", "test.xlsx");
                                XLWorkbook workbook = new XLWorkbook(filePath);
                                var sheet = workbook.Worksheet("Избранное");

                                if (sheet == null)
                                    throw new ArgumentException("Лист не найден");

                                int rowCount = sheet.LastRowUsed()?.RowNumber() ?? 0;

                                for (int i = 1; i <= rowCount; i++)
                                {
                                    if (!string.IsNullOrEmpty(sheet.Cell(i, 1).GetString()) && sheet.Cell(i, 1).GetString() == OldName)
                                    {
                                        sheet.Cell(i, 1).Value = NewMaterial.Name;
                                        break;
                                    }
                                }
                                workbook.SaveAs(filePath);
                            }

                            //:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

                            string fp = Path.Combine(AppContext.BaseDirectory.Substring(0, AppContext.BaseDirectory.IndexOf("bin") - 1), "Materials", "materials.xlsx");
                            XLWorkbook wb = new XLWorkbook(fp);

                            foreach (var ws in wb.Worksheets)
                            {
                                foreach (var cell in ws.CellsUsed(c => c.HasFormula))
                                {
                                    if (string.IsNullOrWhiteSpace(cell.FormulaA1))
                                        cell.Clear(); // Или cell.Value = null;
                                }
                            }

                            var currentWorksheet = wb.Worksheet(NewMaterial.GroupNavigation.Name);

                            int lastRow = currentWorksheet.LastRowUsed()?.RowNumber() ?? 0;
                            for (int i = 1; i < lastRow; i++)
                            {
                                if (currentWorksheet.Cell(i, 2).Value.ToString() == OldName)
                                {
                                    currentWorksheet.Cell(i, 2).Value = NewMaterial.Name;
                                    currentWorksheet.Cell(i, 3).Value = NewMaterial.Measurement;
                                    currentWorksheet.Cell(i, 4).Value = NewMaterial.Analogs;
                                    currentWorksheet.Cell(i, 5).Value = NewMaterial.Note;
                                    break;
                                }
                            }

                            wb.SaveAs(fp);
                            CloseAction?.Invoke();
                            MainWindowViewModel.Instance.Us = new DetailsView();
                            DetailsVMObj.ShowSuccess("Успех", "Материал изменен");
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DetailsVMObj.ShowError("AddEdit: Ошибка!", ex.ToString());
            }
        }
    }
}
