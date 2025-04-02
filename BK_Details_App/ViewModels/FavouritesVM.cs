using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BK_Details_App.Models;
using ReactiveUI;

namespace BK_Details_App.ViewModels
{
    internal class FavouritesVM : ViewModelBase
    {
        List<Materials> _materialsList = new();
        public List<Materials> MaterialsList
        {
            get => _materialsList;
            set => this.RaiseAndSetIfChanged(ref _materialsList, value);
        }

        public FavouritesVM()
        {
            DetailsVM detailsVM = new DetailsVM();
            List<string> buf = [.. detailsVM.ReadFavorites()];
            MaterialsList = detailsVM.MaterialsList.Where(x => buf.Contains(x.Name)).ToList();
        }
    }
}
