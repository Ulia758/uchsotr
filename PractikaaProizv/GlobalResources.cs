using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PractikaaProizv
{
    public static class GlobalResources
    {
        // Наблюдательная коллекция записей об увольнении
        public static ObservableCollection<Uvolnenie> UvolnenieItems { get; set; } = new ObservableCollection<Uvolnenie>();
    }
}
