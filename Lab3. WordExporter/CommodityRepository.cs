using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab3.WordExporter
{
    public class CommodityRepository
    {
        private ObservableCollection<Commodity> commodity;

        public double Sum
        {
            get
            {
                return commodity.Sum(x => x.Amount);
            }
        }

        public ObservableCollection<Commodity> Commodity
        {
            get
            {
                return commodity;            
            }
            set
            {
                commodity = value;
            }
        }

        public CommodityRepository()
        {
            commodity = new ObservableCollection<Commodity>();
        }


    }
}
