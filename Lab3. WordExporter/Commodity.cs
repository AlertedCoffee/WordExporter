using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Lab3.WordExporter
{
    public class Commodity : INotifyPropertyChanged
    {
        private int id;
        public int Id
        {
            get { return id; }
            set
            {
                id = value;
                OnPropertyChanged(nameof(Id));
            }
        }

        private string product;
        public string Product
        {
            get { return product; }
            set
            {
                product = value;
                OnPropertyChanged(nameof(Product));
            }
        }

        private int count;
        public int Count
        {
            get { return count; } 
            set
            {
                count = value;
                OnPropertyChanged("Count");
                OnPropertyChanged("Amount");
            }
        }

        private double price;
        public double Price
        {
            get
            {
                return price;
            }
            set
            {
                price = value;
                OnPropertyChanged("Price");
                OnPropertyChanged("Amount");
            }
        }

        public double Amount
        {
            get { return count * price; }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string atr)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(atr));

            }
        }
    }
}
