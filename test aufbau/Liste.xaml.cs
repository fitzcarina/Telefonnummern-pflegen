using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace test_aufbau
{
    /// <summary>
    /// Interaktionslogik für Liste.xaml
    /// </summary>
    public partial class Liste : Window
    {
        public Liste()
        {
            InitializeComponent();
        }
        //Ruft eine Methode auf, die für die ausgabe der Excel + PDF für Firmenhandy ist
        private void handynummer_p(object sender, RoutedEventArgs e)
        {
            Excel_aufrufe.Firmenhandy();
            this.Close();
        }
        //Ruft eine Methode auf, die für die ausgabe der Excel + PDF für TelefonSchmal ist
        private void einspaltig_p(object sender, RoutedEventArgs e)
        {
            Excel_aufrufe.TelefonSchmal();
            this.Close();
        }
        //Ruft eine Methode auf, die für die ausgabe der Excel + PDF für TelefonZweiSpalten ist
        private void zweispaltig_p(object sender, RoutedEventArgs e)
        {
            Excel_aufrufe.TelefonZweiSpalten();
            this.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Excel_aufrufe.Firmenhandy();
            Excel_aufrufe.TelefonSchmal();
            Excel_aufrufe.TelefonZweiSpalten();
            this.Close();
        }
    }
}
