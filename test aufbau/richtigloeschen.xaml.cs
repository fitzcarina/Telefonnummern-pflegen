using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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
    public partial class richtigloeschen : Window
    {
        public string geben;
        public richtigloeschen(string geben)
        {
            InitializeComponent();
            //holt die Daten die im loeschen.xaml.cs übergeben wurden an richtigloeschen.xaml.cs
            this.geben = geben;
            ueberschrift.Content = "Sind Sie sicher das Sie den Benutzer : "+geben;
            ueberschrift2.Content = "unwiederruflich löschen möchten?";
        }
        private void Ja(object sender, RoutedEventArgs e)
        {
            string Nachname = "";
            string ID = "";
            //Delete SQl statement wird aufgerufen
            Class1.sqlDelete(geben, Nachname, ID );
            MessageBox.Show("User wurde unwiedeerruflich gelöscht !");
         }
        private void Nein(object sender, RoutedEventArgs e)
        {
            //es wird nichts gelöscht
            this.Close();
        }
    }
}
    
