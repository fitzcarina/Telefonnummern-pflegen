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
    public partial class loeschen : Window
    {
    public string geben { get; set; }   
        public loeschen()
        {
            //beim aufruf des buttons werden im Drop Down alle Mitarbeiter schon angezeigt
            InitializeComponent();
            using (SqlConnection conn = new SqlConnection(@"server=vmsql01\prod;database=schnupp; trusted_connection=yes"))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("Select Nachname,Vorname, ID from tbl_Telefonnummern", conn);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    Mitarbeiter.Items.Add(reader[0].ToString() + " " + reader[1].ToString() + " " + reader[2].ToString());
                }
                reader.Close();
            } 
        }
       
        private void löschen(object sender, RoutedEventArgs e)
        {
            //übergabe an richtigloeschen.xaml.cs 
                if(Mitarbeiter.SelectedItem == null)
                {
                    MessageBox.Show("Es wurde kein Mitarbeiter gewählt, den Sie löschen möchten");
                }
                else
                {
                    geben = Mitarbeiter.SelectedItem.ToString();
                    Window richtigloeschen = new richtigloeschen(geben);
                richtigloeschen.Owner = this;
                    richtigloeschen.ShowDialog();
                }
    }
    }
}
