using System;
using System.Data.SqlClient;
using System.Windows;

namespace test_aufbau
{ 
    public partial class Bearbeitenneu : Window
    {
        public Bearbeitenneu()
        {
            //beim aufruf von dem Button Bearbeiten wird beim laden eine SQL verbindung aufgebaut, damit Datensätze in dem Drop Down sind
            InitializeComponent();
            Ueberschrift.Content = "Bitte wählen Sie den Mitarbeiter, den Sie bearbeiten möchten";
            using (SqlConnection conn = new SqlConnection(@"server=vmsql01\prod;database=schnupp; trusted_connection=yes"))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("Select Nachname,Vorname,ID from tbl_Telefonnummern where Nachname IS NOT NULL AND Nachname != ' ' Order by Nachname", conn);
                SqlDataReader reader = cmd.ExecuteReader();
                //reader liest solange, bis er alle Elemente durch hat und schreibt sie dann in das Drop Down Menü
                while (reader.Read())
                {
                    Mitarbeiter.Items.Add(reader[0].ToString() + " " + reader[1].ToString() + " " + reader[2].ToString());
                }
                reader.Close();
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Label und Textboxen werden befüllt, mit aussagen
                Ueberschrift.Content = "Bitte geben Sie die neuen Daten ein und drücken Sie auf speichern";
                //der Textbox wird das Gewählte element übergeben
                Nachname.Text = Mitarbeiter.SelectedItem.ToString();
                //Datenbankverbindung wird aufgebaut, damit ein Mitarbeiter, von dem die Telefonnummer bearbeitet werden soll gewählt werden kann
                using (SqlConnection conn = new SqlConnection(@"server=vmsql01\prod;database=schnupp; trusted_connection=yes"))
                {
                    conn.Open();
                    string[] authorlist = Nachname.Text.ToString().Split(" ");
                    Nachname.Text = authorlist[0];
                    Vornames.Text = authorlist[1];
                    string IDs = authorlist[2];
                    SqlCommand cmd = new SqlCommand("Select DW,kurzw,Handy,ID from tbl_Telefonnummern  where ID=" + IDs + "", conn);
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        Durchwahl.Text = reader[0].ToString();
                        Kurzwahl.Text = reader[1].ToString();
                        Handy.Text = reader[2].ToString();
                        ID.Text = reader[3].ToString();
                    }
                    reader.Close();
                }
            }
            catch
            {
                MessageBox.Show("Der gewünschte Mitarbeiter konnte nicht ausgewählt werden");
            }
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string vorname = Vornames.Text;
            int s;
            //es wird getestet, ob ein String ist, damit keine Zahlen eingetragen werden können
            bool vorname_p = Int32.TryParse(Vornames.Text, out s);
            string nachname = Nachname.Text;
            bool nachname_p = Int32.TryParse(Nachname.Text, out s);
            //es wird getestet, ob ein String ist, damit sich nur Zahlen in der Durchwahl + Kurzwahl befinden
            bool durchwahl_p = Int32.TryParse(Durchwahl.Text, out s);
            string handy = Handy.Text;
            int n;
            bool kurzwahl_p = Int32.TryParse(Kurzwahl.Text, out n);
            string durchwahl_string = Durchwahl.Text;
            string kurzwahl_string = Kurzwahl.Text;
            string id = ID.Text;
            if (Class1.IsInGroup()== true)
            {
                //parameter werden übergeben an die Klasse, in dem weitere Prüfungen abgearbeitet werden
                if (Class1.pruefen(id, vorname, vorname_p, nachname_p, nachname, handy, kurzwahl_p, n, durchwahl_string, kurzwahl_string, durchwahl_p) == true)
                {
                    //wenn die Prüfung erfolgreich war wird der SQL update durchgeführt
                    Class1.sqlUpdate(id, vorname, vorname_p, nachname_p, nachname, handy, kurzwahl_p, n, durchwahl_string, kurzwahl_string, durchwahl_p);
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Bei dem Update ist ein Fehler aufgetreten");
                }
            }
            else
            {
                MessageBox.Show("Ihnen fehlt die Role_IT Berechtigung");
                this.Close();
            }

        }
    }
 }

