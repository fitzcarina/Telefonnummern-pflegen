using System;
using System.Collections.Generic;
using System.Windows;




namespace test_aufbau
{
    public partial class Hinzufügen : Window
    {
        public Hinzufügen()
        {
            InitializeComponent();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Es wird getestet, ob der User der das Programm aufruft in der Role_IT Gruppenrichtline im AD ist
            if (Class1.IsInGroup() == true)
            {
                int s;
                int n;
                string id = "";
                    if (Class1.pruefen(id,  Vorname.Text, Int32.TryParse(Vorname.Text, out s), Int32.TryParse(Nachname.Text, out s), Nachname.Text, Handy.Text, Int32.TryParse(Kurzwahl.Text, out n), n, Durchwahl.Text, Kurzwahl.Text, Int32.TryParse(Durchwahl.Text, out s)) == true)
                    {
                        Class1.sqlInsert(id,  Vorname.Text, Int32.TryParse(Vorname.Text, out s), Int32.TryParse(Nachname.Text, out s), Nachname.Text, Handy.Text, Int32.TryParse(Kurzwahl.Text, out n), n, Durchwahl.Text, Kurzwahl.Text, Int32.TryParse(Durchwahl.Text, out s));
                    this.Close();
                }
            }
            else
            {
                    MessageBox.Show("Ihnen Fehlt die Role_IT Berechtigung");
            }
        }
    }
}



