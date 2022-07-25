using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;



namespace test_aufbau
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void liste(object sender, RoutedEventArgs e)
        {

            //neues Fenster namens Liste wird erzeugt
         //  MainWindow = ownedWindow;
            Window Liste = new Liste();
            Liste.Owner = this;
            //Liste.WindowStartupLocation = Window.StartupLocation;
            Liste.ShowDialog();
        }
        private void hinzufügen(object sender, RoutedEventArgs e)
        {   //neues Fenster namens Hinzufügen wird erzeugt
            Window Hinzufügen = new Hinzufügen();
            Hinzufügen.Owner = this;
            Hinzufügen.ShowDialog();
        }
        private void bearbeiten(object sender, RoutedEventArgs e)
        {
            //neues Fenster namens Bearbeitenneu wird erzeugt
            Window Bearbeitenneu = new Bearbeitenneu();
            Bearbeitenneu.Owner = this;
            Bearbeitenneu.ShowDialog();
        }

        private void löschen(object sender, RoutedEventArgs e)
        {
            //neues Fenster namens loeschen wird erzeugt
            Window loeschen = new loeschen();
            loeschen.Owner = this;
            loeschen.ShowDialog();
        }
    }
}
