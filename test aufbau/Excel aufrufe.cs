using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows;

namespace test_aufbau
{
    internal class Excel_aufrufe : Helper_DB
    {
        public static void Firmenhandy()
        {
            //Heutiges datum für die Excel + PDF Table
            DateTime now = DateTime.Now;
            using (SqlConnection conn = new SqlConnection(db_connection()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("  Select Count(*) from tbl_Telefonnummern   where Handy is not null   ", conn);
                int anzahl = (Int32)cmd.ExecuteScalar();
                cmd = new SqlCommand("  Select Vorname, Nachname, DW, Kurzw, Handy from tbl_Telefonnummern where Handy is not null order by Nachname ", conn);
                SqlDataReader reader = cmd.ExecuteReader();
                string[] vorname = new string[anzahl];
                string[] nachname = new string[anzahl];
                string[] durchwahl = new string[anzahl];
                string[] kurzwahl = new string[anzahl];
                string[] handy = new string[anzahl];
                for (int i = 0; reader.Read(); i++)
                {
                    vorname[i] = reader[0].ToString();
                    nachname[i] = reader[1].ToString();
                    durchwahl[i] = reader[2].ToString();
                    kurzwahl[i] = reader[3].ToString();
                    handy[i] = reader[4].ToString();
                }
                reader.Close();
                try
                {
                    
                    MessageBox.Show("Bitte Warten Sie Kurz Ihre Excel Liste wird generiert");
                    //Excel wird erstellt
                    Microsoft.Office.Interop.Excel.Application xlApp;
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;
                    Microsoft.Office.Interop.Excel.Range chartRange;
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    string[] NachVor = new string[anzahl];
                    for (int i = 0; i < anzahl; i++)
                    {
                        NachVor[i] = nachname[i] + " " + vorname[i];
                    }
                    string stand = "Stand: ";
                    xlWorkSheet.Cells[1, 1] = stand + DateTime.Now.ToString(" d/M/yyyy");
                    xlWorkSheet.Cells[1, 2] = "Rufnummer";
                    xlWorkSheet.Cells[1, 3] = "Kurzwahl";
                    xlWorkSheet.Cells[1, 4] = stand + DateTime.Now.ToString(" d/M/yyyy");
                    xlWorkSheet.Cells[1, 5] = "Rufnummern";
                    xlWorkSheet.Cells[1, 6] = "Kurzwahl";
                    for (int i = 0; i < vorname.Length; i++)
                    {
                        xlWorkSheet.Cells[i + 2, 1] = NachVor[i];
                        xlWorkSheet.Cells[i + 2, 3] = kurzwahl[i];
                        xlWorkSheet.Cells[i + 2, 2] = handy[i];
                        xlWorkSheet.Cells[i + 2, 4] = NachVor[i];
                        xlWorkSheet.Cells[i + 2, 6] = kurzwahl[i];
                        xlWorkSheet.Cells[i + 2, 5] = handy[i];
                    }
                    chartRange = xlWorkSheet.get_Range("a1", "f1");
                    chartRange = xlWorkSheet.get_Range("a1", "f1");
                    chartRange.Font.Bold = true;
                    int rand = anzahl + 1;
                    chartRange = xlWorkSheet.get_Range("a1", "f" + (anzahl + 1));
                    chartRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic);
                    foreach (Microsoft.Office.Interop.Excel.Range cell in chartRange.Rows[1].Cells)
                    {
                        cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        cell.Font.Bold = true;
                    }

                    chartRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xlApp.DisplayAlerts = false;
                    xlWorkSheet.Columns["A:F"].AutoFit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    xlWorkSheet.PageSetup.Zoom = false;
                    xlWorkBook.SaveAs(@"M:\Kollegen\Telefonlisten\Firmenhandys.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    // Excel Format wird auch als PDF gespeichert
                    xlWorkBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, @"M:\Kollegen\Telefonlisten\Firmenhandys.pdf");
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                    releaseObject(xlApp);
                    releaseObject(xlWorkBook);
                    releaseObject(xlWorkSheet); var p = new System.Diagnostics.Process();
                    p.StartInfo = new ProcessStartInfo(@"M:\Kollegen\Telefonlisten\Firmenhandys.xls")
                    {
                        UseShellExecute = true
                    };
                    p.Start();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Die Datei ist gerade noch geöffnet, bitte versuchen Sie es zu einem Späteren Zeitpunkt erneut");
                }
                void releaseObject(object obj)
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                        obj = null;
                    }
                    catch (Exception ex)
                    {
                        obj = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                    
                }










            }
        }
        public static void TelefonSchmal()
        {
            DateTime now = DateTime.Now;
            using (SqlConnection conn = new SqlConnection(db_connection()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("  Select Count(*) from tbl_Telefonnummern    ", conn);
                int anzahl = (Int32)cmd.ExecuteScalar();
                cmd = new SqlCommand("  Select Vorname, Nachname, DW, Kurzw, Handy from tbl_Telefonnummern  order by Nachname ", conn);
                SqlDataReader reader = cmd.ExecuteReader();
                string[] vorname = new string[anzahl];
                string[] nachname = new string[anzahl];
                string[] durchwahl = new string[anzahl];
                string[] kurzwahl = new string[anzahl];
                string[] handy = new string[anzahl];
                for (int i = 0; reader.Read(); i++)
                {
                    vorname[i] = reader[0].ToString();
                    nachname[i] = reader[1].ToString();
                    durchwahl[i] = reader[2].ToString();
                    kurzwahl[i] = reader[3].ToString();
                    handy[i] = reader[4].ToString();
                }
                reader.Close();
                MessageBox.Show("Bitte Warten Sie Kurz Ihre Excel Liste wird generiert");
                try
                {

                    Microsoft.Office.Interop.Excel.Application xlApp;
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

                    object misValue = System.Reflection.Missing.Value;
                    Microsoft.Office.Interop.Excel.Range chartRange;
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    string[] NachVor = new string[anzahl];
                    for (int i = 0; i < anzahl; i++)
                    {
                        NachVor[i] = nachname[i] + " " + vorname[i];
                    }

                    xlWorkSheet.Cells[1, 1] = "Stand: " + DateTime.Now.ToString("d/M/yyyy");
                    xlWorkSheet.Cells[1, 2] = "Nr.";
                    xlWorkSheet.Cells[1, 3] = "Kurzwa.";
                    xlWorkSheet.Cells[1, 4] = "Stand: " + DateTime.Now.ToString("d/M/yyyy");
                    xlWorkSheet.Cells[1, 6] = "Kurzwa.";
                    xlWorkSheet.Cells[1, 5] = "Nr.";
                    xlWorkSheet.Cells[1, 7] = "Stand: " + DateTime.Now.ToString("d/M/yyyy");
                    xlWorkSheet.Cells[1, 9] = "Kurzwa.";
                    xlWorkSheet.Cells[1, 8] = "Nr.";
                    for (int i = 0; i < vorname.Length; i++)
                    {
                        xlWorkSheet.Cells[i + 2, 1] = NachVor[i];
                        xlWorkSheet.Cells[i + 2, 3] = kurzwahl[i];
                        xlWorkSheet.Cells[i + 2, 2] = durchwahl[i];
                        xlWorkSheet.Cells[i + 2, 4] = NachVor[i];
                        xlWorkSheet.Cells[i + 2, 5] = durchwahl[i];
                        xlWorkSheet.Cells[i + 2, 6] = kurzwahl[i];
                        xlWorkSheet.Cells[i + 2, 7] = NachVor[i];
                        xlWorkSheet.Cells[i + 2, 8] = durchwahl[i];
                        xlWorkSheet.Cells[i + 2, 9] = kurzwahl[i];
                    }
                    chartRange = xlWorkSheet.get_Range("a1", "i1");
                    chartRange = xlWorkSheet.get_Range("a1", "i1");
                    chartRange.Font.Bold = true;
                    int rand = anzahl + 1;
                    chartRange = xlWorkSheet.get_Range("a1", "i" + (anzahl + 1));
                    chartRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic);
                    foreach (Microsoft.Office.Interop.Excel.Range cell in chartRange.Rows[1].Cells)
                    {
                        cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        cell.Font.Bold = true;
                    }
                    chartRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xlApp.DisplayAlerts = false;
                    //excel spalten zentieren
                    xlWorkSheet.Columns["A:I"].AutoFit();
                    xlWorkSheet.PageSetup.Zoom = false;
                   // xlWorkSheet.OnSheetactivate(xlWorkSheet);
                   // xlWorkSheet.PageSetup = 1;


                    //Worksheets("Sheet1").PageSetup
                    //xlWorkBook.
                    xlWorkBook.SaveAs(@"M:\Kollegen\Telefonlisten\TelefonSchmal.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, @"M:\Kollegen\Telefonlisten\TelefonSchmal.pdf");
                    xlWorkBook.Close(true, misValue, misValue);
                    //xlApp.Quit();
                    releaseObject(xlApp);
                    releaseObject(xlWorkBook);
                    releaseObject(xlWorkSheet); var p = new System.Diagnostics.Process();
                    p.StartInfo = new ProcessStartInfo(@"M:\Kollegen\Telefonlisten\TelefonSchmal.xls")
                    {
                        UseShellExecute = true
                    };
                    p.Start();
                }
                catch( Exception ex)
                {
                    MessageBox.Show("Die Excel Datei ist gerade noch geöffnet, bitte probieren sie es zu einem Spätern Zeitpunkt erneut");
                }
             
            }
            void releaseObject(object obj)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
                catch (Exception ex)
                {
                    obj = null;
                }
                finally
                {
                    GC.Collect();
                }
            }
        }

        public static void TelefonZweiSpalten()
        { 
                DateTime now = DateTime.Now;
                using (SqlConnection conn = new SqlConnection(db_connection()))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("  Select Count(*) from tbl_Telefonnummern    ", conn);
                    int anzahl = (Int32)cmd.ExecuteScalar();
                    cmd = new SqlCommand("  Select Vorname, Nachname, DW, Kurzw, Handy from tbl_Telefonnummern  order by Nachname ", conn);
                    SqlDataReader reader = cmd.ExecuteReader();
                    string[] vorname = new string[anzahl];
                    string[] nachname = new string[anzahl];
                    string[] durchwahl = new string[anzahl];
                    string[] kurzwahl = new string[anzahl];
                    string[] handy = new string[anzahl];
                    for (int i = 0; reader.Read(); i++)
                    {
                        vorname[i] = reader[0].ToString();
                        nachname[i] = reader[1].ToString();
                        durchwahl[i] = reader[2].ToString();
                        kurzwahl[i] = reader[3].ToString();
                        handy[i] = reader[4].ToString();
                    }
                reader.Close();
                MessageBox.Show("Bitte Warten Sie Kurz Ihre Excel Liste wird generiert");
                try
                {
                    int spalte1 = anzahl / 2;
                    int spalte2 = anzahl / 2;
                    int geteilt = anzahl % 2;
                    int j = 6;
                    if ((anzahl + 6) % 2 != 0)
                    {
                        spalte1 = (anzahl / 2) + 1;
                        spalte2 = (anzahl / 2) - 1;
                        j = 7;
                    }
                    else if (anzahl % 2 == 0)
                    {
                        spalte1 = (anzahl / 2);
                        spalte1 = (anzahl / 2);
                        j = 6;
                    }
                    {
                    }
                    Microsoft.Office.Interop.Excel.Application xlApp;
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;
                    Microsoft.Office.Interop.Excel.Range chartRange;
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    string[] NachVor = new string[anzahl];
                    for (int i = 0; i < anzahl; i++)
                    {
                        NachVor[i] = nachname[i] + " " + vorname[i];
                    }
                    xlWorkSheet.Cells[1, 1] = "Dw";
                    xlWorkSheet.Cells[1, 2] = "Stand: " + DateTime.Now.ToString("d/M/yyyy");
                    xlWorkSheet.Cells[1, 3] = "Kurzw.";
                    xlWorkSheet.Cells[1, 4] = "Handy";
                    xlWorkSheet.Cells[1, 5] = "Dw";
                    xlWorkSheet.Cells[1, 6] = "Stand: " + DateTime.Now.ToString("d/M/yyyy");
                    xlWorkSheet.Cells[1, 7] = "Kurzw.";
                    xlWorkSheet.Cells[1, 8] = "Handy";
                    int zahler = 0;
                    int spalte1_neu = (anzahl / 2) + 6;
                    int mit_fax = anzahl + 6;
                    int rechts = 0;
                    int links = 0;
                    if (anzahl % 2 != 0)
                    {
                        rechts = ((anzahl + 6) / 2) + 1;
                        links = ((anzahl - 6) / 2);
                    }
                    else
                    {
                        rechts = (anzahl + 6) / 2;
                        links = (anzahl - 6) / 2;
                    }

                    for (int i = 0; i < rechts; i++)
                    {
                        xlWorkSheet.Cells[i + 2, 1] = durchwahl[i];
                        xlWorkSheet.Cells[i + 2, 2] = NachVor[i];
                        xlWorkSheet.Cells[i + 2, 3] = kurzwahl[i];
                        xlWorkSheet.Cells[i + 2, 4] = handy[i];
                    }
                    int zwischen = anzahl - rechts;


                    for (int i = 0; i < links; i++)
                    {
                        xlWorkSheet.Cells[i + 2, 5] = durchwahl[j + links];
                        xlWorkSheet.Cells[i + 2, 6] = NachVor[j + links];
                        xlWorkSheet.Cells[i + 2, 7] = kurzwahl[j + links];
                        xlWorkSheet.Cells[i + 2, 8] = handy[j + links];
                        j++;
                        if (i + 1 == links)
                        {
                            xlWorkSheet.Cells[i + 3, 5] = "46";
                            xlWorkSheet.Cells[i + 3, 6] = "Fax Buchhaltung";
                            xlWorkSheet.Cells[i + 4, 6] = "Fax 1. STock Altbau";
                            xlWorkSheet.Cells[i + 4, 8] = "09422-5550";
                            xlWorkSheet.Cells[i + 5, 5] = "10";
                            xlWorkSheet.Cells[i + 5, 6] = "Fax 2. Stock Altbau";
                            xlWorkSheet.Cells[i + 6, 5] = "827";
                            xlWorkSheet.Cells[i + 6, 6] = "Fax 2. Stock Neubau";
                            xlWorkSheet.Cells[i + 7, 5] = "825";
                            xlWorkSheet.Cells[i + 7, 6] = "Fax Konstruktion";
                            xlWorkSheet.Cells[i + 8, 5] = "828";
                            xlWorkSheet.Cells[i + 8, 6] = "Fax Herr Schnupp";
                        }
                        //xlWorkSheet.Cells.AutoFit();
                    }
                    chartRange = xlWorkSheet.get_Range("a1", "h1");
                    chartRange = xlWorkSheet.get_Range("a1", "h1");
                    chartRange.Font.Bold = true;
                    int rand = (anzahl / 2) + 7;
                    chartRange = xlWorkSheet.get_Range("a1", "h" + (spalte1 + 4));


                    chartRange.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic);
                    foreach (Microsoft.Office.Interop.Excel.Range cell in chartRange.Rows[1].Cells)
                    {
                        cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        cell.Font.Bold = true;
                    }
                    chartRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xlApp.DisplayAlerts = false;
                    xlWorkSheet.Columns["A:J"].AutoFit();
                    xlWorkSheet.PageSetup.Zoom = false;
                    xlWorkBook.SaveAs(@"M:\Kollegen\Telefonlisten\Telefon2spaltig.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, @"M:\Kollegen\Telefonlisten\Telefon2spaltig.pdf");
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                    releaseObject(xlApp);
                    releaseObject(xlWorkBook);
                    releaseObject(xlWorkSheet); var p = new System.Diagnostics.Process();
                    p.StartInfo = new ProcessStartInfo(@"M:\Kollegen\Telefonlisten\Telefon2spaltig.xls")
                    {
                        UseShellExecute = true
                    };
                    p.Start();
                }
               catch(Exception ex)
                {
                    MessageBox.Show("Die Datei ist gerade noch geöffnet, bitte versuchen sie es zu einem späteren Zeitpunkt erneut");
                }
                }
                void releaseObject(object obj)
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                        obj = null;
                    }
                    catch (Exception ex)
                    {
                        obj = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                }
            }
        }
    }

