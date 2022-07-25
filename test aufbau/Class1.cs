using System;
using System.Data;
using System.Data.SqlClient;
using System.DirectoryServices.AccountManagement;
using System.Windows;
using DirectoryEntry = System.DirectoryServices.DirectoryEntry;


namespace test_aufbau
{
    // Hier findet die Überprüfung statt
    internal class Class1 : Helper_DB
    {
        
        public static bool pruefen(string id, string vorname, bool vorname_p, bool nachname_p, string nachname, string handy, bool kurzwahl_p, int n, string durchwahl_string, string kurzwahl_string, bool durchwahl_p)
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "schnupp.de");
            //nur wenn der Vorname weniger als 1 Buchstaben hat
            if (vorname.Length < 1)
            {
                MessageBox.Show("Bitte Geben Sie einen Vornamen ein");
            }
            else
            {
                //Vorname hat mehr als 1 Buchstaben 
                if (vorname_p == true)
                {
                    //es wird geschaut, sind es nur Buchstaben oder auch Zeichen 
                    MessageBox.Show("Zahlen sind im Vorname nicht erlaubt");
                }
                else
                {
                    //Vorname hat nur Buchstaben und nachname wird geprüfut ob er Zahlen hat
                    if (nachname_p == true)
                    {
                        MessageBox.Show("Zahlen sind im Nachnamen nicht erlaubt");
                    }
                    else
                    {
                        //Nachname hat keine Zahlen, es wird geprüft ob die länge des Nachnamens länger als 0 ist
                        if (nachname.Length < 1)
                        {
                            MessageBox.Show("Bitte Geben Sie einen Nachnamen ein");
                        }
                        else
                        {
                            //Nachname ist min 1 Buchstaben lang
                            int a;
                            bool istzahl = true;
                            int kurwzahl_int = 0;
                            istzahl = int.TryParse(kurzwahl_string, out a);
                            //Es wird geprüft, ob die kurzwahl eine Zahl ist
                            if(istzahl == true )
                            {
                                // die Kurzwahl wird in einen INT konvertiert, aber nur wenn es keine Buchstaben hat
                                kurwzahl_int = Convert.ToInt32(kurzwahl_string);
                            }
                            //Kurzwahl muss zwischen eine Zahl am anfang haben mit 60 und darf nur 5 stellig sein oder eine länge von 0
                            if (kurzwahl_p == true  && kurwzahl_int < 60000 ||  kurzwahl_p == true &&  kurwzahl_int > 60999 )
                            {
                                if(kurwzahl_int == 0)
                                {

                                }
                                else if(kurwzahl_int != 0 && kurwzahl_int != 5 && kurwzahl_int < 60000 ||kurwzahl_int > 60999)

                                {
                                    MessageBox.Show("Bitte geben Sie bei Kurzwahl eine Zahl ein die mit 60 beginnt und min 5 zahlen besitzt");
                                }
                            }
                            //wenn die Kurzwahl Buchstaben hat 
                            if(kurzwahl_p == false && kurzwahl_string != "")
                            {
                                MessageBox.Show("Es sind nur Zahlen erlaubt"); 
                            }
                            if(kurzwahl_string.Length== 0 || kurzwahl_p == true && kurwzahl_int >= 60000 && kurwzahl_int <= 60999 && kurzwahl_string.Length == 5) 
                            {
                                if(durchwahl_p == false && durchwahl_string.Length != 0)
                                {
                                    MessageBox.Show("Es sind nur Zahlen erlaubt !");
                                }
                                //durchwahl darf nur 2 Zeichen haben oder gar keine
                              if (durchwahl_p == true && durchwahl_string.Length > 3 || durchwahl_p == true && durchwahl_string.Length ==1)
                                {
                                    MessageBox.Show("Bitte geben Sie eine Durchwahl ein, die 2 Zeichen besitzt");
                                }
                              //es wird geprüft, ob eine der 3 felder gefüllt sind, weil 1 min gefüllt sein muss
                                else if(durchwahl_p == true && durchwahl_string.Length == 2 || durchwahl_p == true && durchwahl_string.Length == 3  ||  durchwahl_string.Length==0)
                                {
                                    if (durchwahl_string.Length == 0 && handy.Length ==0  && kurzwahl_string == "")

                                    {
                                        MessageBox.Show("Eines der Felder Kurzwahl, Durchwahl oder Handy muss gefüllt werden");
                                    }
                                    //gibt true zurück und stoßt das jeweilige SQL statement somit ab und gibt zurück das die Prüfung erfolgreich war
                                   else if (durchwahl_string.Length == 2 || durchwahl_string.Length == 2 || durchwahl_string.Length == 0 || handy.Length >1 || kurzwahl_string.Length >1) //
                                    {
                                            return true;
                                    }
                                    else  if(durchwahl_string.Length !=2 && durchwahl_string.Length != 0 && durchwahl_string.Length != 2)
                                    {
                                        //gibt false zurück und sagt das die Prüfung nicht erfolgreich war und es kommt zu keinem SQl Statement
                                        return false;
                                    }
                                    return false;
                                } 
                                return false;
                             }
                            return false;
                        }
                        return false;
                    } 
                    return false;
                }
                return false;
            }
            return false;

        }

    // Es wird geschaut, ob derjenige, der das Programm ausführt die Berechtigung hat die SQl Statements durchzuführen
        public static bool IsInGroup()
        {
            // set up domain context
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "schnupp.de");
            // find a user
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, "SomeUserName");
            // find the group in question
            GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, "Role_IT");
            return true;
        }
        //SQL Satement mit Update
        public static void sqlUpdate(string id, string vorname, bool vorname_p, bool nachname_p, string nachname, string handy, bool kurzwahl_p, int n, string durchwahl_string, string kurzwahl_string, bool durchwahl_p)
        {
            using (SqlConnection conn = new SqlConnection(db_connection()))
            {
                using (SqlCommand cmd = new SqlCommand(@"UPDATE tbl_Telefonnummern set   Handy=" + handy + ", Nachname= '" + nachname + "', DW= '" + durchwahl_string + "', kurzw ='" + kurzwahl_string + " ' where ID ='" + id + "'", conn))
                {
                    cmd.CommandType = CommandType.Text;
                    conn.Open();
                    int rowsAffected = cmd.ExecuteNonQuery();
                    conn.Close();
                    MessageBox.Show("Telefonnummer wurde erfolgreich bearbeitet ");
                    
                }
            }
        }
        //SQL Statement mit Delete
        public static void sqlDelete(string geben, string Nachname, string ID)
        {
            using (SqlConnection conn = new SqlConnection(db_connection()))
            {
                conn.Open();
                string[] authorlist = geben.Split(" ");
                Nachname = authorlist[0];
                ID = authorlist[2];
                SqlCommand cmd = new SqlCommand("Delete  from tbl_Telefonnummern where ID=" + ID + "", conn);
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                }
                reader.Close();
            }
        }
  // SQL Statement mit Insert
        public static void sqlInsert(string id, string vorname, bool vorname_p, bool nachname_p, string nachname, string handy, bool kurzwahl_p, int n, string durchwahl_string, string kurzwahl_string, bool durchwahl_p)
        {
            System.Data.SqlClient.SqlConnection sqlConnection1 =
                                   new System.Data.SqlClient.SqlConnection(db_connection());
            System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = " Insert tbl_Telefonnummern (Vorname, Nachname, DW, Kurzw ,Handy) VALUES ('" + vorname + "','" + nachname + "','" + durchwahl_string + "','" + kurzwahl_string + "','" + handy + "')";
            cmd.Connection = sqlConnection1;
            sqlConnection1.Open();
            cmd.ExecuteNonQuery();
            sqlConnection1.Close();
            MessageBox.Show("Die Telefonnummern wurden erfolgreich Hinzugefügt");
        }
        public static string getUsername(string vorname, string nachname)
        {
            string vorname_null = vorname[0].ToString().ToLower();
            return vorname_null + "." + nachname.ToLower();
        }
    }
}
                
