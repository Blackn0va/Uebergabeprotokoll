using System.Data.SQLite;
using System;
using System.Data;
using System.Windows.Forms;

namespace Übergabeprotokoll
{
    public class SQL_Befehl
    {
       public static string con = "Data Source=Protokolle.db";
        // Create a connection to the database file
        public static SQLiteConnection connection = new SQLiteConnection(con);
        public static SQLiteCommand command = connection.CreateCommand();
        public static frmHauptprogramm frmHauptprogramm = System.Windows.Forms.Application.OpenForms[0] as frmHauptprogramm;


        [STAThread]
        public static void Delete()
        {
            if (frmHauptprogramm == null)
            {
                MessageBox.Show("Keine Hauptform vorhanden!");
                return;
            }

            DialogResult result = System.Windows.Forms.MessageBox.Show("Datensatz wirklich löschen?", "Abfrage", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                connection.Open();
                command.CommandText = "delete from TBL_Protokoll where ProtokollID='" + frmHauptprogramm.lblID.Text + "'";
                command.ExecuteNonQuery();
                connection.Close();
                Display_Data();
                frmHauptprogramm.txtSuchen.Text = "";
            }
            try
            {

            }
            catch { }
        }

        [STAThread]
        public static void Search()
        {
            try
            {
                if (frmHauptprogramm == null)
                {
                    MessageBox.Show("Keine Hauptform vorhanden!");
                    return;
                }

                connection.Open();
                command.CommandText = "select * from TBL_Protokoll where Nachname like" + "'%" + frmHauptprogramm.txtSuchen.Text + "%'" + "OR Vorname like" + "'%" + frmHauptprogramm.txtSuchen.Text + "%'";
                System.Data.DataTable dt = new System.Data.DataTable();
                SQLiteDataAdapter da = new SQLiteDataAdapter(command);
                da.Fill(dt);
                frmHauptprogramm.dgvProtokolle.DataSource = dt;
                connection.Close();
                frmHauptprogramm.txtSuchen.Text = "";
            }
            catch { }
        }

        [STAThread]
        public static void Update()
        {
            if (frmHauptprogramm == null)
            {
               MessageBox.Show("Keine Hauptform vorhanden!");
                return;
            }

                //cast text from frmHauptprogramm.lblID.Text to int
                int id = Convert.ToInt32(frmHauptprogramm.lblID.Text);
                connection.Open();
                command.CommandText = "update TBL_Protokoll set " +
                    " ProtokollID='" + id + "'," +
                    " Vorname='" + frmHauptprogramm.txtVorname.Text + "'," +
                    " Nachname='" + frmHauptprogramm.txtNachname.Text + "'," +
                    " Rueckgabe_notebook='" + frmHauptprogramm.ausgabe_Rueckgabe_NotebookComboBox.Text + "' ," +
                    " Notebook_modell='" + frmHauptprogramm.notebook_ModellComboBox.Text + "' ," +
                    " Notebook_Seriennummer='" + frmHauptprogramm.notebook_SeriennummerTextBox.Text + "' ," +
                    " Notebook_Inventarnummer='" + frmHauptprogramm.notebook_InventarnummerTextBox.Text + "' ," +
                    " Zustand_notebook='" + frmHauptprogramm.zustand_NotebookComboBox.Text + "' ," +
                    " Rueckgabe_zubehor='" + frmHauptprogramm.ausgabe_Rueckgabe_ZubehorComboBox.Text + "' ," +
                    " Dockingstation='" + frmHauptprogramm.zusatzliche_DockingstationComboBox.Text + "' ," +
                    " Ladegerat='" + frmHauptprogramm.zusatzliche_LadegeratComboBox.Text + "' ," +
                    " Zustand_zubehor='" + frmHauptprogramm.zustand_ZubehorComboBox.Text + "' ," +
                    " Rueckgabe_smartphone='" + frmHauptprogramm.ausgabe_Rueckgabe_SmartphoneComboBox.Text + "' ," +
                    " Smartphone_modell='" + frmHauptprogramm.smartphone_ModellComboBox.Text + "' ," +
                    " Smartphone_Seriennummer='" + frmHauptprogramm.smartphone_SeriennummerTextBox.Text + "' ," +
                    " Zustand_smartphone='" + frmHauptprogramm.zustand_SmartphoneComboBox.Text + "' ," +
                    " Rueckgabe_sim='" + frmHauptprogramm.ausgabe_Rueckgabe_SIMComboBox.Text + "' ," +
                    " SIM_Seriennummer='" + frmHauptprogramm.sim_SeriennummerTextBox.Text + "' ," +
                    " SIM_Telefonnummer='" + frmHauptprogramm.sim_TelefonnummerTextBox.Text + "' ," +
                    " Rueckgabe_datenkarte='" + frmHauptprogramm.ausgabe_Rueckgabe_DatenkarteComboBox.Text + "' ," +
                    " Datenkarte_Seriennummer='" + frmHauptprogramm.datenkarte_SeriennummerTextBox.Text + "' ," +
                    " Anmerkungen='" + frmHauptprogramm.anmerkungenTextBox.Text + "' ," +
                    " Datum='" + frmHauptprogramm.lblDatum.Text + "' ," +
                    " Kennwort='" + frmHauptprogramm.txtKennwort.Text + "' ," +
                    " Benutzername='" + frmHauptprogramm.txtBenutzername.Text + "' ," +
                    " Agent='" + frmHauptprogramm.agentTextBox.Text + "' " +
                    "where ProtokollID='" + frmHauptprogramm.lblID.Text + "'";
                command.ExecuteNonQuery();
                connection.Close();
                Display_Data();
  
        }

        [STAThread]
        public static void Speichern()
        {
            try
            {
                if (frmHauptprogramm == null)
                {
                    MessageBox.Show("Keine Hauptform vorhanden!");
                    return;
                }

                if (frmHauptprogramm.txtVorname.Text != "" & frmHauptprogramm.txtNachname.Text != "")
                {

                    connection.Open();
                    command.CommandText = "insert into TBL_Protokoll values(null,'"  + frmHauptprogramm.txtVorname.Text + "'," +
                        " '" + frmHauptprogramm.txtNachname.Text + "'," +
                        "'" + frmHauptprogramm.ausgabe_Rueckgabe_NotebookComboBox.Text + "'," +
                        " '" + frmHauptprogramm.notebook_ModellComboBox.Text + "'," +
                        " '" + frmHauptprogramm.notebook_SeriennummerTextBox.Text + "'," +
                        " '" + frmHauptprogramm.notebook_InventarnummerTextBox.Text + "'," +
                        " '" + frmHauptprogramm.zustand_NotebookComboBox.Text + "'," +
                        " '" + frmHauptprogramm.ausgabe_Rueckgabe_ZubehorComboBox.Text + "'," +
                        " '" + frmHauptprogramm.zusatzliche_DockingstationComboBox.Text + "'," +
                        " '" + frmHauptprogramm.zusatzliche_LadegeratComboBox.Text + "'," +
                        " '" + frmHauptprogramm.zustand_ZubehorComboBox.Text + "'," +
                        " '" + frmHauptprogramm.ausgabe_Rueckgabe_SmartphoneComboBox.Text + "'," +
                        " '" + frmHauptprogramm.smartphone_ModellComboBox.Text + "'," +
                        " '" + frmHauptprogramm.smartphone_SeriennummerTextBox.Text + "'," +
                        " '" + frmHauptprogramm.zustand_SmartphoneComboBox.Text + "'," +
                        " '" + frmHauptprogramm.ausgabe_Rueckgabe_SIMComboBox.Text + "'," +
                        " '" + frmHauptprogramm.sim_SeriennummerTextBox.Text + "'," +
                        " '" + frmHauptprogramm.sim_TelefonnummerTextBox.Text + "'," +
                        " '" + frmHauptprogramm.ausgabe_Rueckgabe_DatenkarteComboBox.Text + "'," +
                        " '" + frmHauptprogramm.datenkarte_SeriennummerTextBox.Text + "'," +
                        " '" + frmHauptprogramm.anmerkungenTextBox.Text + "'," +
                        " '" + frmHauptprogramm.lblDatum.Text + "'," +
                        " '" + frmHauptprogramm.agentTextBox.Text + "'," +
                        " '" + frmHauptprogramm.txtBenutzername.Text + "'," +
                        " '" + frmHauptprogramm.txtKennwort.Text + "')";
                    command.ExecuteNonQuery();
                    connection.Close();
                    Display_Data();

                   FelderZuruckSetzen();
                }
                else
                {
                     MessageBox.Show("Bitte geben Sie mindestens einen namen ein");
                }

            }
            catch { }
        }

        [STAThread]
        public static void Display_Data()
        {
            try
            {
                connection.Open();

                // Create a command to select all data from the table
                command.CommandText = "select * from TBL_Protokoll";

                // Create a data table and a data adapter to fill it with the data
                DataTable dt = new DataTable();
        
               //add data to dataadapter
               SQLiteDataAdapter da = new SQLiteDataAdapter(command);


                da.Fill(dt);



                // Check if the form variable is not null
                if (frmHauptprogramm == null)
                {
                    MessageBox.Show("Keine Hauptform vorhanden!");
                    return;
                }


                frmHauptprogramm.dgvProtokolle.DataSource = dt;
                connection.Close();
            }
            catch { }

        }

        [STAThread]
        public static void CreateDatabase()
        {
 
            // Open the connection
            connection.Open();

   

            // Create the table if it does not exist
            command.CommandText = @"CREATE TABLE IF NOT EXISTS TBL_Protokoll (
                ProtokollID INTEGER primary key AUTOINCREMENT,
                Vorname varchar(20),
                Nachname varchar(20),
                Rueckgabe_notebook varchar(20),
                Notebook_modell varchar(20),
                Notebook_seriennummer varchar(20),
                Notebook_inventarnummer varchar(20),
                Zustand_notebook varchar(20),
                Rueckgabe_zubehor varchar(20),
                Dockingstation varchar(20),
                Ladegerat varchar(20),
                Zustand_zubehor varchar(20),
                Rueckgabe_smartphone varchar(20),
                Smartphone_modell varchar(20),
                Smartphone_seriennummer varchar(20),
                Zustand_smartphone varchar(20),
                Rueckgabe_sim varchar(20),
                Sim_seriennummer varchar(20),
                Sim_telefonnummer varchar(20),
                Rueckgabe_datenkarte varchar(20),
                Datenkarte_seriennummer varchar(20),
                Anmerkungen varchar(100),
                Datum date,
                Agent varchar(20),
                Benutzername varchar(20),
                Kennwort varchar(20)
            )";

            // Execute the command
            command.ExecuteNonQuery();

            // Close the connection
            connection.Close();

           Display_Data();
        }

        [STAThread]
        private static void FelderZuruckSetzen()
        {
            try
            {
                if (frmHauptprogramm == null)
                {
                    MessageBox.Show("Keine Hauptform vorhanden!");
                    return;
                }

                frmHauptprogramm.txtVorname.Text = "";
                frmHauptprogramm.txtNachname.Text = "";
                frmHauptprogramm.ausgabe_Rueckgabe_NotebookComboBox.SelectedIndex = -1;
                frmHauptprogramm.notebook_ModellComboBox.SelectedIndex = -1;
                frmHauptprogramm.notebook_SeriennummerTextBox.Text = "";
                frmHauptprogramm.notebook_InventarnummerTextBox.Text = "";
                frmHauptprogramm.zustand_NotebookComboBox.SelectedIndex = -1;
                frmHauptprogramm.ausgabe_Rueckgabe_ZubehorComboBox.SelectedIndex = -1;
                frmHauptprogramm.zusatzliche_DockingstationComboBox.SelectedIndex = -1;
                frmHauptprogramm.zusatzliche_LadegeratComboBox.SelectedIndex = -1;
                frmHauptprogramm.zustand_ZubehorComboBox.SelectedIndex = -1;
                frmHauptprogramm.ausgabe_Rueckgabe_SmartphoneComboBox.SelectedIndex = -1;
                frmHauptprogramm.smartphone_ModellComboBox.SelectedIndex = -1;
                frmHauptprogramm.smartphone_SeriennummerTextBox.Text = "";
                frmHauptprogramm.zustand_SmartphoneComboBox.SelectedIndex = -1;
                frmHauptprogramm.ausgabe_Rueckgabe_SIMComboBox.SelectedIndex = -1;
                frmHauptprogramm.sim_SeriennummerTextBox.Text = "";
                frmHauptprogramm.sim_TelefonnummerTextBox.Text = "";
                frmHauptprogramm.ausgabe_Rueckgabe_DatenkarteComboBox.SelectedIndex = -1;
                frmHauptprogramm.datenkarte_SeriennummerTextBox.Text = "";
                frmHauptprogramm.anmerkungenTextBox.Text = "";
                frmHauptprogramm.agentTextBox.Text = "";
                frmHauptprogramm.txtKennwort.Text = "";
                frmHauptprogramm.txtBenutzername.Text = "";
            }
            catch { }
        }
    }
}
