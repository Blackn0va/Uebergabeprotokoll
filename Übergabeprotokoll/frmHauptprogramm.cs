using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;


namespace Übergabeprotokoll
{
    public partial class frmHauptprogramm : Form
    {
        //get workingdirectory
         string SpeicherpfadProtokolle = Environment.CurrentDirectory + @"\Protokolle\";
        string Dateiname = "";

         public frmHauptprogramm()
        {
            InitializeComponent();
            SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            DoubleBuffered = true;
            AutoScaleMode = 0;
        }

        [STAThread]
        private void frmHauptprogramm_Load(object sender, EventArgs e)
        {
            //CheckForIllegalCrossThreadCalls = false;

            //Software Version FormText anzeigen
            this.Text = "Übergabeprotokoll Tool - " + this.ProductVersion.ToString();

            this.StartPosition = FormStartPosition.WindowsDefaultLocation;
            typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic |
            BindingFlags.Instance | BindingFlags.SetProperty, null,
            dgvProtokolle, new object[] { true });

            //das Label beinhaltet das Datum in kurzform
            lblDatum.Text = Convert.ToString(DateTime.Now.ToString("yyyy-MM-dd"));

            //string to workingDirectory
            string workingDirectory =  Environment.CurrentDirectory;
            //check if Directory exists
            if (!System.IO.Directory.Exists(workingDirectory + @"\Protokolle\"))
            {
                //if not create it
                System.IO.Directory.CreateDirectory(workingDirectory + @"\Protokolle\");
            }
            System.IO.DirectoryInfo ParentDirectory = new System.IO.DirectoryInfo(workingDirectory + @"\Protokolle\");

            //Dateien in Verzeichnis einlesen und anzeigen
            foreach (System.IO.FileInfo f in ParentDirectory.GetFiles())
            {
                if (f.Name.Contains(".pdf"))
                    listBox1.Items.Add(f.Name);
            }

            //wenn dgvProtokolle nicht leer ist zum letztenDatensatz springen
            if (dgvProtokolle.Rows.Count > 0)
            {
                dgvProtokolle.FirstDisplayedCell = dgvProtokolle.Rows[dgvProtokolle.RowCount - 1].Cells[0];
            }



            //datenbank erstellen wenn nicht vorhanden
            SQL_Befehl.CreateDatabase();

            SQL_Befehl.Display_Data();

        }

        [STAThread]
        private void cmdOK_Click(object sender, EventArgs e)
        {
            SQL_Befehl.Speichern();
        }

        [STAThread]
        private void cmdUpdate_Click(object sender, EventArgs e)
        {
            SQL_Befehl.Update();

        }

        [STAThread]
        private void cmdDelete_Click(object sender, EventArgs e)
        {
            SQL_Befehl.Delete();

        }

        [STAThread]
        private void cmdErstellen(object sender, EventArgs e)
        {
 
                    Protokoll_Variable_ersetzen.ersetzen(Dateiname, SpeicherpfadProtokolle);
  
        }

        [STAThread]
        private void cmdSearch_Click(object sender, EventArgs e)
        {
            SQL_Befehl.Search();

        }

        [STAThread]
        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                Process.Start(SpeicherpfadProtokolle + @"" + listBox1.GetItemText(listBox1.SelectedItem));
            }
            catch { }
        }

        [STAThread]
        private void dgvProtokolle_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            this.lblID.Text = dgvProtokolle.Rows[e.RowIndex].Cells[0].Value.ToString();
            this.txtVorname.Text = dgvProtokolle.Rows[e.RowIndex].Cells[1].Value.ToString();
            this.txtNachname.Text = dgvProtokolle.Rows[e.RowIndex].Cells[2].Value.ToString();
            this.ausgabe_Rueckgabe_NotebookComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[3].Value.ToString();
            this.notebook_ModellComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[4].Value.ToString();
            this.notebook_SeriennummerTextBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[5].Value.ToString();
            this.notebook_InventarnummerTextBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[6].Value.ToString();
            this.zustand_NotebookComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[7].Value.ToString();
            this.ausgabe_Rueckgabe_ZubehorComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[8].Value.ToString();
            this.zusatzliche_DockingstationComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[9].Value.ToString();
            this.zusatzliche_LadegeratComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[10].Value.ToString();
            this.zustand_ZubehorComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[11].Value.ToString();
            this.ausgabe_Rueckgabe_SmartphoneComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[12].Value.ToString();
            this.smartphone_ModellComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[13].Value.ToString();
            this.smartphone_SeriennummerTextBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[14].Value.ToString();
            this.zustand_SmartphoneComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[15].Value.ToString();
            this.ausgabe_Rueckgabe_SIMComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[16].Value.ToString();
            this.sim_SeriennummerTextBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[17].Value.ToString();
            this.sim_TelefonnummerTextBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[18].Value.ToString();
            this.ausgabe_Rueckgabe_DatenkarteComboBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[19].Value.ToString();
            this.datenkarte_SeriennummerTextBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[20].Value.ToString();
            this.anmerkungenTextBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[21].Value.ToString();
            this.agentTextBox.Text = dgvProtokolle.Rows[e.RowIndex].Cells[23].Value.ToString();
            this.txtBenutzername.Text = dgvProtokolle.Rows[e.RowIndex].Cells[24].Value.ToString();
            this.txtKennwort.Text = dgvProtokolle.Rows[e.RowIndex].Cells[25].Value.ToString();
        }

        [STAThread]
        private void ckUebergabe_CheckedChanged(object sender, EventArgs e)
        {
            if (ckUebergabe.Checked == true)
            {
                ckRueckgabe.Checked = false;
            }
        }

        [STAThread]
        private void ckRueckgabe_CheckedChanged(object sender, EventArgs e)
        {
            if (ckRueckgabe.Checked == true)
            {
                ckUebergabe.Checked = false;
            }
        }

        [STAThread]
        private void txtSuchen_KeyPress(object sender, KeyPressEventArgs e)
        {
            //wenn die taste Enter gedrückt wurde
            if (e.KeyChar == (char)13)
            {
                SQL_Befehl.Search();
            }
        }

        [STAThread]
        private void cmdPdferstellen_Click(object sender, EventArgs e)
        {
            try
            {
                Protokoll_Variable_ersetzen.ersetzen(Dateiname, SpeicherpfadProtokolle);
            }
            catch { }
        }

        [STAThread]
        public static void SetDoubleBuffered(Control control)
        {
            // set instance non-public property with name "DoubleBuffered" to true
            typeof(Control).InvokeMember("DoubleBuffered",
                BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.NonPublic,
                null, control, new object[] { true });
        }

        private void tabProtokolle_Enter(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            this.Location = new Point(frmHauptprogramm.MousePosition.X, frmHauptprogramm.MousePosition.Y);


        }
    }
}
