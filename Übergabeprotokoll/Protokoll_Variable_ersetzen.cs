using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Net.Sockets;
using System.Threading;
using System.Windows.Forms;

namespace Übergabeprotokoll
{
    class Protokoll_Variable_ersetzen
    {
        public static void ersetzen(string Dateiname, string SpeicherpfadProtokolle)
        {
            Action t_Datei_loeschen = () =>
            {
                System.IO.File.Delete(SpeicherpfadProtokolle + Dateiname + ".docx");
            };

            Thread t_variablen_ersetzen = new Thread(() =>
            {
                try
                {
                    frmHauptprogramm frmHauptprogramm = System.Windows.Forms.Application.OpenForms[0] as frmHauptprogramm;
                    if (frmHauptprogramm == null)
                    {
                        MessageBox.Show("Keine Hauptform vorhanden!");
                        return;
                    }

                    if (frmHauptprogramm.ckUebergabe.Checked == true)
                    {
                        //J:\IT\01_Protokolle\01_Übergabeprotokole
                        Dateiname = frmHauptprogramm.lblDatum.Text + "_Übergabe_" + frmHauptprogramm.txtNachname.Text + "_" + frmHauptprogramm.txtVorname.Text;

                    }
                    else if (frmHauptprogramm.ckRueckgabe.Checked == true)
                    {
                        //J:\IT\01_Protokolle\01_Übergabeprotokole
                        Dateiname = frmHauptprogramm.lblDatum.Text + "_Rückgabe_" + frmHauptprogramm.txtNachname.Text + "_" + frmHauptprogramm.txtVorname.Text;

                    }


                    File.WriteAllBytes(SpeicherpfadProtokolle + Dateiname + ".docx", Übergabeprotokoll.Properties.Resources.Vorlage_mit_Variablen);

                    object fileName = SpeicherpfadProtokolle + Dateiname + ".docx";


                    Type wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType != null) // Check if the type is valid
                    {
                        dynamic msword = Activator.CreateInstance(wordType);
                        // Do something with msword
                    }
                    else
                    {
                        MessageBox.Show("Word ist nicht installiert!");
                        return;
                    }

                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };
                    Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);
                    Microsoft.Office.Interop.Word.Document aDoc2 = wordApp.Documents.Open(fileName, ReadOnly: false, Visible: false);

                    aDoc.Activate();

                    FindAndReplace(wordApp, "{DisplayName}", frmHauptprogramm.txtNachname.Text + " " + frmHauptprogramm.txtVorname.Text);
                    FindAndReplace(wordApp, "{DATUM}", frmHauptprogramm.lblDatum.Text);

                    //Notebook
                    if (frmHauptprogramm.ausgabe_Rueckgabe_NotebookComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{NB_AUSGABE}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{NB_AUSGABE}", frmHauptprogramm.ausgabe_Rueckgabe_NotebookComboBox.Text);
                    }

                    if (frmHauptprogramm.notebook_ModellComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{NB_MODELL}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{NB_MODELL}", frmHauptprogramm.notebook_ModellComboBox.Text);
                    }

                    if (frmHauptprogramm.notebook_SeriennummerTextBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{NB_SN}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{NB_SN}", frmHauptprogramm.notebook_SeriennummerTextBox.Text);
                    }

                    if (frmHauptprogramm.notebook_InventarnummerTextBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{NB_IV}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{NB_IV}", frmHauptprogramm.notebook_InventarnummerTextBox.Text);
                    }

                    if (frmHauptprogramm.zustand_NotebookComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{NB_ZUST}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{NB_ZUST}", frmHauptprogramm.zustand_NotebookComboBox.Text);
                    }

                    if (frmHauptprogramm.ausgabe_Rueckgabe_ZubehorComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{ZB_AUSGABE}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{ZB_AUSGABE}", frmHauptprogramm.ausgabe_Rueckgabe_ZubehorComboBox.Text);
                    }

                    if (frmHauptprogramm.zusatzliche_DockingstationComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{ZB_TYP1}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{ZB_TYP1}", frmHauptprogramm.zusatzliche_DockingstationComboBox.Text);
                    }


                    if (frmHauptprogramm.zusatzliche_LadegeratComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{ZB_TYP2}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{ZB_TYP2}", frmHauptprogramm.zusatzliche_LadegeratComboBox.Text);
                    }

                    if (frmHauptprogramm.zustand_ZubehorComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{ZB_ZUST}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{ZB_ZUST}", frmHauptprogramm.zustand_ZubehorComboBox.Text);
                    }

                    if (frmHauptprogramm.ausgabe_Rueckgabe_SmartphoneComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{SP_AUSGABE}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{SP_AUSGABE}", frmHauptprogramm.ausgabe_Rueckgabe_SmartphoneComboBox.Text);
                    }

                    if (frmHauptprogramm.smartphone_ModellComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{SP_MODELL}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{SP_MODELL}", frmHauptprogramm.smartphone_ModellComboBox.Text);
                    }


                    if (frmHauptprogramm.smartphone_SeriennummerTextBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{SP_SN}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{SP_SN}", frmHauptprogramm.smartphone_SeriennummerTextBox.Text);
                    }

                    if (frmHauptprogramm.zustand_SmartphoneComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{SP_ZUST}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{SP_ZUST}", frmHauptprogramm.zustand_SmartphoneComboBox.Text);
                    }

                    if (frmHauptprogramm.ausgabe_Rueckgabe_SIMComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{SIM_AUSGABE}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{SIM_AUSGABE}", frmHauptprogramm.ausgabe_Rueckgabe_SIMComboBox.Text);
                    }

                    if (frmHauptprogramm.sim_SeriennummerTextBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{SIM_SN}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{SIM_SN}", frmHauptprogramm.sim_SeriennummerTextBox.Text);
                    }

                    if (frmHauptprogramm.sim_TelefonnummerTextBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{SIM_RUFNUMMER}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{SIM_RUFNUMMER}", frmHauptprogramm.sim_TelefonnummerTextBox.Text);
                    }

                    var count = frmHauptprogramm.anmerkungenTextBox.Lines.Length;

                    //wenn zähler null, dann beide zeilen löschen
                    //wenn zähler 1 dann nur zeile 2 löschen
                    //wenn zähler 2 dann keine zeile löschen
                    //wenn zähler 3 dann 

                    if (count == 0)
                    {
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT}", "");
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_2}", "");
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_3}", "");
                    }
                    else if (count == 1)
                    {
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT}", frmHauptprogramm.anmerkungenTextBox.Lines[0].ToString());
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_2}", "");
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_3}", "");
                    }
                    else if (count == 2)
                    {
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT}", frmHauptprogramm.anmerkungenTextBox.Lines[0].ToString());
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_2}", frmHauptprogramm.anmerkungenTextBox.Lines[1].ToString());
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_3}", "");
                    }
                    else if (count == 3)
                    {
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT}", frmHauptprogramm.anmerkungenTextBox.Lines[0].ToString());
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_2}", frmHauptprogramm.anmerkungenTextBox.Lines[1].ToString());
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_3}", frmHauptprogramm.anmerkungenTextBox.Lines[2].ToString());
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT}", frmHauptprogramm.anmerkungenTextBox.Lines[0].ToString());
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_2}", frmHauptprogramm.anmerkungenTextBox.Lines[1].ToString());
                        FindAndReplace(wordApp, "{ANMERKUNGENTEXT_3}", frmHauptprogramm.anmerkungenTextBox.Lines[2].ToString());
                    }

                    if (frmHauptprogramm.ausgabe_Rueckgabe_DatenkarteComboBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{DK_AUSGABE}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{DK_AUSGABE}", frmHauptprogramm.ausgabe_Rueckgabe_DatenkarteComboBox.Text);
                    }

                    if (frmHauptprogramm.datenkarte_SeriennummerTextBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{DK_SN}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{DK_SN}", frmHauptprogramm.datenkarte_SeriennummerTextBox.Text);
                    }

                    if (frmHauptprogramm.txtNachname.Text is null)
                    {
                        FindAndReplace(wordApp, "{MITARBEITER}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{MITARBEITER}", frmHauptprogramm.txtNachname.Text + " " + frmHauptprogramm.txtVorname.Text);
                    }
                    if (frmHauptprogramm.agentTextBox.Text is null)
                    {
                        FindAndReplace(wordApp, "{IT_MITARBEITER}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{IT_MITARBEITER}", frmHauptprogramm.agentTextBox.Text);
                    }
                    if (frmHauptprogramm.txtBenutzername.Text is null)
                    {
                        FindAndReplace(wordApp, "{BENUTZERNAME}", "");

                    }
                    else
                    {
                        FindAndReplace(wordApp, "{BENUTZERNAME}", frmHauptprogramm.txtBenutzername.Text);
                    }
                    if (frmHauptprogramm.txtKennwort.Text is null)
                    {
                        FindAndReplace(wordApp, "{KENNWORT}", "");
                    }
                    else
                    {
                        FindAndReplace(wordApp, "{KENNWORT}", frmHauptprogramm.txtKennwort.Text);
                    }



                    aDoc.Save();
                    aDoc.Close();

                    Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                    aDoc2 = appWord.Documents.Open(SpeicherpfadProtokolle + Dateiname + ".docx");
                    aDoc2.ExportAsFixedFormat(SpeicherpfadProtokolle + Dateiname + ".pdf", WdExportFormat.wdExportFormatPDF);

                    aDoc2.Save();
                    aDoc2.Close();


                    Socket s = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                    Socket s2 = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

                    if (s.Connected == true)
                    {
                        s.Disconnect(true);
                        s.Close();
                    }

                    if (s2.Connected == true)
                    {
                        s2.Disconnect(true);
                        s2.Close();
                    }

                    //datei wird 2 mal an den Drucker gesendet
                    //s.Connect(Drucker, 9100);
                    //s.SendFile(SpeicherpfadProtokolle + Dateiname + ".pdf");
                    //s.Disconnect(true);
                    //s.Close(2);


                    //s2.Connect(Drucker, 9100);
                    //s2.SendFile(SpeicherpfadProtokolle + Dateiname + ".pdf");
                    //s2.Disconnect(true);
                    //s2.Close(2);

                }
                finally
                {
                    t_Datei_loeschen();
                }

            });

            t_variablen_ersetzen.Start();
        }

        [STAThread]
        public static void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = false;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }
}
