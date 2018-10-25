using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportTM
{
    public partial class frmImport : Form
    {
        Excel.Application xlApplication;
        Excel.Workbook xlWorkbook;
        Excel.Worksheet xlWorksheet;
        object misValue = System.Reflection.Missing.Value;
        private SqlConnection cnnOnderwijs = new SqlConnection("Data Source=" + Globals.DB_SERVER + ";Initial Catalog=" + Globals.DB_NAME + ";User ID=" + Globals.DB_USER + ";Password=" + Globals.DB_PASSWORD + ";MultipleActiveResultSets=true;");
        private int intDoelen = 0;

        public frmImport()
        {
            InitializeComponent();
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnOpenTM_Click(object sender, EventArgs e)
        {
            //Importeren in DB?
            bool blnImport = false;
            if (chkImportInDB.Checked)
            {
                blnImport = MessageBox.Show("Weet je zeker dat je wilt importeren in de DB? Checkbox staat aangevinkt!!!", "!!!LET OP!!!", MessageBoxButtons.OKCancel) == DialogResult.OK;
            }
            chkImportInDB.Checked = blnImport;

            bool blnVersieImport = false;
            if (chkVersie.Checked)
            {
                blnVersieImport = MessageBox.Show("Weet je zeker dat je het versiebeheer wilt importeren in de DB? Checkbox staat aangevinkt!!!", "!!!LET OP!!!", MessageBoxButtons.OKCancel) == DialogResult.OK;
            }
            chkVersie.Checked = blnVersieImport;

            bool blnToetsvormImport = false;
            if (chkToetsvorm.Checked)
            {
                blnToetsvormImport = MessageBox.Show("Weet je zeker dat je de toetsvorm wilt importeren in de DB? Checkbox staat aangevinkt!!!", "!!!LET OP!!!", MessageBoxButtons.OKCancel) == DialogResult.OK;
            }
            chkToetsvorm.Checked = blnToetsvormImport;


            //Initialiseren controls:
            txtSucces.Text = "";
            txtFail.Text = "";
            lstToetscodes.Items.Clear();

            //Selecteer een bestand zodat de map met toetsmatrijzen bekend is:
            if (openTM.FileName.Equals("<leeg>"))
            {
                openTM.FileName = "Toetsmatrijs";
            }
            openTM.ShowDialog();
            txtFilenameTM.Text = Path.GetDirectoryName(openTM.FileName);
            txtSucces.Text += "==================================================" + "\r\n";
            txtSucces.Text += "Map geselecteerd: " + Path.GetDirectoryName(openTM.FileName) + "\r\n";
            txtSucces.Text += "==================================================" + "\r\n";

            //Haal alle bestandsnamen uit de geselecteerde map:
            String[] filePaths = Directory.GetFiles(Path.GetDirectoryName(openTM.FileName), "*.xlsx", SearchOption.TopDirectoryOnly);
            for (int t = 0; t < filePaths.Length; t++)
            {
                lstToetscodes.Items.Add(filePaths[t]);
            }
            frmImport.ActiveForm.Refresh();

            for (int t = 0; t < filePaths.Length; t++)
            {

                //Initialiseren controls:
                intDoelen = 0;
                int fkToets = -1;
                foreach (Control ctrInstance in Controls)
                {
                    switch (ctrInstance.GetType().Name)
                    {
                        case "TextBox":
                            if (((TextBox)ctrInstance).Name != "txtSucces" && ((TextBox)ctrInstance).Name != "txtFail")
                            {
                                ((TextBox)ctrInstance).Text = "";
                            }
                            break;
                        case "ListBox":
                            if (((ListBox)ctrInstance).Name != "lstToetscodes")
                            {
                                ((ListBox)ctrInstance).Items.Clear();
                            }
                            break;
                        default:
                            break;
                    }
                }

                //Open bestand:
                xlApplication = new Excel.Application();
                xlWorkbook = xlApplication.Workbooks.Open(filePaths[t], 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                txtSucces.Text += "Bestand geopend: " + filePaths[t] + "\r\n";
                txtFail.Text += "Bestand geopend: " + filePaths[t] + "\r\n";

                //Toetscode:
                txtToetscode.Text = xlWorksheet.get_Range("B3", "B3").Value2.ToString();
                txtSucces.Text += "Toetscode: " + txtToetscode.Text + "\r\n";
                using (SqlCommand cmdToets = new SqlCommand("SELECT * FROM tblToets WHERE strCode = '" + txtToetscode.Text + "'", cnnOnderwijs))
                {
                    using (SqlDataReader rdrToets = cmdToets.ExecuteReader())
                    {
                        if (rdrToets.HasRows)
                        {
                            rdrToets.Read();
                            fkToets = Convert.ToInt32(rdrToets["pkId"]);
                            txtPKToets.Text = rdrToets["pkId"].ToString();
                            txtSucces.Text += "Toetscode gevonden in DB, PK: " + txtPKToets.Text + "\r\n";
                        }
                        else
                        {
                            txtPKToets.Text = "???";
                            txtFail.Text += "Toetscode niet gevonden in DB: " + txtToetscode.Text + "\r\n";
                        }
                        rdrToets.Close();
                    }
                }

                //Toetsvorm:
                txtToetsvorm.Text = xlWorksheet.get_Range("B6", "B6").Value2.ToString();
                txtSucces.Text += "Toetsvorm: " + txtToetsvorm.Text + "\r\n";
                using (SqlCommand cmdToetsvorm = new SqlCommand("SELECT * FROM tblToetsvorm WHERE strVolledig = '" + txtToetsvorm.Text + "'", cnnOnderwijs))
                {
                    using (SqlDataReader rdrToetsvorm = cmdToetsvorm.ExecuteReader())
                    {
                        if (rdrToetsvorm.HasRows)
                        {
                            rdrToetsvorm.Read();
                            txtPKToetsvorm.Text = rdrToetsvorm["pkId"].ToString();
                            txtSucces.Text += "Toetsvorm gevonden in DB, PK: " + txtPKToetsvorm.Text + "\r\n";
                        }
                        else
                        {
                            txtPKToetsvorm.Text = "???";
                            txtFail.Text += "Toetsvorm niet gevonden in DB: " + txtToetsvorm.Text + "\r\n";
                        }
                        rdrToetsvorm.Close();
                    }
                }

                //Beoordelingswijze:
                txtBeoordelingswijze.Text = xlWorksheet.get_Range("B7", "B7").Value2.ToString();
                txtSucces.Text += "Beoordelingswijze: " + txtBeoordelingswijze.Text + "\r\n";
                using (SqlCommand cmdBeoordelingswijze = new SqlCommand("SELECT * FROM tblBeoordelingswijze WHERE strNaam = '" + txtBeoordelingswijze.Text + "'", cnnOnderwijs))
                {
                    using (SqlDataReader rdrBeoordelingswijze = cmdBeoordelingswijze.ExecuteReader())
                    {
                        if (rdrBeoordelingswijze.HasRows)
                        {
                            rdrBeoordelingswijze.Read();
                            txtPKBeoordelingswijze.Text = rdrBeoordelingswijze["pkId"].ToString();
                            txtSucces.Text += "Beoordelingswijze gevonden in DB, PK: " + txtPKBeoordelingswijze.Text + "\r\n";
                        }
                        else
                        {
                            txtPKBeoordelingswijze.Text = "???";
                            txtFail.Text += "Beoordelingswijze niet gevonden in DB: " + txtBeoordelingswijze.Text + "\r\n";
                        }
                        rdrBeoordelingswijze.Close();
                    }
                }

                //Leerdoelen (meteen ook het aantal doelen bepalen):
                for (int i = 10; i < 100; i = i + 4)
                {
                    String strValue = "";
                    if (xlWorksheet.get_Range("A" + Convert.ToString(i), "A" + Convert.ToString(i)).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range("A" + Convert.ToString(i), "A" + Convert.ToString(i)).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            intDoelen = (i - 6) / 4;
                            ((TextBox)Controls["txtLeerdoel" + Convert.ToString(intDoelen)]).Text = strValue;
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                txtSucces.Text += "Aantal leerdoelen: " + Convert.ToString(intDoelen) + "\r\n";

                //Onderwerpen:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    String strCel = "D" + Convert.ToString(intDoel * 4 + 6);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range(strCel, strCel).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            ((TextBox)Controls["txtOnderwerpen" + Convert.ToString(intDoel)]).Text = strValue;
                        }
                    }
                }

                //Weging:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    String strCel = "H" + Convert.ToString(intDoel * 4 + 6);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = (100 * xlWorksheet.get_Range(strCel, strCel).Value2).ToString();
                        if (!strValue.Equals(""))
                        {
                            ((TextBox)Controls["txtWeging" + Convert.ToString(intDoel)]).Text = strValue;
                        }
                    }
                }

                //Check totaal van weging:
                double dblTotaal = 0.0;
                for (int intTotaal = 1; intTotaal <= intDoelen; intTotaal++)
                {
                    dblTotaal += Convert.ToDouble(((TextBox)Controls["txtWeging" + Convert.ToString(intTotaal)]).Text);
                }
                if (dblTotaal == 100.0)
                {
                    txtSucces.Text += "Totaal weegfactoren: " + Convert.ToString(dblTotaal) + "\r\n";
                }
                else
                {
                    txtFail.Text += "Totaal weegfactoren ongelijk 100: " + Convert.ToString(dblTotaal) + "\r\n";
                }

                //Onthouden:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    String strCel = "I" + Convert.ToString(intDoel * 4 + 6);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range(strCel, strCel).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            if (strValue.Contains("x") || strValue.Contains("X"))
                            {
                                ((TextBox)Controls["txtO" + Convert.ToString(intDoel)]).Text = "X";
                                txtSucces.Text += "Bloom niveau Onthouden geselecteerd." + "\r\n";
                            }
                        }
                    }
                }

                //Begrijpen:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    String strCel = "J" + Convert.ToString(intDoel * 4 + 6);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range(strCel, strCel).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            if (strValue.Contains("x") || strValue.Contains("X"))
                            {
                                ((TextBox)Controls["txtB" + Convert.ToString(intDoel)]).Text = "X";
                                txtSucces.Text += "Bloom niveau Begrijpen geselecteerd." + "\r\n";
                            }
                        }
                    }
                }

                //Toepassen:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    String strCel = "K" + Convert.ToString(intDoel * 4 + 6);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range(strCel, strCel).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            if (strValue.Contains("x") || strValue.Contains("X"))
                            {
                                ((TextBox)Controls["txtT" + Convert.ToString(intDoel)]).Text = "X";
                                txtSucces.Text += "Bloom niveau Toepassen geselecteerd." + "\r\n";
                            }
                        }
                    }
                }

                //Analyseren:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    String strCel = "L" + Convert.ToString(intDoel * 4 + 6);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range(strCel, strCel).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            if (strValue.Contains("x") || strValue.Contains("X"))
                            {
                                ((TextBox)Controls["txtA" + Convert.ToString(intDoel)]).Text = "X";
                                txtSucces.Text += "Bloom niveau Analyseren geselecteerd." + "\r\n";
                            }
                        }
                    }
                }

                //Evalueren:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    String strCel = "M" + Convert.ToString(intDoel * 4 + 6);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range(strCel, strCel).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            if (strValue.Contains("x") || strValue.Contains("X"))
                            {
                                ((TextBox)Controls["txtE" + Convert.ToString(intDoel)]).Text = "X";
                                txtSucces.Text += "Bloom niveau Evalueren geselecteerd." + "\r\n";
                            }
                        }
                    }
                }

                //Creeren:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    String strCel = "N" + Convert.ToString(intDoel * 4 + 6);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range(strCel, strCel).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            if (strValue.Contains("x") || strValue.Contains("X"))
                            {
                                ((TextBox)Controls["txtC" + Convert.ToString(intDoel)]).Text = "X";
                                txtSucces.Text += "Bloom niveau Creeren geselecteerd." + "\r\n";
                            }
                        }
                    }
                }

                //BoKS:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    String strCel = "F" + Convert.ToString(intDoel * 4 + 6);
                    String strCel2 = "G" + Convert.ToString(intDoel * 4 + 6);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range(strCel, strCel).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            ((TextBox)Controls["txtBoKSa" + Convert.ToString(intDoel)]).Text = strValue;
                        }

                        using (SqlCommand cmdBoKS = new SqlCommand("SELECT * FROM tblBoKS WHERE strAfkortingToetsmatrijs = '" + strValue + "'", cnnOnderwijs))
                        {
                            using (SqlDataReader rdrBoKS = cmdBoKS.ExecuteReader())
                            {
                                if (rdrBoKS.HasRows)
                                {
                                    rdrBoKS.Read();
                                    ((TextBox)Controls["txtPKBoKSa" + Convert.ToString(intDoel)]).Text = rdrBoKS["pkId"].ToString();
                                    txtSucces.Text += "BoKS gevonden in DB: " + strValue + "\r\n";
                                }
                                else
                                {
                                    ((TextBox)Controls["txtPKBoKSa" + Convert.ToString(intDoel)]).Text = "???";
                                    txtFail.Text += "BoKS niet gevonden in DB: " + strValue + "\r\n";
                                }
                                rdrBoKS.Close();
                            }
                        }

                        //Haal inhoud van cel met BoKS-codes op:
                        strValue = xlWorksheet.get_Range(strCel2, strCel2).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            //Destilleer items uit string:
                            strValue = strValue.Replace("\r\n", " ");
                            strValue = strValue.Replace("\n", " ");
                            strValue = strValue.Replace("\r", " ");
                            strValue = strValue.Replace(";", ".");
                            strValue = strValue.Replace(":", ".");
                            strValue = strValue.Replace(",", ".");
                            String strCode = "";
                            for (int i = 0; i < strValue.Length; i++)
                            {
                                if (strValue[i] >= '0' && strValue[i] <= '9')
                                {
                                    strCode += Convert.ToString(strValue[i]);
                                }
                                else if (strValue[i] == '.')
                                {
                                    strCode += ".";
                                }
                                else
                                {
                                    if (strCode.Length > 0)
                                    {
                                        ((ListBox)Controls["lstBoKSa" + Convert.ToString(intDoel)]).Items.Add(strCode);
                                        strCode = "";
                                    }
                                }
                            }
                            if (strCode.Length > 0)
                            {
                                ((ListBox)Controls["lstBoKSa" + Convert.ToString(intDoel)]).Items.Add(strCode);
                            }
                        }
                        if (((ListBox)Controls["lstBoKSa" + Convert.ToString(intDoel)]).Items.Count > 0)
                        {
                            txtSucces.Text += "Aantal BoKS-items voor " + ((TextBox)Controls["txtBoKSa" + Convert.ToString(intDoel)]).Text + ": " + ((ListBox)Controls["lstBoKSa" + Convert.ToString(intDoel)]).Items.Count.ToString() + "\r\n";
                        }
                        else
                        {
                            txtFail.Text += "Geen BoKS-items gevonden voor " + ((TextBox)Controls["txtBoKSa" + Convert.ToString(intDoel)]).Text + "\r\n";
                        }
                    }


                    strValue = "";
                    strCel = "F" + Convert.ToString(intDoel * 4 + 8);
                    strCel2 = "G" + Convert.ToString(intDoel * 4 + 8);
                    if (xlWorksheet.get_Range(strCel, strCel).Value2 != null)
                    {
                        strValue = xlWorksheet.get_Range(strCel, strCel).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            ((TextBox)Controls["txtBoKSb" + Convert.ToString(intDoel)]).Text = strValue;
                        }

                        using (SqlCommand cmdBoKS = new SqlCommand("SELECT * FROM tblBoKS WHERE strAfkortingToetsmatrijs = '" + strValue + "'", cnnOnderwijs))
                        {
                            using (SqlDataReader rdrBoKS = cmdBoKS.ExecuteReader())
                            {
                                if (rdrBoKS.HasRows)
                                {
                                    rdrBoKS.Read();
                                    ((TextBox)Controls["txtPKBoKSb" + Convert.ToString(intDoel)]).Text = rdrBoKS["pkId"].ToString();
                                    txtSucces.Text += "BoKS gevonden in DB: " + strValue + "\r\n";
                                }
                                else
                                {
                                    ((TextBox)Controls["txtPKBoKSb" + Convert.ToString(intDoel)]).Text = "???";
                                    txtFail.Text += "BoKS niet gevonden in DB: " + strValue + "\r\n";
                                }
                                rdrBoKS.Close();
                            }
                        }

                        //Haal inhoud van cel met BoKS-codes op (tweede regel):
                        strValue = xlWorksheet.get_Range(strCel2, strCel2).Value2.ToString();
                        if (!strValue.Equals(""))
                        {
                            //Destilleer items uit string:
                            strValue = strValue.Replace("\r\n", " ");
                            strValue = strValue.Replace("\n", " ");
                            strValue = strValue.Replace("\r", " ");
                            strValue = strValue.Replace(";", ".");
                            strValue = strValue.Replace(":", ".");
                            strValue = strValue.Replace(",", ".");
                            String strCode = "";
                            for (int i = 0; i < strValue.Length; i++)
                            {
                                if (strValue[i] >= '0' && strValue[i] <= '9')
                                {
                                    strCode += Convert.ToString(strValue[i]);
                                }
                                else if (strValue[i] == '.')
                                {
                                    strCode += ".";
                                }
                                else
                                {
                                    if (strCode.Length > 0)
                                    {
                                        ((ListBox)Controls["lstBoKSb" + Convert.ToString(intDoel)]).Items.Add(strCode);
                                        strCode = "";
                                    }
                                }
                            }
                            if (strCode.Length > 0)
                            {
                                ((ListBox)Controls["lstBoKSb" + Convert.ToString(intDoel)]).Items.Add(strCode);
                            }
                        }
                        if (((ListBox)Controls["lstBoKSb" + Convert.ToString(intDoel)]).Items.Count > 0)
                        {
                            txtSucces.Text += "Aantal BoKS-items voor " + ((TextBox)Controls["txtBoKSb" + Convert.ToString(intDoel)]).Text + ": " + ((ListBox)Controls["lstBoKSb" + Convert.ToString(intDoel)]).Items.Count.ToString() + "\r\n";
                        }
                        else
                        {
                            txtFail.Text += "Geen BoKS-items gevonden voor " + ((TextBox)Controls["txtBoKSb" + Convert.ToString(intDoel)]).Text + "\r\n";
                        }
                    }
                }

                //Competenties:
                for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                {
                    String strValue = "";
                    for (int intComp = 1; intComp <= 4; intComp++)
                    {
                        String strCelComp = "O" + Convert.ToString(intDoel * 4 + intComp + 5);
                        String strCelAvT = "P" + Convert.ToString(intDoel * 4 + intComp + 5);
                        String strCelAvC = "Q" + Convert.ToString(intDoel * 4 + intComp + 5);
                        String strCelMvZ = "R" + Convert.ToString(intDoel * 4 + intComp + 5);
                        String strCelGKM = "S" + Convert.ToString(intDoel * 4 + intComp + 5);
                        bool blnCompAanwezig = false;

                        if (xlWorksheet.get_Range(strCelComp, strCelComp).Value2 != null)
                        {
                            strValue = xlWorksheet.get_Range(strCelComp, strCelComp).Value2.ToString();
                            strValue = strValue.Replace(" ", "");
                            if (!strValue.Equals(""))
                            {
                                if (strValue.Equals("Analyseren")
                                    || strValue.Equals("Ontwerpen")
                                    || strValue.Equals("Realiseren")
                                    || strValue.Equals("Beheren")
                                    || strValue.Equals("Managen")
                                    || strValue.Equals("Adviseren")
                                    || strValue.Equals("Onderzoeken")
                                    || strValue.Equals("Professionaliseren"))
                                {
                                    ((ListBox)Controls["lstComp" + Convert.ToString(intDoel)]).Items.Add(strValue);
                                    txtSucces.Text += "Competentie gevonden: " + strValue + "\r\n";
                                    blnCompAanwezig = true;
                                }
                                else
                                {
                                    ((ListBox)Controls["lstComp" + Convert.ToString(intDoel)]).Items.Add("???");
                                    txtFail.Text += "Competentie niet gevonden: " + strValue + "\r\n";
                                    blnCompAanwezig = true; //Discutabel...
                                }
                            }
                        }

                        if (blnCompAanwezig)
                        {
                            //AvT:
                            if (xlWorksheet.get_Range(strCelAvT, strCelAvT).Value2 != null)
                            {
                                strValue = xlWorksheet.get_Range(strCelAvT, strCelAvT).Value2.ToString();
                                strValue = strValue.Replace(" ", "");
                                if (strValue.Equals("0") || strValue.Equals("I") || strValue.Equals("II") || strValue.Equals("III"))
                                {
                                    ((ListBox)Controls["lstAvT" + Convert.ToString(intDoel)]).Items.Add(strValue);
                                    txtSucces.Text += "Competentieniveau AvT: " + strValue + "\r\n";
                                }
                                else
                                {
                                    ((ListBox)Controls["lstAvT" + Convert.ToString(intDoel)]).Items.Add("???");
                                    txtFail.Text += "Competentieniveau AvT onduidelijk: " + strValue + "\r\n";
                                }
                            }
                            else
                            {
                                ((ListBox)Controls["lstAvT" + Convert.ToString(intDoel)]).Items.Add("???");
                                txtFail.Text += "Competentieniveau AvT onduidelijk bij competentie: " + strValue + "\r\n";
                            }

                            //AvC:
                            if (xlWorksheet.get_Range(strCelAvC, strCelAvC).Value2 != null)
                            {
                                strValue = xlWorksheet.get_Range(strCelAvC, strCelAvC).Value2.ToString();
                                strValue = strValue.Replace(" ", "");
                                if (strValue.Equals("0") || strValue.Equals("I") || strValue.Equals("II") || strValue.Equals("III"))
                                {
                                    ((ListBox)Controls["lstAvC" + Convert.ToString(intDoel)]).Items.Add(strValue);
                                    txtSucces.Text += "Competentieniveau AvC: " + strValue + "\r\n";
                                }
                                else
                                {
                                    ((ListBox)Controls["lstAvC" + Convert.ToString(intDoel)]).Items.Add("???");
                                    txtFail.Text += "Competentieniveau AvC onduidelijk: " + strValue + "\r\n";
                                }
                            }
                            else
                            {
                                ((ListBox)Controls["lstAvC" + Convert.ToString(intDoel)]).Items.Add("???");
                                txtFail.Text += "Competentieniveau AvC onduidelijk bij competentie: " + strValue + "\r\n";
                            }

                            //MvZ:
                            if (xlWorksheet.get_Range(strCelMvZ, strCelMvZ).Value2 != null)
                            {
                                strValue = xlWorksheet.get_Range(strCelMvZ, strCelMvZ).Value2.ToString();
                                strValue = strValue.Replace(" ", "");
                                if (strValue.Equals("0") || strValue.Equals("I") || strValue.Equals("II") || strValue.Equals("III"))
                                {
                                    ((ListBox)Controls["lstMvZ" + Convert.ToString(intDoel)]).Items.Add(strValue);
                                    txtSucces.Text += "Competentieniveau MvZ: " + strValue + "\r\n";
                                }
                                else
                                {
                                    ((ListBox)Controls["lstMvZ" + Convert.ToString(intDoel)]).Items.Add("???");
                                    txtFail.Text += "Competentieniveau MvZ onduidelijk: " + strValue + "\r\n";
                                }
                            }
                            else
                            {
                                ((ListBox)Controls["lstMvZ" + Convert.ToString(intDoel)]).Items.Add("???");
                                txtFail.Text += "Competentieniveau MvZ onduidelijk bij competentie: " + strValue + "\r\n";
                            }

                            //GKM:
                            if (xlWorksheet.get_Range(strCelGKM, strCelGKM).Value2 != null)
                            {
                                strValue = xlWorksheet.get_Range(strCelGKM, strCelGKM).Value2.ToString();
                                if (!strValue.Equals(""))
                                {
                                    //String opschonen:
                                    strValue = strValue.ToLower();
                                    strValue = strValue.Replace("\r\n", "");
                                    strValue = strValue.Replace("\n", "");
                                    strValue = strValue.Replace("\r", "");
                                    strValue = strValue.Replace(";", "");
                                    strValue = strValue.Replace(":", "");
                                    strValue = strValue.Replace(",", "");
                                    strValue = strValue.Replace(" ", "");
                                    String strGKM = "";
                                    for (int i = 0; i < strValue.Length; i++)
                                    {
                                        if (strValue[i] >= 'a' && strValue[i] <= 'f')
                                        {
                                            strGKM += strValue[i];
                                            txtSucces.Text += "Gedragskenmerk toegevoegd: " + strValue[i] + "\r\n";
                                        }
                                        else
                                        {
                                            strGKM += '?';
                                            txtFail.Text += "Gedragskenmerk onduidelijk: " + strValue[i] + "\r\n";
                                        }
                                    }
                                    ((ListBox)Controls["lstGKM" + Convert.ToString(intDoel)]).Items.Add(strGKM);
                                    txtSucces.Text += "Gedragskenmerken: " + strGKM + "\r\n";
                                }
                                else
                                {
                                    ((ListBox)Controls["lstGKM" + Convert.ToString(intDoel)]).Items.Add("???");
                                    txtFail.Text += "Gedragskenmerken onduidelijk!" + "\r\n";
                                }
                            }
                            else
                            {
                                ((ListBox)Controls["lstGKM" + Convert.ToString(intDoel)]).Items.Add("???");
                                txtFail.Text += "Gedragskenmerken onduidelijk!" + "\r\n";
                            }
                        }
                    }
                }

                //Versiebeheer:

                //Zoek naar begin tabel versiebeheer:
                int intBeginVersietabel = -1;
                for (int i = 14; i < 100; i++)
                {
                    String celZoek = "A" + Convert.ToString(i);
                    if (xlWorksheet.get_Range(celZoek, celZoek).Value2 != null)
                    {
                        String strValue = xlWorksheet.get_Range(celZoek, celZoek).Value2.ToString();
                        if (strValue.Equals("Versie"))
                        {
                            intBeginVersietabel = i;
                        }
                    }
                }

                if (intBeginVersietabel > 0)
                {
                    //De versietabel is gevonden!
                    for (int i = intBeginVersietabel + 1; i < intBeginVersietabel + 10; i++)
                    {
                        String celVersie = "A" + Convert.ToString(i);
                        String strVersie;
                        if (xlWorksheet.get_Range(celVersie, celVersie).Value2 != null)
                        {
                            strVersie = xlWorksheet.get_Range(celVersie, celVersie).Value2.ToString();
                            if (!strVersie.Equals(""))
                            {
                                strVersie = strVersie.Trim();
                            }
                            else
                            {
                                if (i == 30) //Er hoort minimaal één versieregel aanwezig te zijn!
                                {
                                    lstVersies.Items.Add("???");
                                    txtFail.Text += "Versiebeheer onduidelijk! Cel -Versie- leeg?" + "\r\n";
                                    continue;
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            if (i == 30) //Er hoort minimaal één versieregel aanwezig te zijn!
                            {
                                txtFail.Text += "Versiebeheer onduidelijk! Cel -Versie- leeg?" + "\r\n";
                                continue;
                            }
                            else
                            {
                                break;
                            }
                        }

                        String celDatum = "B" + Convert.ToString(i);
                        String strDatum;
                        if (xlWorksheet.get_Range(celDatum, celDatum).Value2 != null)
                        {
                            strDatum = xlWorksheet.get_Range(celDatum, celDatum).Value2.ToString();
                            if (!strDatum.Equals(""))
                            {
                                try
                                {
                                    strDatum = DateTime.FromOADate(Convert.ToDouble(strDatum.Trim())).ToString();
                                }
                                catch (Exception ex1)
                                {
                                    try
                                    {
                                        strDatum = DateTime.Parse(strDatum.Trim()).ToString();
                                    }
                                    catch (Exception ex2)
                                    {
                                        strDatum = "iets vreemds";
                                        txtFail.Text += "Versiebeheer onduidelijk! Cel -Datum- leeg?" + "\r\n";
                                    }
                                }
                            }
                            else
                            {
                                lstVersies.Items.Add("???");
                                txtFail.Text += "Versiebeheer onduidelijk! Cel -Datum- leeg?" + "\r\n";
                                continue;
                            }
                        }
                        else
                        {
                            txtFail.Text += "Versiebeheer onduidelijk! Cel -Datum- leeg?" + "\r\n";
                            continue;
                        }

                        String celAuteur = "C" + Convert.ToString(i);
                        String strAuteur;
                        if (xlWorksheet.get_Range(celAuteur, celAuteur).Value2 != null)
                        {
                            strAuteur = xlWorksheet.get_Range(celAuteur, celAuteur).Value2.ToString();
                            if (!strAuteur.Equals(""))
                            {
                                strAuteur = strAuteur.Trim();
                            }
                            else
                            {
                                lstVersies.Items.Add("???");
                                txtFail.Text += "Versiebeheer onduidelijk! Cel -Auteur- leeg?" + "\r\n";
                                continue;
                            }
                        }
                        else
                        {
                            txtFail.Text += "Versiebeheer onduidelijk! Cel -Auteur- leeg?" + "\r\n";
                            continue;
                        }

                        String celBeschrijving = "E" + Convert.ToString(i);
                        String strBeschrijving;
                        if (xlWorksheet.get_Range(celBeschrijving, celBeschrijving).Value2 != null)
                        {
                            strBeschrijving = xlWorksheet.get_Range(celBeschrijving, celBeschrijving).Value2.ToString();
                            if (!strBeschrijving.Equals(""))
                            {
                                strBeschrijving = strBeschrijving.Trim();
                            }
                            else
                            {
                                lstVersies.Items.Add("???");
                                txtFail.Text += "Versiebeheer onduidelijk! Cel -Beschrijving- leeg?" + "\r\n";
                                continue;
                            }
                        }
                        else
                        {
                            txtFail.Text += "Versiebeheer onduidelijk! Cel -Beschrijving- leeg?" + "\r\n";
                            continue;
                        }

                        //Plak alle versiedata aan elkaar en plaats in listbox:
                        lstVersies.Items.Add(strVersie + "|" + strDatum + "|" + strAuteur + "|" + strBeschrijving);
                        txtSucces.Text += strVersie + "|" + strDatum + "|" + strAuteur + "|" + strBeschrijving + "\r\n";

                        //Importeer de versieregel in de database:
                        if (blnImport && false)
                        {
                            // Stap 1: haal pkId op uit de database:
                            int intMaxId = 0;
                            using (SqlCommand cmdMaxId = new SqlCommand("SELECT MAX(pkId) AS maxId FROM tblTMVersie", cnnOnderwijs))
                            {
                                using (SqlDataReader rdrMaxId = cmdMaxId.ExecuteReader())
                                {
                                    rdrMaxId.Read();
                                    if (!rdrMaxId.IsDBNull(0))
                                    {
                                        intMaxId = (int)rdrMaxId["maxId"];
                                    }
                                    rdrMaxId.Close();
                                }
                            }
                            intMaxId++;

                            // Stap 2: Stel Query samen:
                            using (SqlCommand cmdInsert = new SqlCommand("INSERT INTO tblTMVersie (pkId, fkToets, decNummer, datDatum, fkAuteur, strBeschrijving, strMoetWeg) " +
                                "VALUES (" +
                                intMaxId.ToString() + ", " +
                                fkToets.ToString() + ", " +
                                strVersie + ", '" +
                                strDatum + "', " +
                                "16" + ", '" +
                                strBeschrijving + "', '" +
                                strAuteur + "')", cnnOnderwijs))
                            {
                                //MessageBox.Show(cmdInsert.CommandText);
                                int intAantalRecords = cmdInsert.ExecuteNonQuery();
                            }
                        }
                    }
                }
                else
                {
                    lstVersies.Items.Add("???");
                    txtFail.Text += "Versiebeheer onduidelijk! Tabel niet gevonden" + "\r\n";
                }

                //**BEGIN** leerdoelen naar database===============================================================================:
                if (blnImport)
                {
                    for (int intDoel = 1; intDoel <= intDoelen; intDoel++)
                    {
                        // Stap 1: Haal leerdoel-data uit form:
                        String strLeerdoel = ((TextBox)Controls["txtLeerdoel" + Convert.ToString(intDoel)]).Text.ToString().Replace("'","`");
                        String strOnderwerpen = ((TextBox)Controls["txtOnderwerpen" + Convert.ToString(intDoel)]).Text.ToString().Replace("'", "`");
                        String strWeging = ((TextBox)Controls["txtWeging" + Convert.ToString(intDoel)]).Text;
                        bool blnO = ((TextBox)Controls["txtO" + Convert.ToString(intDoel)]).Text.Equals("X");
                        bool blnB = ((TextBox)Controls["txtB" + Convert.ToString(intDoel)]).Text.Equals("X");
                        bool blnT = ((TextBox)Controls["txtT" + Convert.ToString(intDoel)]).Text.Equals("X");
                        bool blnA = ((TextBox)Controls["txtA" + Convert.ToString(intDoel)]).Text.Equals("X");
                        bool blnE = ((TextBox)Controls["txtE" + Convert.ToString(intDoel)]).Text.Equals("X");
                        bool blnC = ((TextBox)Controls["txtC" + Convert.ToString(intDoel)]).Text.Equals("X");

                        // Stap 2: haal pkId op uit de database:
                        int intMaxId = 0;
                        using (SqlCommand cmdMaxId = new SqlCommand("SELECT MAX(pkId) AS maxId FROM tblDoel", cnnOnderwijs))
                        {
                            using (SqlDataReader rdrMaxId = cmdMaxId.ExecuteReader())
                            {
                                rdrMaxId.Read();
                                if (!rdrMaxId.IsDBNull(0))
                                {
                                    intMaxId = (int)rdrMaxId["maxId"];
                                }
                                rdrMaxId.Close();
                            }
                        }
                        int pkDoel = intMaxId + 1;

                        // Stap 3: Stel Query samen en voer 'm uit:
                        using (SqlCommand cmdInsert = new SqlCommand("INSERT INTO tblDoel (pkId, fkDoeltype, fkToets, strOmschrijving, strOnderwerpen, decWeging, blnOnthouden, blnBegrijpen, blnToepassen, blnAnalyseren, blnEvalueren, blnCreeren) " +
                            "VALUES (" +
                            pkDoel.ToString() + ", " +
                            "1" + ", " +
                            fkToets.ToString() + ", '" +
                            strLeerdoel + "', '" +
                            strOnderwerpen + "', " +
                            strWeging + ", " +
                            (blnO ? 1 : 0) + ", " +
                            (blnB ? 1 : 0) + ", " +
                            (blnT ? 1 : 0) + ", " +
                            (blnA ? 1 : 0) + ", " +
                            (blnE ? 1 : 0) + ", " +
                            (blnC ? 1 : 0) + ")", cnnOnderwijs))
                        {
                            //MessageBox.Show(cmdInsert.CommandText);
                            int intAantalRecords = cmdInsert.ExecuteNonQuery();
                        }

                        // Stap 4a: BoKS-data:
                        if (!((TextBox)Controls["txtPKBoKSa" + Convert.ToString(intDoel)]).Text.ToString().Equals(""))
                        {
                            int pkBoKS = Convert.ToInt32(((TextBox)Controls["txtPKBoKSa" + Convert.ToString(intDoel)]).Text);


                            // Stap 4a.1: Haal alle BoKS-items uit lijst:
                            for (int iItem = 0; iItem < ((ListBox)Controls["lstBoKSa" + Convert.ToString(intDoel)]).Items.Count; iItem++)
                            {
                                //Stap 4a.1.1: Categorie- en Itemnummer:
                                String strItem = ((ListBox)Controls["lstBoKSa" + Convert.ToString(intDoel)]).Items[iItem].ToString();
                                String[] strSplit = strItem.Split('.');
                                int intCat = Convert.ToInt32(strSplit[0]);
                                int intItm = Convert.ToInt32(strSplit[1]);

                                // Stap 4a.1.2: Haal PK van BoKS-item uit DB:
                                int pkItem = 0;
                                using (SqlCommand cmdPKItem = new SqlCommand("SELECT pkItem FROM qryBoKSItem WHERE pkBoKS = " + pkBoKS.ToString() + " AND Categorienummer = " + intCat.ToString() + " AND Itemnummer = " + intItm.ToString(), cnnOnderwijs))
                                {
                                    using (SqlDataReader rdrPKItem = cmdPKItem.ExecuteReader())
                                    {
                                        rdrPKItem.Read();
                                        if (!rdrPKItem.IsDBNull(0))
                                        {
                                            pkItem = (int)rdrPKItem["pkItem"];
                                        }
                                        rdrPKItem.Close();
                                    }
                                }

                                // Stap 4a.1.3: Lees de maximale pk-waarde uit tblDoelBoKSItem:
                                int pkDoelBoKSItem = 0;
                                using (SqlCommand cmdMaxId = new SqlCommand("SELECT MAX(pkId) AS maxId FROM tblDoelBoKSItem", cnnOnderwijs))
                                {
                                    using (SqlDataReader rdrMaxId = cmdMaxId.ExecuteReader())
                                    {
                                        rdrMaxId.Read();
                                        if (!rdrMaxId.IsDBNull(0))
                                        {
                                            pkDoelBoKSItem = (int)rdrMaxId["maxId"];
                                        }
                                        rdrMaxId.Close();
                                    }
                                }
                                pkDoelBoKSItem++;

                                // Stap 4a.1.4: Stel Query samen en voer 'm uit:
                                using (SqlCommand cmdInsert = new SqlCommand("INSERT INTO tblDoelBoKSItem (pkId, fkDoel, fkBoKSItem) " +
                                    "VALUES (" +
                                    pkDoelBoKSItem.ToString() + ", " +
                                    pkDoel.ToString() + ", " +
                                    pkItem.ToString() + ")", cnnOnderwijs))
                                {
                                    //MessageBox.Show(cmdInsert.CommandText);
                                    int intAantalRecords = cmdInsert.ExecuteNonQuery();
                                }
                            }
                        }

                        // Stap 4b: BoKS-data:
                        if (!((TextBox)Controls["txtPKBoKSb" + Convert.ToString(intDoel)]).Text.ToString().Equals(""))
                        {
                            int pkBoKS = Convert.ToInt32(((TextBox)Controls["txtPKBoKSb" + Convert.ToString(intDoel)]).Text);


                            // Stap 4b.1: Haal alle BoKS-items uit lijst:
                            for (int iItem = 0; iItem < ((ListBox)Controls["lstBoKSb" + Convert.ToString(intDoel)]).Items.Count; iItem++)
                            {
                                //Stap 4b.1.1: Categorie- en Itemnummer:
                                String strItem = ((ListBox)Controls["lstBoKSb" + Convert.ToString(intDoel)]).Items[iItem].ToString();
                                String[] strSplit = strItem.Split('.');
                                int intCat = Convert.ToInt32(strSplit[0]);
                                int intItm = Convert.ToInt32(strSplit[1]);

                                // Stap 4b.1.2: Haal PK van BoKS-item uit DB:
                                int pkItem = 0;
                                using (SqlCommand cmdPKItem = new SqlCommand("SELECT pkItem FROM qryBoKSItem WHERE pkBoKS = " + pkBoKS.ToString() + " AND Categorienummer = " + intCat.ToString() + " AND Itemnummer = " + intItm.ToString(), cnnOnderwijs))
                                {
                                    using (SqlDataReader rdrPKItem = cmdPKItem.ExecuteReader())
                                    {
                                        rdrPKItem.Read();
                                        if (!rdrPKItem.IsDBNull(0))
                                        {
                                            pkItem = (int)rdrPKItem["pkItem"];
                                        }
                                        rdrPKItem.Close();
                                    }
                                }

                                // Stap 4b.1.3: Lees de maximale pk-waarde uit tblDoelBoKSItem:
                                int pkDoelBoKSItem = 0;
                                using (SqlCommand cmdMaxId = new SqlCommand("SELECT MAX(pkId) AS maxId FROM tblDoelBoKSItem", cnnOnderwijs))
                                {
                                    using (SqlDataReader rdrMaxId = cmdMaxId.ExecuteReader())
                                    {
                                        rdrMaxId.Read();
                                        if (!rdrMaxId.IsDBNull(0))
                                        {
                                            pkDoelBoKSItem = (int)rdrMaxId["maxId"];
                                        }
                                        rdrMaxId.Close();
                                    }
                                }
                                pkDoelBoKSItem++;

                                // Stap 4b.1.4: Stel Query samen en voer 'm uit:
                                using (SqlCommand cmdInsert = new SqlCommand("INSERT INTO tblDoelBoKSItem (pkId, fkDoel, fkBoKSItem) " +
                                    "VALUES (" +
                                    pkDoelBoKSItem.ToString() + ", " +
                                    pkDoel.ToString() + ", " +
                                    pkItem.ToString() + ")", cnnOnderwijs))
                                {
                                    //MessageBox.Show(cmdInsert.CommandText);
                                    int intAantalRecords = cmdInsert.ExecuteNonQuery();
                                }
                            }
                        }

                        //Competentie-data:
                        for (int iComp = 0; iComp < ((ListBox)Controls["lstComp" + Convert.ToString(intDoel)]).Items.Count; iComp++)
                        {
                            //Haal data op uit form:
                            int fkComp = pkCompetentie(((ListBox)Controls["lstComp" + Convert.ToString(intDoel)]).Items[iComp].ToString());
                            int fkNiveauT = pkNiveau("AvT", ((ListBox)Controls["lstAvT" + Convert.ToString(intDoel)]).Items[iComp].ToString());
                            int fkNiveauC = pkNiveau("AvC", ((ListBox)Controls["lstAvC" + Convert.ToString(intDoel)]).Items[iComp].ToString());
                            int fkNiveauZ = pkNiveau("MvZ", ((ListBox)Controls["lstMvZ" + Convert.ToString(intDoel)]).Items[iComp].ToString());

                            //Lees de maximale pk-waarde uit tblDoelCompetentie:
                            int pkDoelComp = 0;
                            using (SqlCommand cmdMaxId = new SqlCommand("SELECT MAX(pkId) AS maxId FROM tblDoelCompetentie", cnnOnderwijs))
                            {
                                using (SqlDataReader rdrMaxId = cmdMaxId.ExecuteReader())
                                {
                                    rdrMaxId.Read();
                                    if (!rdrMaxId.IsDBNull(0))
                                    {
                                        pkDoelComp = (int)rdrMaxId["maxId"];
                                    }
                                    rdrMaxId.Close();
                                }
                            }
                            pkDoelComp++;

                            //Stel Query samen en voer 'm uit:
                            using (SqlCommand cmdInsert = new SqlCommand("INSERT INTO tblDoelCompetentie (pkId, fkCompetentie, fkDoel, fkNiveauT, fkNiveauC, fkNiveauZ) " +
                                "VALUES (" +
                                pkDoelComp.ToString() + ", " +
                                fkComp.ToString() + ", " +
                                pkDoel.ToString() + ", " +
                                fkNiveauT.ToString() + ", " +
                                fkNiveauC.ToString() + ", " +
                                fkNiveauZ.ToString() + ")", cnnOnderwijs))
                            {
                                //MessageBox.Show(cmdInsert.CommandText);
                                int intAantalRecords = cmdInsert.ExecuteNonQuery();
                            }

                            //Gedragskenmerken (GKM)==============================================
                            //Lees de maximale pk-waarde uit tblDoelCompetentieGedragskenmerk:
                            int pkDoelCompetentieGKM = 0;
                            using (SqlCommand cmdMaxId = new SqlCommand("SELECT MAX(pkId) AS maxId FROM tblDoelCompetentieGedragskenmerk", cnnOnderwijs))
                            {
                                using (SqlDataReader rdrMaxId = cmdMaxId.ExecuteReader())
                                {
                                    rdrMaxId.Read();
                                    if (!rdrMaxId.IsDBNull(0))
                                    {
                                        pkDoelCompetentieGKM = (int)rdrMaxId["maxId"];
                                    }
                                    rdrMaxId.Close();
                                }
                            }
                            pkDoelCompetentieGKM++;

                            //Lees GKM-letters:
                            foreach (char chrGKM in ((ListBox)Controls["lstGKM" + Convert.ToString(intDoel)]).Items[iComp].ToString().ToCharArray())
                            {
                                //Zoek in de DB naar de PK die bij het GKM hoort:
                                int pkGKM = 0;
                                using (SqlCommand cmdPkGKM = new SqlCommand("SELECT pkGedragskenmerk FROM qryGedragskenmerk WHERE pkCompetentie = " + fkComp.ToString() + " AND Gedragskenmerkindex = '" + chrGKM.ToString() + "'", cnnOnderwijs))
                                {
                                    using (SqlDataReader rdrPkGKM = cmdPkGKM.ExecuteReader())
                                    {
                                        rdrPkGKM.Read();
                                        if (!rdrPkGKM.IsDBNull(0))
                                        {
                                            pkGKM = (int)rdrPkGKM["pkGedragskenmerk"];
                                        }
                                        rdrPkGKM.Close();
                                    }
                                }

                                //Stel Query samen en voer 'm uit:
                                using (SqlCommand cmdInsert = new SqlCommand("INSERT INTO tblDoelCompetentieGedragskenmerk (pkId, fkGedragskenmerk, fkDoelCompetentie) " +
                                    "VALUES (" +
                                    pkDoelCompetentieGKM.ToString() + ", " +
                                    pkGKM.ToString() + ", " +
                                    pkDoelComp.ToString() + ")", cnnOnderwijs))
                                {
                                    //MessageBox.Show(cmdInsert.CommandText);
                                    int intAantalRecords = cmdInsert.ExecuteNonQuery();
                                }

                                //PK ophogen...
                                pkDoelCompetentieGKM++;
                            }
                        }
                    }
                }
                //**EIND** leerdoelen naar database===============================================================================:

                //Toetsvorm naar DB (ben ik wat laat achter gekomen):
                if (blnToetsvormImport)
                {
                    //Toetsvorm:
                    using (SqlCommand cmdInsert = new SqlCommand("UPDATE tblToets SET fkToetsvorm = " + txtPKToetsvorm.Text + " WHERE pkId = " + txtPKToets.Text, cnnOnderwijs))
                    {
                        //MessageBox.Show(cmdInsert.CommandText);
                        int intAantalRecords = cmdInsert.ExecuteNonQuery();
                    }

                    //Beoordelingswijze:
                    using (SqlCommand cmdInsert = new SqlCommand("UPDATE tblToets SET fkBeoordelingswijze = " + txtPKBeoordelingswijze.Text + " WHERE pkId = " + txtPKToets.Text, cnnOnderwijs))
                    {
                        //MessageBox.Show(cmdInsert.CommandText);
                        int intAantalRecords = cmdInsert.ExecuteNonQuery();
                    }
                }


                //Toetsmatrijs afsluiten...
                xlWorkbook.Close(true, misValue, misValue);
                xlApplication.Quit();

                txtSucces.Text += "Bestand gesloten: " + filePaths[t] + "\r\n";
                txtSucces.Text += "==================================================" + "\r\n";

                txtFail.Text += "Bestand gesloten: " + filePaths[t] + "\r\n";
                txtFail.Text += "==================================================" + "\r\n";

                releaseObject(xlWorksheet);
                releaseObject(xlWorkbook);
                releaseObject(xlApplication);
            }
        }

        private void frmImport_Load(object sender, EventArgs e)
        {
            txtSucces.Text += "Applicatie gestart" + "\r\n";
            cnnOnderwijs.Open();
            txtSucces.Text += "Databaseconnectie geopend" + "\r\n";
        }

        private int pkCompetentie(String strCompetentie)
        {
            switch (strCompetentie)
            {
                case "Analyseren": return 1;
                case "Ontwerpen": return 2;
                case "Realiseren": return 3;
                case "Beheren": return 4;
                case "Managen": return 5;
                case "Adviseren": return 6;
                case "Onderzoeken": return 7;
                case "Professionaliseren": return 8;
                default: return -1;
            }
        }

        private int pkNiveau(String strFactor, String strCode)
        {
            switch (strFactor)
            {
                case "AvT":
                    switch (strCode)
                    {
                        case "0":
                            return 1;
                        case "I":
                            return 4;
                        case "II":
                            return 7;
                        case "III":
                            return 10;
                        default:
                            return -1;
                    }
                case "AvC":
                    switch (strCode)
                    {
                        case "0":
                            return 2;
                        case "I":
                            return 5;
                        case "II":
                            return 8;
                        case "III":
                            return 11;
                        default:
                            return -1;
                    }
                case "MvZ":
                    switch (strCode)
                    {
                        case "0":
                            return 3;
                        case "I":
                            return 6;
                        case "II":
                            return 9;
                        case "III":
                            return 12;
                        default:
                            return -1;
                    }
                default:
                    return -1;
            }
        }
    }
}
