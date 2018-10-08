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
            //Initialiseren controls:
            txtSucces.Text = "";
            txtFail.Text = "";
            lstToetscodes.Items.Clear();

            //Selecteer een bestand zodat de map met toetsmatrijzen bekend is:Pick/Wick#12
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
    }
}
