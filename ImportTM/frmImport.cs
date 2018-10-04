using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
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

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            //Selecteer bestand:
            if (openTM.FileName.Equals("<leeg>"))
            {
                openTM.FileName = "Toetsmatrijs";
            }
            openTM.ShowDialog();
            txtFilenameTM.Text = openTM.FileName;

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
            intDoelen = 0;
            for (int i = 1; i <= 8; i++)
            {
                ((TextBox)Controls["txtLeerdoel" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtOnderwerpen" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtWeging" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtO" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtB" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtT" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtA" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtE" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtC" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtBoKSa" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtBoKSb" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtPKBoKSa" + Convert.ToString(i)]).Text = "";
                ((TextBox)Controls["txtPKBoKSb" + Convert.ToString(i)]).Text = "";
                ((ListBox)Controls["lstBoKSa" + Convert.ToString(i)]).Items.Clear();
                ((ListBox)Controls["lstBoKSb" + Convert.ToString(i)]).Items.Clear();
            }

            //Open bestand:
            xlApplication = new Excel.Application();
            xlWorkbook = xlApplication.Workbooks.Open(openTM.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

            //Toetscode:
            txtToetscode.Text = xlWorksheet.get_Range("B3", "B3").Value2.ToString();
            using (SqlCommand cmdToets = new SqlCommand("SELECT * FROM tblToets WHERE strCode = '" + txtToetscode.Text + "'", cnnOnderwijs))
            {
                using (SqlDataReader rdrToets = cmdToets.ExecuteReader())
                {
                    if (rdrToets.HasRows)
                    {
                        rdrToets.Read();
                        txtPKToets.Text = rdrToets["pkId"].ToString();
                    }
                    else
                    {
                        txtPKToets.Text = "???";
                    }
                    rdrToets.Close();
                }
            }

            //Toetsvorm:
            txtToetsvorm.Text = xlWorksheet.get_Range("B6", "B6").Value2.ToString();
            using (SqlCommand cmdToetsvorm = new SqlCommand("SELECT * FROM tblToetsvorm WHERE strVolledig = '" + txtToetsvorm.Text + "'", cnnOnderwijs))
            {
                using (SqlDataReader rdrToetsvorm = cmdToetsvorm.ExecuteReader())
                {
                    if (rdrToetsvorm.HasRows)
                    {
                        rdrToetsvorm.Read();
                        txtPKToetsvorm.Text = rdrToetsvorm["pkId"].ToString();
                    }
                    else
                    {
                        txtPKToetsvorm.Text = "???";
                    }
                    rdrToetsvorm.Close();
                }
            }

            //Beoordelingswijze:
            txtBeoordelingswijze.Text = xlWorksheet.get_Range("B7", "B7").Value2.ToString();
            using (SqlCommand cmdBeoordelingswijze = new SqlCommand("SELECT * FROM tblBeoordelingswijze WHERE strNaam = '" + txtBeoordelingswijze.Text + "'", cnnOnderwijs))
            {
                using (SqlDataReader rdrBeoordelingswijze = cmdBeoordelingswijze.ExecuteReader())
                {
                    if (rdrBeoordelingswijze.HasRows)
                    {
                        rdrBeoordelingswijze.Read();
                        txtPKBeoordelingswijze.Text = rdrBeoordelingswijze["pkId"].ToString();
                    }
                    else
                    {
                        txtPKBeoordelingswijze.Text = "???";
                    }
                    rdrBeoordelingswijze.Close();
                }
            }

            //Leerdoelen (meteen ook het aantal doelen bepalen):
            for (int i = 10; i < 100; i=i+4)
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

            //Onderwerpen:
            for (int intDoel = 1; intDoel<=intDoelen; intDoel++)
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
                    strValue = (100*xlWorksheet.get_Range(strCel, strCel).Value2).ToString();
                    if (!strValue.Equals(""))
                    {
                        ((TextBox)Controls["txtWeging" + Convert.ToString(intDoel)]).Text = strValue;
                    }
                }
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
                        ((TextBox)Controls["txtO" + Convert.ToString(intDoel)]).Text = "X";
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
                        ((TextBox)Controls["txtB" + Convert.ToString(intDoel)]).Text = "X";
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
                        ((TextBox)Controls["txtT" + Convert.ToString(intDoel)]).Text = "X";
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
                        ((TextBox)Controls["txtA" + Convert.ToString(intDoel)]).Text = "X";
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
                        ((TextBox)Controls["txtE" + Convert.ToString(intDoel)]).Text = "X";
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
                        ((TextBox)Controls["txtC" + Convert.ToString(intDoel)]).Text = "X";
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
                            }
                            else
                            {
                                ((TextBox)Controls["txtPKBoKSa" + Convert.ToString(intDoel)]).Text = "???";
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
                        for (int i=0; i<strValue.Length; i++)
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
                            }
                            else
                            {
                                ((TextBox)Controls["txtPKBoKSb" + Convert.ToString(intDoel)]).Text = "???";
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

                }
            }





        }

        private void btnCloseTM_Click(object sender, EventArgs e)
        {
            xlWorkbook.Close(true, misValue, misValue);
            xlApplication.Quit();

            releaseObject(xlWorksheet);
            releaseObject(xlWorkbook);
            releaseObject(xlApplication);
        }

        private void frmImport_Load(object sender, EventArgs e)
        {
            cnnOnderwijs.Open();
        }
    }
}
