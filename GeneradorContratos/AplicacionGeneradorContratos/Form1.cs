using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Core;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using System.Drawing.Drawing2D;

using System.Data.OleDb;

namespace AplicacionGeneradorContratos
{
    public partial class Form1 : Form
    {

        string properties;
        Boolean boolFuenteSeteado = false;
        Boolean boolDestinoSeteado = false;
        string rutaWordTemplate;

        public Form1()
        {
            rutaWordTemplate = GenerarRecursoTemporal("AplicacionGeneradorContratos", "Templates", "MODELO_CONTRATO_RENOVACION_PARAMETRIZADO.doc");
            InitializeComponent();
        }

        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref findText,
                        ref matchCase, ref matchWholeWord,
                        ref matchWildCards, ref matchSoundLike,
                        ref nmatchAllForms, ref forward,
                        ref wrap, ref format, ref replaceWithText,
                        ref replace, ref matchKashida,
                        ref matchDiactitics, ref matchAlefHamza,
                        ref matchControl);
        }

        private void CreateWordDocument(object filename, string pathExcel)
        {
            List<int> processesaftergen;
            DataSet dsExcel = new DataSet();
            setupPropertiesBD();
            using (OleDbConnection conn = new OleDbConnection(properties))
            {
                conn.Open();
                string[] columnNames = new string[] { "NOMBRE", "APELLIDO", "DNI", "[ESTADO CIVIL]", "[NIVEL DE ESTUDIO]", "TITULO", "DOMICILIO", "CP", "LOCALIDAD", "PROVINCIA", "FUNCION", "[HS DE TRABAJO]", "AREA", "[NIVEL Y GRADO (LETRA Y NUMERO)]" };
                string columns = String.Join(",", columnNames);
                using (OleDbDataAdapter da = new OleDbDataAdapter(
                    "SELECT " + columns + " FROM [MODELO$]", conn))
                {
                    da.TableMappings.Add("Table", "Modelo");
                    da.Fill(dsExcel);
                }
                conn.Close();
            }

            foreach (DataRow dr in dsExcel.Tables["Modelo"].Rows)
            {
                try
                {
                    List<int> processesbeforegen = getRunningProcesses();
                    object missing = Missing.Value;

                    Word.Application wordApp = new Word.Application();

                    Word.Document aDoc = null;

                    string nombre = dr["NOMBRE"].ToString();
                    string apellido = dr["APELLIDO"].ToString();
                    string dni = "";
                    if (dr["DNI"].ToString() != "")
                    {
                        dni = dr["DNI"].ToString().Substring(0, 2) + "." + dr["DNI"].ToString().Substring(2, 3) + "." + dr["DNI"].ToString().Substring(5, 3);
                    }
                    string estado_civil = dr["ESTADO CIVIL"].ToString();
                    string nivel_de_estudio = dr["NIVEL DE ESTUDIO"].ToString();
                    string titulo = dr["TITULO"].ToString();
                    string domicilio = dr["DOMICILIO"].ToString();
                    string cp = dr["CP"].ToString();
                    string localidad = dr["LOCALIDAD"].ToString();
                    string provincia = dr["PROVINCIA"].ToString();
                    string funcion = dr["FUNCION"].ToString();
                    string hs_de_trabajo = dr["HS DE TRABAJO"].ToString();
                    string area = dr["AREA"].ToString();
                    string nivel_y_grado = dr["NIVEL Y GRADO (LETRA Y NUMERO)"].ToString();

                    if (File.Exists((string)filename))
                    {
                        DateTime today = DateTime.Now;

                        object readOnly = false; //default
                        object isVisible = false;

                        wordApp.Visible = false;

                        aDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing, ref missing);

                        aDoc.Activate();

                        //Find and replace:
                        this.FindAndReplace(wordApp, "<APELLIDO>", apellido);
                        this.FindAndReplace(wordApp, "<NOMBRE>", nombre);
                        this.FindAndReplace(wordApp, "<DNI>", dni);
                        this.FindAndReplace(wordApp, "<ESTADO CIVIL>", estado_civil);
                        this.FindAndReplace(wordApp, "<NIVEL DE ESTUDIO>", nivel_de_estudio);
                        this.FindAndReplace(wordApp, "<TITULO>", titulo);
                        this.FindAndReplace(wordApp, "<DOMICILIO>", domicilio);
                        this.FindAndReplace(wordApp, "<LOCALIDAD>", localidad);
                        this.FindAndReplace(wordApp, "<CP>", cp);
                        this.FindAndReplace(wordApp, "<PROVINCIA>", provincia);
                        this.FindAndReplace(wordApp, "<FUNCION>", funcion);
                        this.FindAndReplace(wordApp, "<HS TRABAJO>", hs_de_trabajo);
                        this.FindAndReplace(wordApp, "<AREA>", area);
                        this.FindAndReplace(wordApp, "<NIVEL Y GRADO>", nivel_y_grado);

                    }
                    else
                    {
                        MessageBox.Show("file does not exist.");
                        processesaftergen = getRunningProcesses();
                        killProcesses(processesbeforegen, processesaftergen);
                        return;
                    }

                    aDoc.SaveAs2(tbDestino.Text + "\\" + apellido + "_" + nombre + "_CONTRATO", ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing);

                    //Close Document:
                    aDoc.Close(ref missing, ref missing, ref missing);

                    processesaftergen = getRunningProcesses();
                    killProcesses(processesbeforegen, processesaftergen);

                } 
                catch(Exception ex)
                {
                    MessageBox.Show("Error: " + ex.ToString());
                }

            }

            MessageBox.Show("Archivos creados.");
        }

        private void killProcesses(List<int> processesbeforegen, List<int> processesaftergen)
        {
            foreach (int pidafter in processesaftergen)
            {
                bool processfound = false;
                foreach (int pidbefore in processesbeforegen)
                {
                    if (pidafter == pidbefore)
                    {
                        processfound = true;
                    }
                }

                if (processfound == false)
                {
                    Process clsProcess = Process.GetProcessById(pidafter);
                    clsProcess.Kill();
                }
            }
        }

        public List<int> getRunningProcesses()
        {
            List<int> ProcessIDs = new List<int>();
            //here we're going to get a list of all running processes on
            //the computer
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (Process.GetCurrentProcess().Id == clsProcess.Id)
                    continue;
                if (clsProcess.ProcessName.Contains("WINWORD"))
                {
                    ProcessIDs.Add(clsProcess.Id);
                }
            }
            return ProcessIDs;
        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            this.Visible = false;

            PleaseWaitForm pleaseWait = new PleaseWaitForm();

            // Display form modelessly
            pleaseWait.Show();
            CreateWordDocument(rutaWordTemplate, tbPathExcel.Text);
            pleaseWait.Close();
            this.Visible = true;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnImportarExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog1.Filter = "Excel Worksheets 2003(*.xls)|*.xls|Excel Worksheets 2007(*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    tbPathExcel.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void setupPropertiesBD()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0";
            props["Data Source"] = tbPathExcel.Text;
            props["Extended Properties"] = "Excel 8.0";

            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }
            properties = sb.ToString();
        }

        private void btnSeleccionarCarpeta_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Elija una carpeta donde se guardaran los archivos generados.";

            folderBrowserDialog.ShowNewFolderButton = true;

            // Default to the My Documents folder.
            folderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop;

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    tbDestino.Text = folderBrowserDialog.SelectedPath;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void tbPathExcel_TextChanged_1(object sender, EventArgs e)
        {
            boolFuenteSeteado = !string.IsNullOrWhiteSpace(tbPathExcel.Text);
            if (boolFuenteSeteado && boolDestinoSeteado)
            {
                btnGenerar.Enabled = true;
            }
        }

        private void tbDestino_TextChanged_1(object sender, EventArgs e)
        {
            boolDestinoSeteado = !string.IsNullOrWhiteSpace(tbDestino.Text);
            if (boolFuenteSeteado && boolDestinoSeteado)
            {
                btnGenerar.Enabled = true;
            }
        }

        private static string GenerarRecursoTemporal(string nameSpace, string internalFilePath, string resourceName, string outDirectory = "")
        {
            Assembly _assembly = Assembly.GetCallingAssembly();

            string exeDirectory = Path.GetDirectoryName(_assembly.Location);

            if(outDirectory == "")
            {
                outDirectory = Path.Combine(exeDirectory, "Temp", internalFilePath);
            }

            if (!Directory.Exists(outDirectory))
            {
                DirectoryInfo di = Directory.CreateDirectory(outDirectory);
            }

            using (Stream s = _assembly.GetManifestResourceStream(nameSpace + "." + (internalFilePath == "" ? "" : internalFilePath + ".") + resourceName))
            {
                using (BinaryReader r = new BinaryReader(s))
                {
                    using (FileStream fs = new FileStream(outDirectory + "\\" + resourceName, FileMode.OpenOrCreate))
                    {
                        using (BinaryWriter w = new BinaryWriter(fs))
                        {
                            w.Write(r.ReadBytes((int)s.Length));
                        }
                    }
                }
            }
            return outDirectory + "\\" + resourceName;
        }

        private void btnDescargarExcel_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.Description = "Elija una carpeta donde se guardara el excel.";

            folderBrowserDialog.ShowNewFolderButton = true;

            // Default to the My Documents folder.
            folderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop;

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    GenerarRecursoTemporal("AplicacionGeneradorContratos", "Templates", "MODELO_EXCEL_CONTRATOS.xlsx", folderBrowserDialog.SelectedPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
            
        }

    }
}
