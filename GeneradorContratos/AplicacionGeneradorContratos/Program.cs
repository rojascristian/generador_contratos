using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Text;

namespace AplicacionGeneradorContratos
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                string rutaAPlantillaWord = GenerarRecursoTemporal("AplicacionGeneradorContratos", "Templates", "MODELO_CONTRATO_RENOVACION_PARAMETRIZADO.doc");

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.ToString());
            }
        }

        private static string GenerarRecursoTemporal(string nameSpace, string internalFilePath, string resourceName)
        {
            Assembly _assembly = Assembly.GetCallingAssembly();

            string exeDirectory = Path.GetDirectoryName(_assembly.Location);

            string outDirectory = Path.Combine(exeDirectory, "Temp", internalFilePath);

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

    }
}
