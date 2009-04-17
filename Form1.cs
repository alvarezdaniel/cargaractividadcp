using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace CargarActividadCP
{
    public partial class Form1 : Form
    {
        // Get a handle to an application window.
        [DllImport("USER32.DLL")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ProcesaClipboard();
        }

        private void ProcesaClipboard()
        {
            if (Clipboard.ContainsText())
            {
                string s = Clipboard.GetText(TextDataFormat.Text).Replace(Environment.NewLine, string.Empty);
                //textBox1.Text = s;

                string[] array = s.Split("\t".ToCharArray());
                /*
                foreach (string a in array)
                {
                    textBox1.AppendText(a + Environment.NewLine);
                }
                */

                if (array.Length <= 1)
                {
                    MessageBox.Show("Invalid clipboard text");
                    return;
                }

                const int COLS = 11;

                int cant = array.Length / COLS;
                textBox1.AppendText("Actividades: " + cant);
                textBox1.AppendText(Environment.NewLine);
                textBox1.AppendText(Environment.NewLine);

                const int K_TAREA = 6;
                const int K_TIPO = 7;
                const int K_TIEMPO = 3;
                const int K_FECHA = 0;
                const int K_DESC = 9;
                const int K_CARGADO = 5;
                
                for (int act = 0; act <= cant-1; act++)
                {
                    textBox1.AppendText(Environment.NewLine);
                    string recurso = "DAE";
                    textBox1.AppendText("Recurso: " + recurso);
                    textBox1.AppendText(Environment.NewLine);

                    string tarea = array[K_TAREA + (act * COLS)];
                    textBox1.AppendText("Tarea: " + tarea);
                    textBox1.AppendText(Environment.NewLine);

                    string tipoActividad = array[K_TIPO + (act * COLS)];
                    textBox1.AppendText("Tipo: " + tipoActividad);
                    textBox1.AppendText(Environment.NewLine);

                    string[] h = array[K_TIEMPO + (act * COLS)].Split(":".ToCharArray());
                    string horas = h[0];
                    string minutos = h[1];
                    textBox1.AppendText("Horas: " + horas);
                    textBox1.AppendText(Environment.NewLine);
                    textBox1.AppendText("Minutos: " + minutos);
                    textBox1.AppendText(Environment.NewLine);

                    string[] f = array[K_FECHA + (act * COLS)].Split("/".ToCharArray());
                    string dia = f[0];
                    string mes = f[1];
                    string anio = f[2];
                    textBox1.AppendText("Fecha: " + dia + "/" + mes + "/" + anio);
                    textBox1.AppendText(Environment.NewLine);

                    string descripcion = array[K_DESC + (act * COLS)];
                    textBox1.AppendText("Desc: " + descripcion);
                    textBox1.AppendText(Environment.NewLine);

                    string cargado = array[K_CARGADO + (act * COLS)];
                    textBox1.AppendText("Cargado en CP: " + cargado);
                    textBox1.AppendText(Environment.NewLine);

                    if ((checkBox1.Checked) && (cargado == "N"))
                        CargarActividad(recurso, tarea, tipoActividad, horas, minutos, descripcion, dia, mes, anio);
                }
            }
        }

        private void CargarActividad(string recurso, string tarea, string tipoActividad, string horas, string minutos, string descripcion, string dia, string mes, string anio)
        {
            // http://www.autoitscript.com/autoit3/docs/appendix/SendKeys.htm
            
            // Get a handle to the Calculator application. The window class
            // and window name were obtained using the Spy++ tool.
            IntPtr cpHandle = FindWindow("TFrmModuleMain", null);
            //IntPtr cpHandle = FindWindow("TFrmABMActividad", null);

            // Verify that Calculator is a running process.
            if (cpHandle == IntPtr.Zero)
            {
                MessageBox.Show("CP is not running.");
                return;
            }

            // Da foco a CP
            SetForegroundWindow(cpHandle);
            System.Threading.Thread.Sleep(1000);

            // Abre una nueva actividad
            SendKeys.SendWait("^N");

            System.Threading.Thread.Sleep(1000);

            SendKeys.SendWait(recurso);
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(100);
            
            SendKeys.SendWait(tarea);
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(100);
            
            SendKeys.SendWait(tipoActividad);
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(100);
            
            SendKeys.SendWait(horas);
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(100);

            SendKeys.SendWait(minutos);
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(100);

            // Primero carga el mes, porque sino, por ejemplo, no se puede cargar 30/01 en febrero
            #region old
            //if (dia.Length == 1)
            //    dia = "0" + dia;
            //SendKeys.SendWait(dia);
            //System.Threading.Thread.Sleep(100);
            //SendKeys.SendWait("{RIGHT}");
            //System.Threading.Thread.Sleep(100);

            //if (mes.Length == 1)
            //    mes = "0" + mes;
            //SendKeys.SendWait(mes);
            //System.Threading.Thread.Sleep(100);
            //SendKeys.SendWait("{RIGHT}");
            //System.Threading.Thread.Sleep(100);
            #endregion

            // Primero pone día 01
            string dia1 = "01";
            SendKeys.SendWait(dia1);
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{RIGHT}");
            System.Threading.Thread.Sleep(100);

            // Pone el mes
            if (mes.Length == 1)
                mes = "0" + mes;
            SendKeys.SendWait(mes);
            System.Threading.Thread.Sleep(100);

            // Vuelve atrás y pone el día
            SendKeys.SendWait("{LEFT}");
            System.Threading.Thread.Sleep(100);
            
            if (dia.Length == 1)
                dia = "0" + dia;
            SendKeys.SendWait(dia);
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{RIGHT}");
            System.Threading.Thread.Sleep(100);

            SendKeys.SendWait("{RIGHT}");
            System.Threading.Thread.Sleep(100);

            // Pone el año
            SendKeys.SendWait(anio);
            System.Threading.Thread.Sleep(100);

            // Va para atrás con Shift-TAB 11 veces para llegar al campo descripción
            for (int i = 1; i<=11; i++)
            {
                SendKeys.SendWait("+{TAB}");
                System.Threading.Thread.Sleep(50);
            }

            //SendKeys.SendWait("{TAB}");
            //System.Threading.Thread.Sleep(100);
            
            // Agrego 2 TABS para pasar de largo los campos de horas restantes
            //SendKeys.SendWait("{TAB}");
            //System.Threading.Thread.Sleep(100);
            //SendKeys.SendWait("{TAB}");
            //System.Threading.Thread.Sleep(100);

            SendKeys.SendWait(descripcion);
            System.Threading.Thread.Sleep(100);
            SendKeys.SendWait("{TAB}");

            System.Threading.Thread.Sleep(1000);
            SendKeys.SendWait("%A");
            System.Threading.Thread.Sleep(200);
            SendKeys.SendWait("G");
        }
    }
}
