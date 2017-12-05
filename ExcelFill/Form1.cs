using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelFill
{

    public partial class Form1 : Form
    {
        Excel.Application eapp = new Excel.Application();
        Excel.Workbook listado;
        String listado_url;
        Excel.Worksheet listado_ws;
        Excel.Range listado_rg;
        Excel.Workbook formulario;
        String formulario_url;
        Excel.Worksheet formulario_ws;
        Excel.Range formulario_rg;
        public Form1()
        {
            InitializeComponent();
            eapp.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //open Listado
            OpenFileDialog openListadoS = new OpenFileDialog();
            openListadoS.Filter = "Listado (.xlsx)|*.xlsx";
            openListadoS.FilterIndex = 1;
            openListadoS.Multiselect = false;
            if (openListadoS.ShowDialog() == DialogResult.OK)
            {
                openListado(openListadoS.FileName);
                // closeListado();
            }

        }
        private void openListado(String fileName)
        {
            listado = eapp.Workbooks.Open(fileName);
            listado_ws = listado.Sheets[1];
            listado_rg = listado_ws.UsedRange;
        }
        private void closeListado()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(listado_rg);
            Marshal.ReleaseComObject(listado_ws);
            listado.Close();
            Marshal.ReleaseComObject(listado);
        }

        private void openFormulario(String fileName)
        {
            formulario = eapp.Workbooks.Open(Filename: fileName, ReadOnly: false, Editable: true);
            formulario_url = fileName;
            Console.WriteLine("SHEETS:" + formulario.Sheets.Count);
            Console.WriteLine("SHEETS2" + formulario.Worksheets.Count);
        }
        private void closeFormulario()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (formulario_rg != null)
            {
                Marshal.ReleaseComObject(formulario_rg);
            }
            if (formulario_ws != null)
            {
                Marshal.ReleaseComObject(formulario_ws);
            }
            formulario.Close();
            Marshal.ReleaseComObject(formulario);
        }
        ~Form1()
        {
            eapp.Quit();
            Marshal.ReleaseComObject(eapp);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFormularioS = new OpenFileDialog();
            openFormularioS.Filter = "Formulario (.xlsx)|*.xlsx";
            openFormularioS.FilterIndex = 1;
            openFormularioS.Multiselect = false;
            if (openFormularioS.ShowDialog() == DialogResult.OK)
            {
                openFormulario(openFormularioS.FileName);
                //closeFormulario();

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //formulario_ws = formulario.Worksheets.get_Item(1);
            //Console.WriteLine(formulario_ws.Name);
            //formulario_ws.Cells.set_Item(3, 3, "JAVIER CIFUENTES");
            //formulario_ws.Cells.set_Item(4, 3, 2048511650101);

            // return;
            int _formulario = 0;
            int rowCount = listado_rg.Rows.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                var codigo_muni = listado_rg.Cells[i, 5] as Excel.Range;
                try
                {
                    if (codigo_muni.Value != null)
                    {
                        if (codigo_muni.Value.ToString() == "713")
                        {
                            String nombre = listado_rg.Cells[i, 13].Value.ToString();
                            String dpi = listado_rg.Cells[i, 14].Value.ToString();
                            String sexo = listado_rg.Cells[i, 15].Value.ToString();
                            String comunidad = listado_rg.Cells[i, 8].Value.ToString();
                            String municipio = listado_rg.Cells[i, 6].Value.ToString();
                            String departamento = listado_rg.Cells[i, 4].Value.ToString();
                            for (int _sheet = 1; _sheet <= formulario.Sheets.Count; _sheet++)
                            {
                                llenarFormulario(_sheet, _formulario, departamento, municipio, comunidad, dpi, nombre, sexo);

                            }
                            _formulario++;
                        }
                    }
                }
                catch (Exception exc)
                {
                    Console.WriteLine("EX:");
                }
            }
            formulario.Save();
            //llenar nombre
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (formulario_rg != null)
            {
                Marshal.ReleaseComObject(formulario_rg);
            }
            if (formulario_ws != null)
            {
                Marshal.ReleaseComObject(formulario_ws);
            }
        }

        private void llenarFormulario(int hoja, int _formulario, string departamento, string municipio, string comunidad, string dpi, string nombre, string sexo)

        {
            int x = _formulario * 6 + 1;
            formulario_ws = formulario.Worksheets.get_Item(hoja);
            Console.WriteLine(formulario_ws.Name);
            formulario_ws.Cells.set_Item(3, x + 2, nombre);
            formulario_ws.Cells.set_Item(4, x + 2, "'" + dpi);
            formulario_ws.Cells.set_Item(11, x + 3, comunidad);
            formulario_ws.Cells.set_Item(12, x + 3, municipio);
            formulario_ws.Cells.set_Item(13, x + 3, departamento);
            if (sexo.ToUpper() == "MUJER")
            {
                formulario_ws.Cells.set_Item(8, x + 4, 1);
            }
            else
            {
                formulario_ws.Cells.set_Item(7, x + 4, 1);
            }
            //_nombre = nombre;
            //_dpi = dpi;
            //_sexo = 1;
            //_comunidad.Value = comunidad;
            //_municipio.Value = municipio;
            //_departamento.Value = departamento;

        }

        private void button4_Click(object sender, EventArgs e)
        {
            var comunidad = textBox1.Text.Trim();
            for (int i = 0; i < 97; i++)
            {
                Console.WriteLine("------- FORMULARIO: " + i + " --------");
                int x = i * 6 + 1;
                for (int _sheet = 1; _sheet <= formulario.Sheets.Count - 1; _sheet++)
                {
                    formulario_ws = formulario.Worksheets.get_Item(_sheet);
                    String comunidadLibro = formulario_ws.Cells[11, x + 3].Value.ToString().Trim();
                    if (comunidad.ToUpper() != comunidadLibro.ToUpper())
                    {
                        //Borrar datos
                        formulario_ws.Cells.set_Item(3, x + 2, null);
                        formulario_ws.Cells.set_Item(4, x + 2, null);
                        formulario_ws.Cells.set_Item(11, x + 3, null);
                        formulario_ws.Cells.set_Item(12, x + 3, null);
                        formulario_ws.Cells.set_Item(13, x + 3, null);
                        formulario_ws.Cells.set_Item(8, x + 4, null);
                        formulario_ws.Cells.set_Item(7, x + 4, null);

                        for (var ii = 0; ii < 20; ii++)
                        {
                            try
                            {
                                formulario_ws.Cells.set_Item(17 + ii, x + 2, null);
                            }
                            catch (Exception ex)
                            {
                                //Console.WriteLine
                            }
                        }
                    }
                }
            }
            formulario.Save();
            //llenar nombre
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (formulario_rg != null)
            {
                Marshal.ReleaseComObject(formulario_rg);
            }
            if (formulario_ws != null)
            {
                Marshal.ReleaseComObject(formulario_ws);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var _form = 0;
            var _hadValue = false;
            for (int i = 0; i < 97; i++)
            {
                Console.WriteLine("------- FORMULARIO: " + i + " --------");
                int x = i * 6 + 1;
                int xx = _form * 6 + 1;
                for (int _sheet = 1; _sheet <= formulario.Sheets.Count; _sheet++)
                {
                    formulario_ws = formulario.Worksheets.get_Item(_sheet);
                    if (formulario_ws.Cells[11, x + 3].Value != null)
                    {
                        _hadValue = true;
                        Console.WriteLine("Tiene valor");
                        //leer la info
                        var nombre = formulario_ws.Cells[3, x + 2].Value.ToString().Trim();
                        var dpi = formulario_ws.Cells[4, x + 2].Value.ToString().Trim();
                        var comunidad = formulario_ws.Cells[11, x + 3].Value.ToString().Trim();
                        var municipio = formulario_ws.Cells[12, x + 3].Value.ToString().Trim();
                        var departamento = formulario_ws.Cells[13, x + 3].Value.ToString().Trim();
                        var esMujer = true;
                        if (formulario_ws.Cells[7, x + 4].Value != null)
                        {
                            esMujer = false;
                        }
                        //mover la info
                        formulario_ws.Cells[3, xx + 2].Value = nombre;
                        formulario_ws.Cells[4, xx + 2].Value = "'" + dpi;
                        formulario_ws.Cells[11, xx + 3].Value =  comunidad;
                        formulario_ws.Cells[12, xx + 3].Value = municipio;
                        formulario_ws.Cells[13, xx + 3].Value = departamento;
                        if (!esMujer)
                        {
                            formulario_ws.Cells[7, xx + 4].Value = 1;
                        }
                        else { 
                            formulario_ws.Cells[8, xx + 4].Value = 1;
                        }
                        //Borrar datos
                        formulario_ws.Cells.set_Item(3, x + 2, null);
                        formulario_ws.Cells.set_Item(4, x + 2, null);
                        formulario_ws.Cells.set_Item(11, x + 3, null);
                        formulario_ws.Cells.set_Item(12, x + 3, null);
                        formulario_ws.Cells.set_Item(13, x + 3, null);
                        formulario_ws.Cells.set_Item(8, x + 4, null);
                        formulario_ws.Cells.set_Item(7, x + 4, null);
                        //campos extras
                        for (var ii = 0; ii < 20; ii++)
                        {
                            try
                            {
                                //leer
                                var valor = formulario_ws.Cells[17 + ii, x + 2].Value;
                                //mover
                                formulario_ws.Cells[17 + ii, xx + 2].Value = valor;
                                //borrar
                                formulario_ws.Cells.set_Item(17 + ii, x + 2, null);
                            }
                            catch (Exception ex)
                            {
                                //Console.WriteLine
                            }
                        }
                        
                    }
                }
                if (_hadValue)
                {
                    _form = _form + 1;
                    _hadValue = false;
                }

            }
            formulario.Save();
            //llenar nombre
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (formulario_rg != null)
            {
                Marshal.ReleaseComObject(formulario_rg);
            }
            if (formulario_ws != null)
            {
                Marshal.ReleaseComObject(formulario_ws);
            }
        }
    }
}
