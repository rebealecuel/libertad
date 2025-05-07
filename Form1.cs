using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace libertad
{
    public partial class Form1 : Form
    {
        acciones acc = new acciones();
        public Form1()
        {
            InitializeComponent();
        }

        private void btnMostrar_Click(object sender, EventArgs e)
        {
            dgTabla.DataSource = acc.Mostrar();
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            if (acc.ExportarExcel())
                MessageBox.Show("Exportado con exito...");

            else
                MessageBox.Show("Fallo catastrofico, no se pudo exportar...");
        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
            if (acc.ImportarExcel())
                MessageBox.Show("Importado con exito...");

            else
                MessageBox.Show("Fallo catastrofico, no se pudo importar...");
        }
    }
}
