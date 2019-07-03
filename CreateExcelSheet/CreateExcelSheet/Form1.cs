using CreateExcelSheet.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CreateExcelSheet
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void extract_Click(object sender, EventArgs e)
        {

            if (Excel.CreateFile())
                {
                    MessageBox.Show("File Created", "Success");
                }
                else
                {
                    MessageBox.Show("File not Created", "Error");
                }
            
            
        }
    }
}
