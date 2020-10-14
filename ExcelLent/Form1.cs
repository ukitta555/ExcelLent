using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelLent
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView1.ColumnCount = 4;
            dataGridView1.RowCount = 4;
            dataGridView1.Columns[0].Name = "Product ID";
            dataGridView1.Columns[1].Name = "Product Name";
            dataGridView1.Columns[2].Name = "Product Price";
            this.dataGridView1.Rows.Insert(0, "one", "two", "three", "four");
            this.dataGridView1.Rows.Insert(0, "Влад", "не ", "любит", "шарп");
            this.dataGridView1.Rows.Add(new string[] { "xd", "lmao" });
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            Console.WriteLine("xd");
        }
    }
}
