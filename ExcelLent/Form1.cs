using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
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
            dataGridView1.ColumnCount = 200;
            dataGridView1.RowCount = 200;
            dataGridView1.Columns[0].Name = "Product ID";
            dataGridView1.Columns[1].Name = "Product Name";
            dataGridView1.Columns[2].Name = "Product Price";
            this.dataGridView1.Rows.Insert(0, "one", "two", "three", "four");
            this.dataGridView1.Rows.Insert(0, "Влад", "не ", "любит", "шарп");
        }

        private void dataGridView1_insertRow(int index)
        {
            dataGridView1.Rows.Insert(index);
        }

        private void dataGridView1_insertCol(int index)
        {
           
           DataGridViewColumn columnToInsert = new DataGridViewColumn();

           columnToInsert.CellTemplate = new DataGridViewTextBoxCell();

           dataGridView1.Columns.Insert(index, columnToInsert);
           dataGridView1.Columns[index].Name = "New Column";
           dataGridView1.Columns[index].SortMode = DataGridViewColumnSortMode.Automatic;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex == dataGridView1.RowCount - 1)
                {
                    dataGridView1_insertCol(e.ColumnIndex + 1);
                }
                else if (e.ColumnIndex == -1)
                {   
                    dataGridView1_insertRow(e.RowIndex + 1);
                    
                }
            }
            catch (IndexOutOfRangeException exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex == dataGridView1.RowCount - 1 && e.Button == MouseButtons.Right)
            {
                dataGridView1.Columns.RemoveAt(e.ColumnIndex );
            }
            else if (e.ColumnIndex == -1 && e.Button == MouseButtons.Right)
            {
                dataGridView1.Rows.RemoveAt(e.RowIndex);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        { 

            // fullscreen form
            TopMost = true;
            WindowState = FormWindowState.Maximized;

            //fullscreen DGV
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            dataGridView1.Dock = DockStyle.Fill;
        }


        /*
         * var p = dataGridView1.PointToClient(Cursor.Position);
            var info = dataGridView1.HitTest(p.X, p.Y);    - to define where the mouse click event happened 
         */

    }
}
