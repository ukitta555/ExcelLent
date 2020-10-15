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
            dataGridView1.ColumnCount = 5;
            dataGridView1.RowCount = 5;
            dataGridView1.Columns[0].Name = "Product ID";
            dataGridView1.Columns[1].Name = "Product Name";
            dataGridView1.Columns[2].Name = "Product Price";
            this.dataGridView1.Rows.Insert(0, "one", "two", "three", "four");
            this.dataGridView1.Rows.Insert(0, "Влад", "не ", "любит", "шарп");
            this.dataGridView1.Rows.Add(new string[] { "xd", "lmao" });

            


            //btn.UseColumnTextForButtonValue = true;
        }

        private void dataGridView1_insertRow(int index)
        {
            //dataGridView1.RowCount++;
            dataGridView1.Rows.Insert(index);
        }

        private void dataGridView1_insertCol(int index)
        {
            // dataGridView1.ColumnCount++;
           //DataGridViewCell tmp_cell = new DataGridViewCell();
           
           DataGridViewColumn columnToInsert = new DataGridViewColumn();

           columnToInsert.CellTemplate = new DataGridViewTextBoxCell();

           dataGridView1.Columns.Insert(index, columnToInsert);
           dataGridView1.Columns[index].Name = "New Column";
           dataGridView1.Columns[index].SortMode = DataGridViewColumnSortMode.Automatic;
           dataGridView1.Refresh();
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

        private void dataGridView1_MouseHover(object sender, EventArgs e)
        {
            var p = dataGridView1.PointToClient(Cursor.Position);
            var info = dataGridView1.HitTest(p.X, p.Y);

            if (info.RowIndex == 0)
            {
                //object value = dataGridView1.Rows[0].Cells[info.ColumnIndex].Value;
                //dataGridView1.Columns[info.ColumnIndex].CellTemplate = new DataGridViewHeader
            }
            //MessageBox.Show(dataGridView1.Columns[0].HeaderCell.GetType().ToString());
            //MessageBox.Show(dataGridView1[2, -1].GetType().ToString());
        }
    }
}
