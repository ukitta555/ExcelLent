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
        private int rowClick;
        private int columnClick;
        public Form1()
        {
            InitializeComponent();
            dataGridView1.ColumnCount = 50;
            dataGridView1.RowCount = 50;
            dataGridView1.Columns[0].Name = "Product ID";
            dataGridView1.Columns[1].Name = "Product Name";
            dataGridView1.Columns[2].Name = "Product Price";
            this.dataGridView1.Rows.Insert(0, "one", "two", "three", "four");
            this.dataGridView1.Rows.Insert(0, "Влад", "не ", "любит", "шарп");
        }

        private void dataGridView1_InsertRow(int index)
        {
            dataGridView1.Rows.Insert(index);
        }

        private void dataGridView1_InsertCol(int index)
        {
           
           DataGridViewColumn columnToInsert = new DataGridViewColumn();

           columnToInsert.CellTemplate = new DataGridViewTextBoxCell();

           dataGridView1.Columns.Insert(index, columnToInsert);
           dataGridView1.Columns[index].Name = "New Column";
           dataGridView1.Columns[index].SortMode = DataGridViewColumnSortMode.Automatic;
        }
        
        private void dataGridView1_DelRow(int index)
        {
            dataGridView1.Rows.RemoveAt(index);
        }

        private void dataGridView1_DelCol(int index)
        {
            dataGridView1.Columns.RemoveAt(index);
        }

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Ignore if a column or row header is clicked
            if (e.RowIndex == -1 || e.ColumnIndex == -1)
            {
                if (e.Button == MouseButtons.Right)
                {

                    // Get mouse position relative to the vehicles grid
                    var relativeMousePosition = dataGridView1.PointToClient(Cursor.Position);

                    rowClick = e.RowIndex;
                    columnClick = e.ColumnIndex;
                    // Show the context menu
                    contextMenuStrip1.Show(dataGridView1, relativeMousePosition);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        { 

            // fullscreen form
            WindowState = FormWindowState.Maximized;

            //fullscreen DGV
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            dataGridView1.Dock = DockStyle.Fill;
        }

        private void AddMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rowClick == -1)
                {
                    dataGridView1_InsertCol(columnClick + 1);
                }
                if (columnClick == -1)
                {
                    dataGridView1_InsertRow(rowClick + 1);

                }
            }
            catch (IndexOutOfRangeException exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void DelMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rowClick == -1 && columnClick != -1)
                {
                    dataGridView1_DelCol(columnClick);
                }
                else if (columnClick == -1 && rowClick != -1)
                {
                    dataGridView1_DelRow(rowClick);

                }
            }
            catch (IndexOutOfRangeException exc)
            {
                MessageBox.Show(exc.Message);
            }
        }


        /*
         * var p = dataGridView1.PointToClient(Cursor.Position);
            var info = dataGridView1.HitTest(p.X, p.Y);    - to define where the mouse click event happened 
         */

    }
}
