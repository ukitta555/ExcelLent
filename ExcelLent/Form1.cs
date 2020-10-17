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
        private string[] alphabet;
        public Form1()
        {
            // set up alphabet for column naming
            alphabet = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};

            // launch component
            InitializeComponent();

            // amount of columns and rows
            dataGridView1.ColumnCount = 150;
            dataGridView1.RowCount = 28;


            //dummy data
            this.dataGridView1.Rows.Insert(0, "one", "two", "three", "four");
            this.dataGridView1.Rows.Insert(0, "Влад", "не ", "любит", "шарп");

            // fill names
            FillColumnNames(); // start from A
            FillRowNames(); // start from 1

            // design decision - fit all columns to the data they contain
            FitColumnWidth();
            FitRowWidth();
        }

        private void FitColumnWidth()
        {

            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
        }
        private  void FitRowWidth()
        {
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
        }
        // change column names to A, B, C, D, .., AA, AB, ...
        private void FillColumnNames()
        {
            DataGridViewColumn col;
            int numberToConvert = 1;
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                col = dataGridView1.Columns[i];
                if (numberToConvert % 27 == 0)
                {
                    numberToConvert++; // degenarate case (two AA columns and so on)...
                }
                Stack<int> base27Number = ConvertToBase27(numberToConvert);
                StringBuilder name = new StringBuilder("");
                while (base27Number.Count != 0)
                {
                    int nextPartOfNumber = base27Number.Pop();
                    string character = alphabet[nextPartOfNumber - 1];
                    name.Append(character);
                }
                col.HeaderText = name.ToString();
                numberToConvert++;
            }
        }

        /*
        // from a certain column after delete / add
        private void FillColumnNames(int index)
        {
            DataGridViewColumn col;
            int numberToConvert = index + 1; // use the trick with naming - it starts from 1 in base 27
            if (numberToConvert % 28 == 0) numberToConvert++; // another degenrate case - double AA etc.
            MessageBox.Show(numberToConvert.ToString());
            // start from the new column 
            for (int i = index; i < dataGridView1.ColumnCount; i++)
            {
                col = dataGridView1.Columns[i];
                if (numberToConvert % 27 == 0)
                {
                    numberToConvert++; // degenarate case (two AA columns and so on)...
                }
                Stack<int> base27Number = ConvertToBase27(numberToConvert);
                StringBuilder name = new StringBuilder("");
                while (base27Number.Count != 0)
                {
                    int nextPartOfNumber = base27Number.Pop();
                    string character = alphabet[nextPartOfNumber - 1];
                    name.Append(character);
                }
                col.HeaderText= name.ToString();
                numberToConvert++;
            }
        }
        */
        // fill the starting table
        private void FillRowNames()
        {
            DataGridViewRow row;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                row = dataGridView1.Rows[i];
                row.HeaderCell.Value = i.ToString();
            }
        }
        /*
        // after deletion / addition
        private void FillRowNames(int index)
        {
            DataGridViewRow row;
            for (int i = index; i < dataGridView1.RowCount; i++)
            {
                row = dataGridView1.Rows[i];
                row.HeaderCell.Value = i.ToString();
            }
        }
        */
        private Stack<int> ConvertToBase27(int number)
        {
            Stack<int> base26Representation = new Stack<int>();
            while (number / 27 > 0)
            {
                base26Representation.Push(number % 27);
                number /= 27;
            }
            base26Representation.Push(number % 27);
            return base26Representation;
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
        // add new column / row
        private void AddMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rowClick == -1 && columnClick != -1)
                {
                    dataGridView1_InsertCol(columnClick + 1);
                    FillColumnNames();
                    dataGridView1.Columns[columnClick + 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                }
                if (columnClick == -1 && rowClick != -1)
                {
                    dataGridView1_InsertRow(rowClick + 1);
                    FillRowNames();
                }
            }
            catch (IndexOutOfRangeException exc)
            {
                MessageBox.Show(exc.Message);
            }
        }
        // delete  column / row
        private void DelMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (rowClick == -1 && columnClick != -1)
                {
                    dataGridView1_DelCol(columnClick);
                    FillColumnNames();
                }
                else if (columnClick == -1 && rowClick != -1)
                {
                    dataGridView1_DelRow(rowClick);
                    FillRowNames();
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
