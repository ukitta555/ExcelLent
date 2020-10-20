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
        private int currentRow;
        private int currentColumn;
        public static string[] alphabet;

        public int CurrentRow
        {
            get { return currentRow; }
        }

        public int CurrentColumn
        {
            get { return currentColumn; }
        }
        public Form1()
        {
            // set up alphabet for column naming
            alphabet = new string[] {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
            // launch component
            InitializeComponent();

   

            int columnCount = 40;
            DataGridViewColumn col;
            // set MyCell as a template for all columns in DataGridView
            for (int i = 0; i < columnCount; i++)
            {
                col = new DataGridViewColumn(new MyCell());
                dataGridView1.Columns.Insert(0, col);
            }

            dataGridView1.RowCount = 40;
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
            Stack<int> base27Number;
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                base27Number = ConvertToBase27(numberToConvert);
                col = dataGridView1.Columns[i];
                while (base27Number == null) 
                {
                    numberToConvert++;
                    base27Number = ConvertToBase27(numberToConvert);
                }
                StringBuilder name = new StringBuilder("");
                while (base27Number.Count != 0)
                {
                    int nextPartOfNumber = base27Number.Pop();
                    string character = alphabet[nextPartOfNumber - 1];
                    name.Append(character);
                }
                col.Name = name.ToString();
                col.HeaderText = name.ToString();
                numberToConvert++;
            }
        }

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
        
        private Stack<int> ConvertToBase27(int number)
        {
            Stack<int> base26Representation = new Stack<int>();
            while (number / 27 > 0)
            {
                if (number % 27 == 0) return null;
                base26Representation.Push(number % 27);
                number /= 27;
            }
            if (number % 27 == 0) return null;
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

           columnToInsert.CellTemplate = new MyCell();

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

                    currentRow = e.RowIndex;
                    currentColumn = e.ColumnIndex;
                    // Show the context menu
                    contextMenuStrip1.Show(dataGridView1, relativeMousePosition);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        { 
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        }
        // add new column / row
        private void AddMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentRow == -1 && currentColumn != -1)
                {
                    dataGridView1_InsertCol(currentColumn + 1);
                    FillColumnNames();
                    dataGridView1.Columns[currentColumn + 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                    dataGridView1.Columns[currentColumn + 1].CellTemplate = new MyCell();
                }
                if (currentColumn == -1 && currentRow != -1)
                {
                    dataGridView1_InsertRow(currentRow + 1);
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
                if (currentRow == -1 && currentColumn != -1)
                {
                    dataGridView1_DelCol(currentColumn);
                    FillColumnNames();
                }
                else if (currentColumn == -1 && currentRow != -1)
                {
                    dataGridView1_DelRow(currentRow);
                    FillRowNames();
                }
            }
            catch (IndexOutOfRangeException exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CleanVariables();
            foreach (MyCell cell in dataGridView1.SelectedCells)
            {
                cell.Expression = ExpressionTextBox.Text;
                
            }
            EvaluateTable();
        }
        private void EvaluateTable ()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (MyCell cell in row.Cells)
                {
                    currentRow = cell.RowIndex;
                    currentColumn = cell.ColumnIndex;
                    cell.Value = Calculator.Evaluate(cell.Expression);
                }
            }
        }

        private void CleanVariables()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (MyCell cell in row.Cells)
                {
                    cell.Variables.Clear();
                }
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            ExpressionTextBox.Text = "";
            if (dataGridView1.SelectedCells.Count == 1)
            { 
                ExpressionTextBox.Text = ((MyCell)dataGridView1.SelectedCells[0]).Expression;
            }
        }

        public DataGridView getDataGridView()
        {
            return dataGridView1;
        }
        /*
         * var p = dataGridView1.PointToClient(Cursor.Position);
            var info = dataGridView1.HitTest(p.X, p.Y);    - to define where the mouse click event happened 
         */

    }
}
