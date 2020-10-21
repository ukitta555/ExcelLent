using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.Json;

namespace ExcelLent
{
    public partial class Form1 : Form
    {
        private int currentRow;
        private int currentColumn;
        private static string dirParameter = AppDomain.CurrentDomain.BaseDirectory + @"\file.txt";
        public static string[] alphabet;
        public Dictionary<string, int> colNameToColIndex;
        
        public int CurrentRow
        {
            get { return currentRow; }
            set { currentRow = value; }
        }

        public int CurrentColumn
        {
            get { return currentColumn; }
            set { currentColumn = value; }
        }
        public Form1()
        {
            colNameToColIndex = new Dictionary<string, int>();
            // set up alphabet for column naming
            alphabet = new string[] {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
            // launch component
            InitializeComponent();



            int columnCount = 4;
            DataGridViewColumn col;
            // set MyCell as a template for all columns in DataGridView
            for (int i = 0; i < columnCount; i++)
            {
                col = new DataGridViewColumn(new MyCell());
                dataGridView1.Columns.Insert(0, col);
            }

            dataGridView1.RowCount = 4;
            // fill names
            FillColumnNames(); // start from A
            FillRowNames(); // start from 1

            // design decision - fit all columns to the data they contain
            FitColumnWidth();
            FitRowWidth();
        }

        // DataGridView event handlers
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            ExpressionTextBox.Text = "";
            if (dataGridView1.SelectedCells.Count == 1)
            {
                ExpressionTextBox.Text = ((MyCell)dataGridView1.SelectedCells[0]).Expression;
            }
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

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            MyCell editedCell = (MyCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            editedCell.Expression = editedCell.Value.ToString();
            editedCell.Value = "";
            ReevaluateTable();
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
        // Form event handlers
        private void Form1_Load(object sender, EventArgs e)
        { 
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Left;
        }

        // Context menu event handlers

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
                    ReevaluateTable();
                }
                if (currentColumn == -1 && currentRow != -1)
                {
                    dataGridView1_InsertRow(currentRow + 1);
                    FillRowNames();
                    ReevaluateTable();
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
                    ReevaluateTable();
                }
                else if (currentColumn == -1 && currentRow != -1)
                {
                    dataGridView1_DelRow(currentRow);
                    FillRowNames();
                    ReevaluateTable();
                }
            }
            catch (IndexOutOfRangeException exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        // Evaluation button event handlers

        private void button1_Click(object sender, EventArgs e)
        {
            ReevaluateTable();
        }
        
        

        
        // Top Menu event handlers.
        private void SaveMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                List<MyCell> cells = new List<MyCell>();
                string json = "";
                json += "\"" + dataGridView1.RowCount + "\"," + "\"" + dataGridView1.ColumnCount + "\",";
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    foreach (MyCell cell in row.Cells)
                    {
                        var settings = new JsonSerializerOptions { MaxDepth = 100 };
                        if (cell.Expression.Contains("\""))
                        {
                            throw new Exception("\" character is forbidden.");
                        }
                        json += JsonSerializer.Serialize(cell.Expression, settings) + ",";
                    }
                }
                json += "\"";
                File.WriteAllText(dirParameter, json);
            }
            catch (Exception exc)
            {
                MessageBox.Show("Error!", exc.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenMenuItem_Click(object sender, EventArgs e)
        {
            string[] fileContents = File.ReadAllLines(dirParameter);
            string content = "";
            foreach (string line in fileContents)
            {
                content += line;
            }
            string[] expressions = content.Split("\",\"");
            MessageBox.Show(expressions[0].Substring(1));
            MessageBox.Show(expressions[1]);
            dataGridView1.RowCount = int.Parse(expressions[0].Substring(1));
            dataGridView1.ColumnCount = int.Parse(expressions[1]);
            int i = 2;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (MyCell cell in row.Cells)
                {

                    cell.Selected = false;
                    if (expressions[i] == "")
                    {
                        cell.Expression = "0";
                    }
                    else
                    {
                        cell.Expression = expressions[i];
                    }
                    i++;
                }
            }
            ReevaluateTable();
        }

        

        // Get DataGridView to access its data (designer made it private, no other way to work around)
        public DataGridView getDataGridView()
        {
            return dataGridView1;
        }

        // Helper methods.

        // reevaluate table
        private void ReevaluateTable()
        {
            CleanVariables();
            foreach (MyCell cell in dataGridView1.SelectedCells)
            {
                cell.Expression = ExpressionTextBox.Text;

            }
            EvaluateTable();
        }

        // evaluate all expressions in cells.
        private void EvaluateTable()
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

        // clear variable references
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
        // Fit widths of rows / cols
        private void FitColumnWidth()
        {
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            }
        }
        private void FitRowWidth()
        {
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
        }
        // change column names to A, B, C, D, .., AA, AB, ...
        private void FillColumnNames()
        {
            DataGridViewColumn col;
            colNameToColIndex.Clear();
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
                colNameToColIndex.Add(name.ToString(), i);
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

        // Convert number to base 27.
        private Stack<int> ConvertToBase27(int number)
        {
            const int BASE = 27;
            Stack<int> base27Representation = new Stack<int>();
            while (number / BASE > 0)
            {
                if (number % BASE == 0) return null;
                base27Representation.Push(number % BASE);
                number /= BASE;
            }
            if (number % BASE == 0) return null;
            base27Representation.Push(number % BASE);
            return base27Representation;
        }

        /*
* var p = dataGridView1.PointToClient(Cursor.Position);
   var info = dataGridView1.HitTest(p.X, p.Y);    - to define where the mouse click event happened 
*/

    }
}
