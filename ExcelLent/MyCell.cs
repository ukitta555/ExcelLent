using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace ExcelLent
{
    class MyCell : DataGridViewTextBoxCell
    {
        private string expression = "";

        public string Expression
        {
            get { return expression; }
            set { expression = value; }
        }

        public override object Clone()
        {
            var objToReturn = (MyCell)base.Clone();
            objToReturn.expression = expression;
            return objToReturn;
        }
    }
}
