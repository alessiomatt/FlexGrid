using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using C1.Win.C1FlexGrid;
using C1.Win.C1BarCode;

namespace Sample
{
    public partial class FlexGridTest : Form
    {
        public FlexGridTest()
        {
            InitializeComponent();


            //FlexGridExtension.x_AddColumn(this.c1FlexGrid1, "ID", false);
            //FlexGridExtension.x_AddColumn(this.c1FlexGrid1, "NAME", false);
            //this.c1FlexGrid1.x_AddColumn("AA", "한글컬럼명", false, 100, TextAlignEnum.CenterBottom);

            var data = TestData(10);
            DataTable dt = new DataTable();

            
            //this._flex.x_InitializeStyle();
            //this._flex.Rows.Fixed = 1;
            

            //this._flex.AutoGenerateColumns = true;
            //this._flex.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;

            
            ////this.c1FlexGrid1.Rows[3].Style.BackColor = Color.Blue;
            //this._flex.DataSource = MockData(10);

            // set up the grid
            //_flex.Rows.Count = 5;
            //_flex.Rows.Fixed = 1;
            //_flex.Cols.Count = 4;
            //_flex.Cols.Fixed = 0;

            //_flex[0, 0] = "차종";
            //_flex[1, 0] = "품목";
            //_flex[2, 0] = "품번";
            //_flex[3, 0] = "수량";
             //populate header cells
            //_flex[0, 0] = _flex[1, 0] = "1998";
            //_flex[0, 1] = _flex[0, 2] = _flex[0, 3] = _flex[0, 4] = "1999";
            //_flex[0, 5] = _flex[1, 5] = "2000";
            //for (int i = 1; i <= 4; i++)
            //    _flex[1, i] = string.Format("Q{0} 1999", i);

            //_flex.DataSource = CreateData();

            FlexGridExtension.x_InitializeStyle(_flex);
            

            DataView dv = CreateData().DefaultView;
            FlexGridExtension.x_DataBind(_flex, dv, false);
            FlexGridExtension.x_AllowMerging(_flex, new int[] { 1,2 }, true);
            //FlexEx fx = new FlexEx();
            //_flex.DataSource = CreateData();
            
            
            //// populate data cells
            //foreach (Column col in _flex.Cols)
            //    col.DataType = typeof(string);
            //_flex[2, 0] = _flex[3, 0] = _flex[4, 0] = 100;
            //_flex[2, 1] = _flex[3, 1] = 100; _flex[4, 1] = 1;
            //_flex[2, 2] = 100; _flex[3, 2] = 9; _flex[4, 2] = 2;
            //_flex[2, 3] = 100; _flex[3, 3] = 9; _flex[4, 3] = 3;
            //_flex[2, 4] = 9; _flex[3, 4] = 9; _flex[4, 4] = 3;
            //_flex[2, 5] = 9; _flex[3, 5] = 9; _flex[4, 5] = 5;

            //_flex.AllowMerging = AllowMergingEnum.Free;

            //foreach (Column col in _flex.Cols)
            //{
            //    col.AllowMerging = true;
            //}

            //foreach (Row row in _flex.Rows)
            //{
            //    row.AllowMerging = true;
            //}

        }
        private DataTable CreateData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("VINCD", typeof(string));
            dt.Columns.Add("ITEMCD", typeof(string));
            dt.Columns.Add("PARTNO", typeof(string));
            dt.Columns.Add("QTY", typeof(int));

            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1000RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1001RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1002RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1003RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1004RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1005RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1006RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1007RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1008RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1009RY", 10 });
            dt.Rows.Add(new object[] { "MD", "1055", "82301-3X1010RY", 10 });

            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1000RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1001RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1002RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1003RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1004RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1005RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1006RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1007RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1008RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1009RY", 10 });
            dt.Rows.Add(new object[] { "DM", "1055", "82301-MD1010RY", 10 });

            return dt;
        }

        
        private object[] TestData(int count)
        {
            string[] firstNames = new string[] { "Ed", "Tommy", "Aaron", "Abe", "Jamie", "Adam", "Dave", "David", "Jay", "Nicolas", "Nige" };
            string[] lastNames = new string[] { "Spencer", "Maintz", "Conran", "Elias", "Avins", "Mishcon", "Kaneda", "Davis", "Robinson", "Ferrero", "White" };
            int[] ratings = new int[] { 1, 2, 3, 4, 5 };
            int[] salaries = new int[] { 100, 400, 900, 1500, 1000000 };

            object[] data = new object[count];
            Random rnd = new Random();

            for (int i = 0; i < count; i++)
            {
                int ratingId = rnd.Next(ratings.Length);
                int salaryId = rnd.Next(salaries.Length);
                int firstNameId = rnd.Next(firstNames.Length);
                int lastNameId = rnd.Next(lastNames.Length);

                int rating = ratings[ratingId];
                int salary = salaries[salaryId];
                string name = String.Format("{0} {1}", firstNames[firstNameId], lastNames[lastNameId]);
                string id = "rec-" + i;

                data[i] = new object[] { id, name, rating, salary };
            }

            return data;
        }



        private DataTable MockData(int count)
        {
            string[] firstNames = new string[] { "Ed", "Tommy", "Aaron", "Abe", "Jamie", "Adam", "Dave", "David", "Jay", "Nicolas", "Nige" };
            string[] lastNames = new string[] { "Spencer", "Maintz", "Conran", "Elias", "Avins", "Mishcon", "Kaneda", "Davis", "Robinson", "Ferrero", "White" };
            int[] ratings = new int[] { 1, 2, 3, 4, 5 };
            int[] salaries = new int[] { 100, 400, 900, 1500, 1000000 };

            //object[] data = new object[count];
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("NAME");
            dt.Columns.Add("RATING");
            dt.Columns.Add("SALARY");

            Random rnd = new Random();

            for (int i = 0; i < count; i++)
            {

                int ratingId = rnd.Next(ratings.Length);
                int salaryId = rnd.Next(salaries.Length);
                int firstNameId = rnd.Next(firstNames.Length);
                int lastNameId = rnd.Next(lastNames.Length);

                int rating = ratings[ratingId];
                int salary = salaries[salaryId];
                string name = String.Format("{0} {1}", firstNames[firstNameId], lastNames[lastNameId]);
                string id = "rec-" + i;
                DataRow dr = dt.NewRow();
                dr["ID"] = id;
                dr["NAME"] = name;
                dr["RATING"] = rating;
                dr["SALARY"] = salary;
                
                //data[i] = new object[] { id, name, rating, salary };
                dt.Rows.Add(dr);
            }

            return dt;
        }

        private void c1FlexGrid1_CellChanged(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            CellStyle s = this._flex.Styles["Yellow"];
            if (s == null)
            {
                s = this._flex.Styles.Add("Yellow");
                s.BackColor = Color.Yellow;
            }
            this._flex.Invalidate(e.Row, -1);
            CellRange rg = this._flex.GetCellRange(e.Row, this._flex.Cols[1].Index);
            //rg.Style = ((bool)this.c1FlexGrid1[e.Row, 1]) ? s : null;
        }
        
        class FlexEx : C1FlexGrid
        {
            private bool _doingMerge = false;
            private int _colIndex;

            public FlexEx()
            {
            }

            public override CellRange GetMergedRange(int row, int col, bool clip)
            {
                // save index of ID column to use in merging logic
                //
                _colIndex = Cols["ID"].Index;

                // set flag to use custom data when merging
                //
                _doingMerge = true;

                // call base class merging logic (will retrieve data using GetData method)
                //
                CellRange cellRange = base.GetMergedRange(row, col, clip);

                // reset flag so GetData behaves as usual
                //
                _doingMerge = false;

                // return the merged range
                return cellRange;
            }

            public override object GetData(int row, int col)
            {
                // getting data to determine merging range:
                // append content of ID column to avoid merging cells in rows with different IDs
                //
                if (_doingMerge && _colIndex > -1 && col != _colIndex)
                    return base.GetDataDisplay(row, col) + base.GetDataDisplay(row, _colIndex);

                // getting data to display, measure, edit etc.
                // let base class handle it as usual
                //
                return base.GetData(row, col);
            }
            
        }
    }
}
