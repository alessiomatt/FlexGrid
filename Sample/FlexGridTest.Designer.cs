namespace Sample
{
    partial class FlexGridTest
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this._flex = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.c1BarCode1 = new C1.Win.C1BarCode.C1BarCode();
            this.c1Button1 = new C1.Win.C1Input.C1Button();
            ((System.ComponentModel.ISupportInitialize)(this._flex)).BeginInit();
            this.SuspendLayout();
            // 
            // _flex
            // 
            this._flex.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this._flex.Location = new System.Drawing.Point(46, 91);
            this._flex.Name = "_flex";
            this._flex.Rows.Count = 0;
            this._flex.Rows.DefaultSize = 20;
            this._flex.Rows.Fixed = 0;
            this._flex.Size = new System.Drawing.Size(558, 398);
            this._flex.SubtotalPosition = C1.Win.C1FlexGrid.SubtotalPositionEnum.BelowData;
            this._flex.TabIndex = 0;
            this._flex.CellChanged += new C1.Win.C1FlexGrid.RowColEventHandler(this.c1FlexGrid1_CellChanged);
            // 
            // c1BarCode1
            // 
            this.c1BarCode1.Location = new System.Drawing.Point(675, 157);
            this.c1BarCode1.Name = "c1BarCode1";
            this.c1BarCode1.Size = new System.Drawing.Size(75, 108);
            this.c1BarCode1.TabIndex = 1;
            this.c1BarCode1.Text = "c1BarCode1";
            // 
            // c1Button1
            // 
            this.c1Button1.Location = new System.Drawing.Point(744, 384);
            this.c1Button1.Name = "c1Button1";
            this.c1Button1.Size = new System.Drawing.Size(75, 23);
            this.c1Button1.TabIndex = 2;
            this.c1Button1.Text = "c1Button1";
            this.c1Button1.UseVisualStyleBackColor = true;
            // 
            // FlexGridTest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(874, 566);
            this.Controls.Add(this.c1Button1);
            this.Controls.Add(this.c1BarCode1);
            this.Controls.Add(this._flex);
            this.Name = "FlexGridTest";
            this.Text = "FlexGridTest";
            ((System.ComponentModel.ISupportInitialize)(this._flex)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private C1.Win.C1FlexGrid.C1FlexGrid _flex;
        private C1.Win.C1BarCode.C1BarCode c1BarCode1;
        private C1.Win.C1Input.C1Button c1Button1;





    }
}