namespace usersMailboxAccess
{
    partial class frmHome
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            label1 = new System.Windows.Forms.Label();
            txtSearch = new System.Windows.Forms.TextBox();
            dataGridView1 = new System.Windows.Forms.DataGridView();
            btnGetAccess = new System.Windows.Forms.Button();
            btnClose = new System.Windows.Forms.Button();
            btnSearch = new System.Windows.Forms.Button();
            label2 = new System.Windows.Forms.Label();
            cmbSearchCat = new System.Windows.Forms.ComboBox();
            btnReset = new System.Windows.Forms.Button();
            chkMatchCase = new System.Windows.Forms.CheckBox();
            btnExport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            label1.Location = new System.Drawing.Point(252, 9);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(286, 20);
            label1.TabIndex = 0;
            label1.Text = "User Mailbox Access List Software";
            // 
            // txtSearch
            // 
            txtSearch.Location = new System.Drawing.Point(193, 47);
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new System.Drawing.Size(308, 23);
            txtSearch.TabIndex = 2;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new System.Drawing.Point(13, 82);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new System.Drawing.Size(737, 529);
            dataGridView1.TabIndex = 3;
            // 
            // btnGetAccess
            // 
            btnGetAccess.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            btnGetAccess.Location = new System.Drawing.Point(13, 617);
            btnGetAccess.Name = "btnGetAccess";
            btnGetAccess.Size = new System.Drawing.Size(111, 23);
            btnGetAccess.TabIndex = 4;
            btnGetAccess.Text = "Retrieve Data";
            btnGetAccess.UseVisualStyleBackColor = true;
            btnGetAccess.Click += btnGetAccess_Click;
            // 
            // btnClose
            // 
            btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            btnClose.Location = new System.Drawing.Point(639, 617);
            btnClose.Name = "btnClose";
            btnClose.Size = new System.Drawing.Size(111, 23);
            btnClose.TabIndex = 5;
            btnClose.Text = "Close";
            btnClose.UseVisualStyleBackColor = true;
            btnClose.Click += btnClose_Click;
            // 
            // btnSearch
            // 
            btnSearch.Location = new System.Drawing.Point(594, 47);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new System.Drawing.Size(75, 23);
            btnSearch.TabIndex = 7;
            btnSearch.Text = "Search";
            btnSearch.UseVisualStyleBackColor = true;
            btnSearch.Click += btnSearch_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(12, 50);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(61, 15);
            label2.TabIndex = 9;
            label2.Text = "Search by:";
            // 
            // cmbSearchCat
            // 
            cmbSearchCat.FormattingEnabled = true;
            cmbSearchCat.Location = new System.Drawing.Point(80, 47);
            cmbSearchCat.Name = "cmbSearchCat";
            cmbSearchCat.Size = new System.Drawing.Size(107, 23);
            cmbSearchCat.TabIndex = 10;
            // 
            // btnReset
            // 
            btnReset.Location = new System.Drawing.Point(675, 47);
            btnReset.Name = "btnReset";
            btnReset.Size = new System.Drawing.Size(75, 23);
            btnReset.TabIndex = 11;
            btnReset.Text = "Reset Data";
            btnReset.UseVisualStyleBackColor = true;
            btnReset.Click += btnReset_Click;
            // 
            // chkMatchCase
            // 
            chkMatchCase.AutoSize = true;
            chkMatchCase.Location = new System.Drawing.Point(507, 49);
            chkMatchCase.Name = "chkMatchCase";
            chkMatchCase.Size = new System.Drawing.Size(88, 19);
            chkMatchCase.TabIndex = 12;
            chkMatchCase.Text = "Match Case";
            chkMatchCase.UseVisualStyleBackColor = true;
            // 
            // btnExport
            // 
            btnExport.Location = new System.Drawing.Point(327, 617);
            btnExport.Name = "btnExport";
            btnExport.Size = new System.Drawing.Size(134, 23);
            btnExport.TabIndex = 13;
            btnExport.Text = "Export Selected to CSV";
            btnExport.UseVisualStyleBackColor = true;
            btnExport.Click += btnExport_Click;
            // 
            // frmHome
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(762, 652);
            Controls.Add(btnExport);
            Controls.Add(chkMatchCase);
            Controls.Add(btnReset);
            Controls.Add(cmbSearchCat);
            Controls.Add(label2);
            Controls.Add(btnSearch);
            Controls.Add(btnClose);
            Controls.Add(btnGetAccess);
            Controls.Add(dataGridView1);
            Controls.Add(txtSearch);
            Controls.Add(label1);
            MaximizeBox = false;
            Name = "frmHome";
            Text = "Home";
            Load += frmHome_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnGetAccess;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbSearchCat;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.CheckBox chkMatchCase;
        private System.Windows.Forms.Button btnExport;
    }
}
