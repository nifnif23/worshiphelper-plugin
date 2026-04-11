namespace WorshipHelperVSTO
{
    partial class InsertScriptureForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            // ---------------------------------------------------------------
            // Control declarations
            // ---------------------------------------------------------------
            this.lblTemplate = new System.Windows.Forms.Label();
            this.cmbTemplate = new System.Windows.Forms.ComboBox();
            this.lblTranslation = new System.Windows.Forms.Label();
            this.cmbTranslation = new System.Windows.Forms.ComboBox();
            this.lblBook = new System.Windows.Forms.Label();
            this.txtBook = new System.Windows.Forms.TextBox();
            this.lblReference = new System.Windows.Forms.Label();
            this.txtReference = new System.Windows.Forms.TextBox();
            this.lblBulk = new System.Windows.Forms.Label();
            this.txtBulk = new System.Windows.Forms.TextBox();
            this.lblBulkHint = new System.Windows.Forms.Label();
            this.chkMultiVerse = new System.Windows.Forms.CheckBox();
            this.btnInsert = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnModeBulk = new System.Windows.Forms.Button();
            this.btnModeSingle = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.lblTitle = new System.Windows.Forms.Label();
            this.panelHeader.SuspendLayout();
            this.SuspendLayout();

            // ---------------------------------------------------------------
            // panelHeader — colored banner at the top
            // ---------------------------------------------------------------
            this.panelHeader.BackColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(520, 48);
            this.panelHeader.TabIndex = 100;
            this.panelHeader.Controls.Add(this.lblTitle);

            // lblTitle
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI Semibold", 14F, System.Drawing.FontStyle.Bold);
            this.lblTitle.ForeColor = System.Drawing.Color.White;
            this.lblTitle.Location = new System.Drawing.Point(16, 10);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Text = "\u2702  Insert Scripture";

            // ---------------------------------------------------------------
            // Layout metrics
            // ---------------------------------------------------------------
            int leftLabel = 20;
            int leftCtrl = 120;
            int ctrlWidth = 370;
            int row = 64;
            int rowH = 34;

            // ---------------------------------------------------------------
            // Row 1: Template
            // ---------------------------------------------------------------
            this.lblTemplate.AutoSize = true;
            this.lblTemplate.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblTemplate.Location = new System.Drawing.Point(leftLabel, row + 3);
            this.lblTemplate.Name = "lblTemplate";
            this.lblTemplate.Text = "Template:";

            this.cmbTemplate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTemplate.Font = new System.Drawing.Font("Segoe UI", 9.5F);
            this.cmbTemplate.FormattingEnabled = true;
            this.cmbTemplate.Location = new System.Drawing.Point(leftCtrl, row);
            this.cmbTemplate.Name = "cmbTemplate";
            this.cmbTemplate.Size = new System.Drawing.Size(ctrlWidth, 25);
            this.cmbTemplate.TabIndex = 0;
            this.cmbTemplate.SelectionChangeCommitted += new System.EventHandler(this.cmbTemplate_SelectionChangeCommitted);

            row += rowH;

            // ---------------------------------------------------------------
            // Row 2: Translation
            // ---------------------------------------------------------------
            this.lblTranslation.AutoSize = true;
            this.lblTranslation.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblTranslation.Location = new System.Drawing.Point(leftLabel, row + 3);
            this.lblTranslation.Name = "lblTranslation";
            this.lblTranslation.Text = "Translation:";

            this.cmbTranslation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTranslation.Font = new System.Drawing.Font("Segoe UI", 9.5F);
            this.cmbTranslation.FormattingEnabled = true;
            this.cmbTranslation.Location = new System.Drawing.Point(leftCtrl, row);
            this.cmbTranslation.Name = "cmbTranslation";
            this.cmbTranslation.Size = new System.Drawing.Size(ctrlWidth, 25);
            this.cmbTranslation.TabIndex = 1;
            this.cmbTranslation.SelectionChangeCommitted += new System.EventHandler(this.cmbTranslation_SelectionChangeCommitted);

            row += rowH + 6;

            // ---------------------------------------------------------------
            // Row 3: Book (single mode)
            // ---------------------------------------------------------------
            this.lblBook.AutoSize = true;
            this.lblBook.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblBook.Location = new System.Drawing.Point(leftLabel, row + 3);
            this.lblBook.Name = "lblBook";
            this.lblBook.Text = "Book:";

            this.txtBook.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.txtBook.Location = new System.Drawing.Point(leftCtrl, row);
            this.txtBook.Name = "txtBook";
            this.txtBook.Size = new System.Drawing.Size(ctrlWidth, 26);
            this.txtBook.TabIndex = 2;
            this.txtBook.TextChanged += new System.EventHandler(this.txtSearchBox_TextChanged);
            this.txtBook.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSearchBox_KeyPress);
            this.txtBook.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBook_KeyDown);

            row += rowH;

            // ---------------------------------------------------------------
            // Row 4: Reference (single mode)
            // ---------------------------------------------------------------
            this.lblReference.AutoSize = true;
            this.lblReference.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblReference.Location = new System.Drawing.Point(leftLabel, row + 3);
            this.lblReference.Name = "lblReference";
            this.lblReference.Text = "Reference:";

            this.txtReference.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.txtReference.Location = new System.Drawing.Point(leftCtrl, row);
            this.txtReference.Name = "txtReference";
            this.txtReference.Size = new System.Drawing.Size(ctrlWidth, 26);
            this.txtReference.TabIndex = 3;
            this.txtReference.TextChanged += new System.EventHandler(this.txtReference_TextChanged);

            // ---------------------------------------------------------------
            // Bulk mode: label + multiline textbox (overlaps rows 3–4 area)
            // ---------------------------------------------------------------
            int bulkTop = this.txtBook.Location.Y;

            this.lblBulk.AutoSize = true;
            this.lblBulk.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblBulk.Location = new System.Drawing.Point(leftLabel, bulkTop + 3);
            this.lblBulk.Name = "lblBulk";
            this.lblBulk.Text = "References:";
            this.lblBulk.Visible = false;

            this.txtBulk.Font = new System.Drawing.Font("Segoe UI", 9.5F);
            this.txtBulk.Location = new System.Drawing.Point(leftCtrl, bulkTop);
            this.txtBulk.Multiline = true;
            this.txtBulk.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtBulk.AcceptsReturn = true;
            this.txtBulk.Name = "txtBulk";
            this.txtBulk.Size = new System.Drawing.Size(ctrlWidth, 62);
            this.txtBulk.TabIndex = 4;
            this.txtBulk.Visible = false;
            this.txtBulk.TextChanged += new System.EventHandler(this.txtBulk_TextChanged);

            this.lblBulkHint.AutoSize = true;
            this.lblBulkHint.Font = new System.Drawing.Font("Segoe UI", 7.5F, System.Drawing.FontStyle.Italic);
            this.lblBulkHint.ForeColor = System.Drawing.Color.Gray;
            this.lblBulkHint.Location = new System.Drawing.Point(leftCtrl, bulkTop + 65);
            this.lblBulkHint.Name = "lblBulkHint";
            this.lblBulkHint.Text = "e.g.  John 3:16-18;  Romans 8:28;  Ps 23:1-6   (one per line or separated by ; )";
            this.lblBulkHint.Visible = false;

            row += rowH + 4;

            // ---------------------------------------------------------------
            // Row 5: Multi-verse checkbox
            // ---------------------------------------------------------------
            this.chkMultiVerse.AutoSize = true;
            this.chkMultiVerse.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.chkMultiVerse.Location = new System.Drawing.Point(leftCtrl, row);
            this.chkMultiVerse.Name = "chkMultiVerse";
            this.chkMultiVerse.Size = new System.Drawing.Size(250, 20);
            this.chkMultiVerse.TabIndex = 5;
            this.chkMultiVerse.Text = "Pack multiple verses per slide (multi-verse)";
            this.chkMultiVerse.UseVisualStyleBackColor = true;
            this.chkMultiVerse.CheckedChanged += new System.EventHandler(this.chkMultiVerse_CheckedChanged);

            row += rowH;

            // ---------------------------------------------------------------
            // Row 6: Mode toggle buttons
            // ---------------------------------------------------------------
            this.btnModeBulk.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnModeBulk.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnModeBulk.ForeColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.btnModeBulk.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.btnModeBulk.Location = new System.Drawing.Point(leftCtrl, row);
            this.btnModeBulk.Name = "btnModeBulk";
            this.btnModeBulk.Size = new System.Drawing.Size(160, 28);
            this.btnModeBulk.TabIndex = 6;
            this.btnModeBulk.Text = "\u2b07  Switch to Bulk Paste";
            this.btnModeBulk.UseVisualStyleBackColor = true;
            this.btnModeBulk.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnModeBulk.Click += new System.EventHandler(this.btnModeBulk_Click);

            this.btnModeSingle.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.btnModeSingle.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnModeSingle.ForeColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.btnModeSingle.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.btnModeSingle.Location = new System.Drawing.Point(leftCtrl, row);
            this.btnModeSingle.Name = "btnModeSingle";
            this.btnModeSingle.Size = new System.Drawing.Size(160, 28);
            this.btnModeSingle.TabIndex = 7;
            this.btnModeSingle.Text = "\u2b06  Switch to Single Entry";
            this.btnModeSingle.UseVisualStyleBackColor = true;
            this.btnModeSingle.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnModeSingle.Visible = false;
            this.btnModeSingle.Click += new System.EventHandler(this.btnModeSingle_Click);

            row += rowH + 2;

            // ---------------------------------------------------------------
            // Row 7: Status label
            // ---------------------------------------------------------------
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.lblStatus.ForeColor = System.Drawing.Color.Gray;
            this.lblStatus.Location = new System.Drawing.Point(leftCtrl, row);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Text = "";

            row += 22;

            // ---------------------------------------------------------------
            // Row 8: Buttons
            // ---------------------------------------------------------------
            this.btnInsert.Font = new System.Drawing.Font("Segoe UI Semibold", 9.5F, System.Drawing.FontStyle.Bold);
            this.btnInsert.BackColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.btnInsert.ForeColor = System.Drawing.Color.White;
            this.btnInsert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInsert.FlatAppearance.BorderSize = 0;
            this.btnInsert.Location = new System.Drawing.Point(leftCtrl, row);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(130, 34);
            this.btnInsert.TabIndex = 8;
            this.btnInsert.Text = "Insert";
            this.btnInsert.UseVisualStyleBackColor = false;
            this.btnInsert.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);

            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Segoe UI", 9.5F);
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnCancel.Location = new System.Drawing.Point(leftCtrl + 140, row);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(110, 34);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "Close";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);

            row += 50;

            // ---------------------------------------------------------------
            // Form
            // ---------------------------------------------------------------
            this.AcceptButton = this.btnInsert;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(520, row + 10);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.BackColor = System.Drawing.Color.White;
            this.Name = "InsertScriptureForm";
            this.Text = "Insert Scripture — WorshipHelper";

            this.Controls.Add(this.panelHeader);
            this.Controls.Add(this.lblTemplate);
            this.Controls.Add(this.cmbTemplate);
            this.Controls.Add(this.lblTranslation);
            this.Controls.Add(this.cmbTranslation);
            this.Controls.Add(this.lblBook);
            this.Controls.Add(this.txtBook);
            this.Controls.Add(this.lblReference);
            this.Controls.Add(this.txtReference);
            this.Controls.Add(this.lblBulk);
            this.Controls.Add(this.txtBulk);
            this.Controls.Add(this.lblBulkHint);
            this.Controls.Add(this.chkMultiVerse);
            this.Controls.Add(this.btnModeBulk);
            this.Controls.Add(this.btnModeSingle);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.btnCancel);

            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        // ---------------------------------------------------------------
        // Field declarations
        // ---------------------------------------------------------------
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblTemplate;
        private System.Windows.Forms.ComboBox cmbTemplate;
        private System.Windows.Forms.Label lblTranslation;
        private System.Windows.Forms.ComboBox cmbTranslation;
        private System.Windows.Forms.Label lblBook;
        private System.Windows.Forms.TextBox txtBook;
        private System.Windows.Forms.Label lblReference;
        private System.Windows.Forms.TextBox txtReference;
        private System.Windows.Forms.Label lblBulk;
        private System.Windows.Forms.TextBox txtBulk;
        private System.Windows.Forms.Label lblBulkHint;
        private System.Windows.Forms.CheckBox chkMultiVerse;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnModeBulk;
        private System.Windows.Forms.Button btnModeSingle;
        private System.Windows.Forms.Label lblStatus;
    }
}
