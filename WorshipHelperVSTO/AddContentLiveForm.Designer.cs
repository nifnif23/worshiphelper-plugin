namespace WorshipHelperVSTO
{
    partial class AddContentLiveForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddContentLiveForm));

            // ---------------------------------------------------------------
            // Control declarations
            // ---------------------------------------------------------------
            this.panelHeader = new System.Windows.Forms.Panel();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblSubtitle = new System.Windows.Forms.Label();
            this.btnScripture = new System.Windows.Forms.Button();
            this.btnSong = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.lblInfo = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panelHeader.SuspendLayout();
            this.SuspendLayout();

            // ---------------------------------------------------------------
            // panelHeader — colored banner at the top (matches InsertScriptureForm)
            // ---------------------------------------------------------------
            this.panelHeader.BackColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(480, 48);
            this.panelHeader.TabIndex = 100;
            this.panelHeader.Controls.Add(this.lblTitle);

            // lblTitle
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI Semibold", 14F, System.Drawing.FontStyle.Bold);
            this.lblTitle.ForeColor = System.Drawing.Color.White;
            this.lblTitle.Location = new System.Drawing.Point(16, 10);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Text = "\u2795  Add Content Live";

            // ---------------------------------------------------------------
            // Layout metrics
            // ---------------------------------------------------------------
            int leftMargin = 24;
            int btnWidth = 200;
            int btnHeight = 120;
            int spacing = 20;
            int row = 68;

            // ---------------------------------------------------------------
            // Subtitle / instruction label
            // ---------------------------------------------------------------
            this.lblSubtitle.AutoSize = true;
            this.lblSubtitle.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblSubtitle.ForeColor = System.Drawing.Color.FromArgb(100, 100, 100);
            this.lblSubtitle.Location = new System.Drawing.Point(leftMargin, row);
            this.lblSubtitle.Name = "lblSubtitle";
            this.lblSubtitle.Text = "Choose what to insert during the live presentation:";

            row += 30;

            // ---------------------------------------------------------------
            // imageList1 — icons for buttons
            // ---------------------------------------------------------------
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "bible.png");
            this.imageList1.Images.SetKeyName(1, "music-note.png");

            // ---------------------------------------------------------------
            // btnScripture — flat modern card-style button
            // ---------------------------------------------------------------
            int btnLeftScripture = (480 - btnWidth * 2 - spacing) / 2;
            int btnLeftSong = btnLeftScripture + btnWidth + spacing;

            this.btnScripture.BackColor = System.Drawing.Color.White;
            this.btnScripture.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnScripture.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.btnScripture.FlatAppearance.BorderSize = 2;
            this.btnScripture.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(235, 238, 248);
            this.btnScripture.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold);
            this.btnScripture.ForeColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.btnScripture.ImageIndex = 0;
            this.btnScripture.ImageList = this.imageList1;
            this.btnScripture.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnScripture.Location = new System.Drawing.Point(btnLeftScripture, row);
            this.btnScripture.Name = "btnScripture";
            this.btnScripture.Padding = new System.Windows.Forms.Padding(0, 12, 0, 6);
            this.btnScripture.Size = new System.Drawing.Size(btnWidth, btnHeight);
            this.btnScripture.TabIndex = 0;
            this.btnScripture.Text = "&Scripture";
            this.btnScripture.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnScripture.UseVisualStyleBackColor = false;
            this.btnScripture.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnScripture.Click += new System.EventHandler(this.btnScripture_Click);

            // ---------------------------------------------------------------
            // btnSong — flat modern card-style button
            // ---------------------------------------------------------------
            this.btnSong.BackColor = System.Drawing.Color.White;
            this.btnSong.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSong.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.btnSong.FlatAppearance.BorderSize = 2;
            this.btnSong.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(235, 238, 248);
            this.btnSong.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold);
            this.btnSong.ForeColor = System.Drawing.Color.FromArgb(55, 71, 133);
            this.btnSong.ImageIndex = 1;
            this.btnSong.ImageList = this.imageList1;
            this.btnSong.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnSong.Location = new System.Drawing.Point(btnLeftSong, row);
            this.btnSong.Name = "btnSong";
            this.btnSong.Padding = new System.Windows.Forms.Padding(0, 12, 0, 6);
            this.btnSong.Size = new System.Drawing.Size(btnWidth, btnHeight);
            this.btnSong.TabIndex = 1;
            this.btnSong.Text = "Song or &Presentation";
            this.btnSong.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnSong.UseVisualStyleBackColor = false;
            this.btnSong.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSong.Click += new System.EventHandler(this.btnSong_Click);

            row += btnHeight + 16;

            // ---------------------------------------------------------------
            // lblInfo — informational text
            // ---------------------------------------------------------------
            this.lblInfo.AutoSize = true;
            this.lblInfo.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.lblInfo.ForeColor = System.Drawing.Color.Gray;
            this.lblInfo.Location = new System.Drawing.Point(leftMargin, row);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Text = "The added content will be inserted after the currently displayed slide.";

            row += 28;

            // ---------------------------------------------------------------
            // btnCancel — flat cancel button
            // ---------------------------------------------------------------
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Segoe UI", 9.5F);
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnCancel.Location = new System.Drawing.Point(leftMargin, row);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(110, 34);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);

            row += 50;

            // ---------------------------------------------------------------
            // AddContentLiveForm
            // ---------------------------------------------------------------
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(480, row + 10);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.BackColor = System.Drawing.Color.White;
            this.Name = "AddContentLiveForm";
            this.Text = "Add Content Live — WorshipHelper";

            this.Controls.Add(this.panelHeader);
            this.Controls.Add(this.lblSubtitle);
            this.Controls.Add(this.btnScripture);
            this.Controls.Add(this.btnSong);
            this.Controls.Add(this.lblInfo);
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
        private System.Windows.Forms.Label lblSubtitle;
        private System.Windows.Forms.Button btnScripture;
        private System.Windows.Forms.Button btnSong;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Button btnCancel;
    }
}
