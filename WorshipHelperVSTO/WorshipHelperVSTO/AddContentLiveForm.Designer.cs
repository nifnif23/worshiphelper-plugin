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

            // ---------------------------------------------------------------
            // Color scheme: Green / White / Gold
            // ---------------------------------------------------------------
            var accentGreen = System.Drawing.Color.FromArgb(46, 125, 50);       // #2E7D32
            var darkGreen   = System.Drawing.Color.FromArgb(27, 94, 32);        // #1B5E20
            var lightGreen  = System.Drawing.Color.FromArgb(232, 245, 233);     // #E8F5E9
            var hoverGreen  = System.Drawing.Color.FromArgb(200, 230, 201);     // #C8E6C9
            var gold        = System.Drawing.Color.FromArgb(184, 134, 11);      // #B8860B
            var textDark    = System.Drawing.Color.FromArgb(33, 33, 33);        // #212121
            var textMuted   = System.Drawing.Color.FromArgb(117, 117, 117);     // #757575

            // ---------------------------------------------------------------
            // Control declarations
            // ---------------------------------------------------------------
            this.panelHeader = new System.Windows.Forms.Panel();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblSubtitle = new System.Windows.Forms.Label();
            this.btnScripture = new System.Windows.Forms.Button();
            this.btnSong = new System.Windows.Forms.Button();
            this.lblInfo = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panelHeader.SuspendLayout();
            this.SuspendLayout();

            // ---------------------------------------------------------------
            // panelHeader — green banner at the top
            // ---------------------------------------------------------------
            this.panelHeader.BackColor = accentGreen;
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(440, 52);
            this.panelHeader.TabIndex = 100;
            this.panelHeader.Controls.Add(this.lblTitle);

            // lblTitle — use BMP-safe cross character
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI Semibold", 13F, System.Drawing.FontStyle.Bold);
            this.lblTitle.ForeColor = System.Drawing.Color.White;
            this.lblTitle.Location = new System.Drawing.Point(18, 13);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Text = "\u271A  Add Content Live";

            // ---------------------------------------------------------------
            // Layout metrics
            // ---------------------------------------------------------------
            int leftMargin = 24;
            int btnWidth = 180;
            int btnHeight = 100;
            int spacing = 16;
            int row = 72;

            // ---------------------------------------------------------------
            // Subtitle / instruction label
            // ---------------------------------------------------------------
            this.lblSubtitle.AutoSize = true;
            this.lblSubtitle.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblSubtitle.ForeColor = textMuted;
            this.lblSubtitle.Location = new System.Drawing.Point(leftMargin, row);
            this.lblSubtitle.Name = "lblSubtitle";
            this.lblSubtitle.Text = "Choose what to insert during the live presentation:";

            row += 28;

            // ---------------------------------------------------------------
            // btnScripture — card-style button with icon from resources
            // FIX: Use Properties.Resources directly instead of broken ImageList.
            //      Replace supplementary-plane Unicode emoji with BMP-safe chars.
            // ---------------------------------------------------------------
            int btnLeftScripture = (440 - btnWidth * 2 - spacing) / 2;
            int btnLeftSong = btnLeftScripture + btnWidth + spacing;

            this.btnScripture.BackColor = System.Drawing.Color.White;
            this.btnScripture.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnScripture.FlatAppearance.BorderColor = accentGreen;
            this.btnScripture.FlatAppearance.BorderSize = 2;
            this.btnScripture.FlatAppearance.MouseOverBackColor = hoverGreen;
            this.btnScripture.Font = new System.Drawing.Font("Segoe UI Semibold", 10F, System.Drawing.FontStyle.Bold);
            this.btnScripture.ForeColor = darkGreen;
            // FIX: Load image directly from compiled resources (not ImageList)
            this.btnScripture.Image = new System.Drawing.Bitmap(global::WorshipHelperVSTO.Properties.Resources.bible, 48, 48);
            this.btnScripture.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnScripture.Location = new System.Drawing.Point(btnLeftScripture, row);
            this.btnScripture.Name = "btnScripture";
            this.btnScripture.Padding = new System.Windows.Forms.Padding(0, 10, 0, 4);
            this.btnScripture.Size = new System.Drawing.Size(btnWidth, btnHeight);
            this.btnScripture.TabIndex = 0;
            // FIX: Use BMP-safe text (no supplementary plane emoji)
            this.btnScripture.Text = "&Scripture";
            this.btnScripture.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnScripture.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnScripture.UseVisualStyleBackColor = false;
            this.btnScripture.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnScripture.Click += new System.EventHandler(this.btnScripture_Click);

            // ---------------------------------------------------------------
            // btnSong — card-style button with icon from resources
            // ---------------------------------------------------------------
            this.btnSong.BackColor = System.Drawing.Color.White;
            this.btnSong.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSong.FlatAppearance.BorderColor = accentGreen;
            this.btnSong.FlatAppearance.BorderSize = 2;
            this.btnSong.FlatAppearance.MouseOverBackColor = hoverGreen;
            this.btnSong.Font = new System.Drawing.Font("Segoe UI Semibold", 10F, System.Drawing.FontStyle.Bold);
            this.btnSong.ForeColor = darkGreen;
            // FIX: Load image directly from compiled resources (not ImageList)
            this.btnSong.Image = new System.Drawing.Bitmap(global::WorshipHelperVSTO.Properties.Resources.music_note, 48, 48);
            this.btnSong.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnSong.Location = new System.Drawing.Point(btnLeftSong, row);
            this.btnSong.Name = "btnSong";
            this.btnSong.Padding = new System.Windows.Forms.Padding(0, 10, 0, 4);
            this.btnSong.Size = new System.Drawing.Size(btnWidth, btnHeight);
            this.btnSong.TabIndex = 1;
            // FIX: Use BMP-safe text (no supplementary plane emoji)
            this.btnSong.Text = "Song / &Presentation";
            this.btnSong.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnSong.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnSong.UseVisualStyleBackColor = false;
            this.btnSong.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSong.Click += new System.EventHandler(this.btnSong_Click);

            row += btnHeight + 14;

            // ---------------------------------------------------------------
            // lblInfo — informational text
            // ---------------------------------------------------------------
            this.lblInfo.AutoSize = true;
            this.lblInfo.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.lblInfo.ForeColor = textMuted;
            this.lblInfo.Location = new System.Drawing.Point(leftMargin, row);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Text = "Content will be inserted after the currently displayed slide.";

            row += 26;

            // ---------------------------------------------------------------
            // btnCancel — flat cancel button
            // ---------------------------------------------------------------
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(189, 189, 189);
            this.btnCancel.ForeColor = textMuted;
            this.btnCancel.Location = new System.Drawing.Point(leftMargin, row);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 32);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);

            row += 46;

            // ---------------------------------------------------------------
            // AddContentLiveForm
            // ---------------------------------------------------------------
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(440, row + 8);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.BackColor = System.Drawing.Color.White;
            this.Name = "AddContentLiveForm";
            this.Text = "Add Content \u2014 WorshipHelper";
            this.TopMost = true;

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
        // FIX: Removed imageList1 — no longer needed since we load
        //      images directly from Properties.Resources.
        // ---------------------------------------------------------------
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblSubtitle;
        private System.Windows.Forms.Button btnScripture;
        private System.Windows.Forms.Button btnSong;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Button btnCancel;
    }
}
