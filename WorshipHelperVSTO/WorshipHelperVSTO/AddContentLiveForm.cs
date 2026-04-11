using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Windows.Forms;

namespace WorshipHelperVSTO
{
    public partial class AddContentLiveForm : Form
    {
        public AddContentLiveForm()
        {
            InitializeComponent();
        }

        private void btnScripture_Click(object sender, EventArgs e)
        {
            try
            {
                // Open the scripture form as a dialog; this form stays open behind it.
                // When the scripture form closes (after insert or cancel), we close too.
                using (var scriptureForm = new InsertScriptureForm())
                {
                    scriptureForm.ShowDialog(this);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"An error occurred while opening the Scripture form:\n\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnSong_Click(object sender, EventArgs e)
        {
            try
            {
                new SongManager().InsertSong();

                // After inserting the song, try to return focus to the presenter view
                try
                {
                    DocumentWindow presenterView = new WindowManager().GetPresenterView();
                    if (presenterView != null)
                    {
                        presenterView.Activate();
                    }
                }
                catch
                {
                    // FIX: Silently ignore focus errors — the song was already inserted
                    // and we don't want to show an error about window focus.
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"An error occurred while inserting the song:\n\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
