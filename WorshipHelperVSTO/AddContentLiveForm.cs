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
            // Open the scripture form as a dialog; this form stays open behind it.
            // When the scripture form closes (after insert or cancel), we close too.
            using (var scriptureForm = new InsertScriptureForm())
            {
                scriptureForm.ShowDialog(this);
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnSong_Click(object sender, EventArgs e)
        {
            new SongManager().InsertSong();
            // After inserting the song, try to return focus to the presenter view
            DocumentWindow presenterView = new WindowManager().GetPresenterView();
            if (presenterView != null)
            {
                presenterView.Activate();
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
