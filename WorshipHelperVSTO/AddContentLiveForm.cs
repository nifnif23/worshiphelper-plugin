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
            new InsertScriptureForm().ShowDialog();
            Close();
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
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
