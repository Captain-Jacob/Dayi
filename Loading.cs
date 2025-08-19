using System;
using System.Windows.Forms;

namespace ProgramKontrol  //yüklenme ekranı millet program bozuldu sanmasın diye koydum
{
    public partial class LoadingForm : Form
    {
        public LoadingForm()
        {
            InitializeComponent();
            this.ControlBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        public void SetMessage(string message)
        {
            labelMessage.Text = message;
            labelMessage.Refresh();
        }
    }
}
