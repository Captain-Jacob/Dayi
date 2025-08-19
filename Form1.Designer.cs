using System;
using System.Drawing;
using System.Windows.Forms;

namespace ProgramKontrol          //thise gerek yok biliyorum ama dursun zararı yok
{
      partial class LoadingForm
    {
        private Label labelMessage;

        private void InitializeComponent()
        {
            this.labelMessage = new Label();
            this.SuspendLayout();

            // labelMessage
            this.labelMessage.AutoSize = false;
            this.labelMessage.Dock = DockStyle.Fill;
            this.labelMessage.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            this.labelMessage.ForeColor = Color.DarkBlue;
            this.labelMessage.TextAlign = ContentAlignment.MiddleCenter;
            this.labelMessage.Text = "Çay demlensin geleceğim";

            // LoadingForm
            this.ClientSize = new Size(400, 100);
            this.Controls.Add(this.labelMessage);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.Name = "Dayı";
            this.Text = "Dayı çayı demliyor. . .";
            this.ResumeLayout(false);
        }
    }


    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private Label labelNot;
        private Label labelNot2;
        private TextBox textBoxNot;
        private TextBox textBoxNot2;
        private Label labelBilgi;
        private Label labelYuzde;
        private ListBox listBoxYuklu;
        private ListBox listBoxEksik;
        private Label labelYukluBaslik;
        private Label labelEksikBaslik;
        private Button btnKaydet;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent() 
        {
            this.labelBilgi = new Label();
            this.labelYuzde = new Label();
            this.listBoxYuklu = new ListBox();
            this.listBoxEksik = new ListBox();
            this.labelYukluBaslik = new Label();
            this.labelEksikBaslik = new Label();
            this.btnKaydet = new Button();
            this.SuspendLayout();

            // Form
            this.ClientSize = new Size(700, 400);
            this.Text = "Dayı";
            this.BackColor = Color.WhiteSmoke;

            // labelBilgi
            this.labelBilgi.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            this.labelBilgi.Location = new Point(20, 20);
            this.labelBilgi.Size = new Size(660, 30);
            this.labelBilgi.Text = "Bilgisayar Bilgisi";

            // labelYuzde
            this.labelYuzde.Font = new Font("Segoe UI", 24F, FontStyle.Bold);
            this.labelYuzde.Location = new Point(550, 60);
            this.labelYuzde.Size = new Size(120, 120);
            this.labelYuzde.TextAlign = ContentAlignment.MiddleCenter;
            this.labelYuzde.BackColor = Color.LightGray;
            this.labelYuzde.Text = "0%";
            this.labelYuzde.BorderStyle = BorderStyle.FixedSingle;

            // labelYukluBaslik
            this.labelYukluBaslik.Text = "✔ Yüklü Programlar";
            this.labelYukluBaslik.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.labelYukluBaslik.ForeColor = Color.Green;
            this.labelYukluBaslik.Location = new Point(20, 70);
            this.labelYukluBaslik.Size = new Size(200, 20);

            // labelEksikBaslik
            this.labelEksikBaslik.Text = "✖ Yüklü Olmayanlar";
            this.labelEksikBaslik.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            this.labelEksikBaslik.ForeColor = Color.Red;
            this.labelEksikBaslik.Location = new Point(270, 70);
            this.labelEksikBaslik.Size = new Size(200, 20);

            // listBoxYuklu
            this.listBoxYuklu.Location = new Point(20, 100);
            this.listBoxYuklu.Size = new Size(230, 250);
            this.listBoxYuklu.ForeColor = Color.Green;
            this.listBoxYuklu.Font = new Font("Segoe UI", 9F);

            // listBoxEksik
            this.listBoxEksik.Location = new Point(270, 100);
            this.listBoxEksik.Size = new Size(230, 250);
            this.listBoxEksik.ForeColor = Color.Red;
            this.listBoxEksik.Font = new Font("Segoe UI", 9F);

            // labelNot
            this.labelNot = new Label();
            this.labelNot.Text = "Eski PC:";
            this.labelNot.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.labelNot.Location = new Point(550, 220);
            this.labelNot.Size = new Size(120, 20);

            // labelNot 2
            this.labelNot2 = new Label();
            this.labelNot2.Text = "Eski Seri no:";
            this.labelNot2.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            this.labelNot2.Location = new Point(550, 270);
            this.labelNot2.Size = new Size(120, 20);


            // textBoxNot
            this.textBoxNot = new TextBox();
            this.textBoxNot.Location = new Point(550, 250);
            this.textBoxNot.Size = new Size(120, 40);
            this.textBoxNot.Font = new Font("Segoe UI", 9F);

            // textBoxNot2
            this.textBoxNot2 = new TextBox();
            this.textBoxNot2.Location = new Point(550, 200);
            this.textBoxNot2.Size = new Size(120, 40);
            this.textBoxNot2.Font = new Font("Segoe UI", 9F);


            // btnKaydet
            this.btnKaydet.Text = "Excele Kaydet";
            this.btnKaydet.Size = new Size(120, 40);
            this.btnKaydet.Location = new Point(550, 300);
            this.btnKaydet.Click += new EventHandler(this.BtnKaydet_Click);

            // Controls Add
            this.Controls.Add(this.labelBilgi);
            this.Controls.Add(this.labelYuzde);
            this.Controls.Add(this.labelYukluBaslik);
            this.Controls.Add(this.labelEksikBaslik);
            this.Controls.Add(this.listBoxYuklu);
            this.Controls.Add(this.listBoxEksik);
            this.Controls.Add(this.btnKaydet); 
            this.Controls.Add(this.labelNot);
            this.Controls.Add(this.labelNot2);
            this.Controls.Add(this.textBoxNot);
            this.Controls.Add(this.textBoxNot2);
            this.ResumeLayout(false);

        }
    }
}
