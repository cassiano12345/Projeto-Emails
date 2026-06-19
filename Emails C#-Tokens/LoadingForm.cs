using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Softnova_Emails
{
    public partial class LoadingForm : Form
    {
        public LoadingForm()
        {
            InitializeComponent();
            this.ControlBox = false; // Remove botões de controle
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None; // Sem bordas
            this.ShowInTaskbar = false;

            // Adiciona uma imagem de carregamento
            PictureBox pictureBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                Image = Properties.Resources.LoadingGif, // Certifique-se de adicionar um GIF aos recursos do projeto
                SizeMode = PictureBoxSizeMode.CenterImage
            };
            this.Controls.Add(pictureBox);
        }

        private void LoadingForm_Load(object sender, EventArgs e)
        {

        }
    }
}
