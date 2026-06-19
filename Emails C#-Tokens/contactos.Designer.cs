
namespace Softnova_Emails
{
    partial class contactos
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(contactos));
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.PictureBoxSearch = new System.Windows.Forms.PictureBox();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBoxSearch)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(12, 9);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(753, 260);
            this.dataGridView.TabIndex = 0;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 280);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(168, 20);
            this.textBox1.TabIndex = 1;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // PictureBoxSearch
            // 
            this.PictureBoxSearch.BackColor = System.Drawing.Color.White;
            this.PictureBoxSearch.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("PictureBoxSearch.BackgroundImage")));
            this.PictureBoxSearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.PictureBoxSearch.InitialImage = ((System.Drawing.Image)(resources.GetObject("PictureBoxSearch.InitialImage")));
            this.PictureBoxSearch.Location = new System.Drawing.Point(181, 280);
            this.PictureBoxSearch.Name = "PictureBoxSearch";
            this.PictureBoxSearch.Size = new System.Drawing.Size(27, 20);
            this.PictureBoxSearch.TabIndex = 34;
            this.PictureBoxSearch.TabStop = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(685, 280);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(80, 23);
            this.button1.TabIndex = 35;
            this.button1.Text = "Fechar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // contactos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClientSize = new System.Drawing.Size(777, 310);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.PictureBoxSearch);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.dataGridView);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "contactos";
            this.Text = "contactos";
            this.Load += new System.EventHandler(this.contactos_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBoxSearch)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.TextBox textBox1;
        internal System.Windows.Forms.PictureBox PictureBoxSearch;
        private System.Windows.Forms.Button button1;
    }
}