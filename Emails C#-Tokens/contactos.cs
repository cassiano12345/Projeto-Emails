using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Softnova_Emails
{

    public partial class contactos : Form
    {

        public contactos()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;


            dataGridView.ReadOnly = true;
            // Adicionar colunas ao DataGridView
            dataGridView.Columns.Add("Name", "Nome");
            dataGridView.Columns.Add("Email", "E-mail");
            dataGridView.Columns["Name"].Width = 130; // Largura da coluna "Nome"
            dataGridView.Columns["Email"].Width = 480; // Largura da coluna "E-mail"
            // Adicionar uma coluna de botão
            var buttonColumn = new DataGridViewButtonColumn
            {
                Name = "ActionButton", // Nome único para a coluna
                HeaderText = "Ações",
                Text = "Adicionar",
                UseColumnTextForButtonValue = true
            };
            dataGridView.Columns.Add(buttonColumn);

            // Adicionar evento para cliques no botão
            dataGridView.CellContentClick += DataGridView_CellContentClick;

            // Adicionar o DataGridView ao formulário
            this.Controls.Add(dataGridView);

            // Carregar os dados
            LoadContacts();
        }

        private void LoadContacts()
        {
            // Simulando a leitura de contatos de um arquivo CSV
            string filePath = "../../../Meus_contactos/contactos.csv";
            var contacts = ReadContactsFromCsv(filePath);

            // Adicionar os contatos ao DataGridView
            foreach (var contact in contacts)
            {
                dataGridView.Rows.Add(contact.Name, contact.Email);
            }
        }

        private List<Contact> ReadContactsFromCsv(string filePath)
        {
            var contacts = new List<Contact>();

            // Ler o arquivo CSV
            using (var reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var values = line.Split(';');

                    if (values.Length >= 2)
                    {
                        contacts.Add(new Contact
                        {
                            Name = values[0],
                            Email = values[1]
                        });
                    }
                }
            }
            return contacts;
        }

        private void DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Verificar se a célula clicada é da coluna de botões
            if (e.ColumnIndex == dataGridView.Columns["ActionButton"].Index && e.RowIndex >= 0)
            {
                string name = dataGridView.Rows[e.RowIndex].Cells["Name"].Value.ToString();
                string email = dataGridView.Rows[e.RowIndex].Cells["Email"].Value.ToString();
                Form1.Instancia_contactos.adicionar_bcc_ficheiro(email);
            }
        }

        private void contactos_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string searchValue = textBox1.Text.Trim().ToLower();

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                // Ignora a nova linha usada para inserção
                if (!row.IsNewRow)
                {
                    // Verifica se a célula da coluna "Name" contém o texto da pesquisa
                    if (row.Cells["Name"].Value != null &&
                        row.Cells["Name"].Value.ToString().ToLower().Contains(searchValue))
                    {
                        row.Visible = true; // Mostra a linha se corresponder
                    }
                    else
                    {
                        row.Visible = false; // Oculta a linha se não corresponder
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
    public class Contact
    {
        public string Name { get; set; }
        public string Email { get; set; }
    }
}
