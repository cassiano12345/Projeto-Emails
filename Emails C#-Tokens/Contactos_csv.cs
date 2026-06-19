using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Data.OleDb;
using System.IO;

namespace Softnova_Emails
{
    public class Contact
    {
        public string Name { get; set; }
        public string Email { get; set; }
    }
    public partial class Contactos_csv : Form
    {
        string valor = "";
        public Contactos_csv()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            dataGridView1.Columns.Add("Name", "Name");
            dataGridView1.Columns.Add("Email", "Email");
            var buttonColumn = new DataGridViewButtonColumn
            {
                HeaderText = "Adicionar",
                Text = "Selecionar",
                UseColumnTextForButtonValue = true
            };
            dataGridView1.Columns.Add(buttonColumn);
            dataGridView1.CellContentClick += DataGridView_CellContentClick;
            LoadContacts();
        }

        private void Contactos_csv_Load(object sender, EventArgs e)
        {
        }
        private void LoadContacts()
        {
            // Simulando a leitura de contatos de um arquivo CSV
            string filePath = "C:/temp/EMAILS/contactos.csv";
            var contacts = ReadContactsFromCsv(filePath);

            // Adicionar os contatos ao DataGridView
            foreach (var contact in contacts)
            {
                dataGridView1.Rows.Add(contact.Name, contact.Email);
            }
        }

        private void DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Verificar se a célula clicada é da coluna de botões
            if (e.ColumnIndex == dataGridView1.Columns["Adicionar"].Index && e.RowIndex >= 0)
            {
                string name = dataGridView1.Rows[e.RowIndex].Cells["Name"].Value.ToString();
                string email = dataGridView1.Rows[e.RowIndex].Cells["Email"].Value.ToString();

                // Exibir os detalhes em uma MessageBox
                MessageBox.Show($"Nome: {name}\nE-mail: {email}", "Detalhes do Contato");
            }
        }

        public static List<Contact> ReadContactsFromCsv(string caminho_ficheiro)
        {
            var contacts = new List<Contact>();

            // Abrir e ler o arquivo linha por linha
            using (var reader = new StreamReader(caminho_ficheiro))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var values = line.Split(';');

                    // Verificar se a linha contém pelo menos dois valores
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
    }
}
