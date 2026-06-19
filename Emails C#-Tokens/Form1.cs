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
using System.Net;
using System.Net.Sockets;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.InteropServices;
using EASendMail;
using SmtpClient = System.Net.Mail.SmtpClient;
using SmtpClient_ = EASendMail.SmtpClient;
using Attachment = System.Net.Mail.Attachment;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimePart = MimeKit.MimePart;
using RtfPipe;
using Font = System.Drawing.Font;
using MailAddress = System.Net.Mail.MailAddress;
using System.Net.Http;

namespace Softnova_Emails
{
    public partial class Form1 : Form
    {
        public static Form1 Instancia_contactos { get; private set; }
        List<string> anexos = new List<string>
        {
        }; // Lista de caminhos dos arquivos que deseja anexar
        private String[] args = Environment.GetCommandLineArgs();
        private static string token = string.Empty;
        public string caminhoDoArquivoAnexo = "";
        string smtp, mail_origem, mail_destino, password, assunto, corpo, recibo_leitura, template, anexo1, anexo2, anexo3,
            anexo4, anexo5, anexo6, mail_origem_mostrar, bcc, login_user, abre_janela, clientID, clientSecret, outlook, template_D, Google;
        bool ssl;
        int porta;
                                            //Outlook
        const string tokenUri = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        const string scope = "https://outlook.office.com/SMTP.Send%20offline_access%20email%20openid";
        const string authUri = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";

                                            //Google
        const string tokenUri_Google = "https://oauth2.googleapis.com/token";
        const string scope_Google = "https://mail.google.com/ https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email";
        const string authUri_Google = "https://accounts.google.com/o/oauth2/auth";

        public Form1()
        {
            InitializeComponent();
            Instancia_contactos = this;

            // Definições da janela
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            //
             
            // Argumentos de entrada
            smtp = args[1]; 
            porta = Convert.ToInt32(args[2]); 
            mail_origem = args[4]; 
            mail_destino = args[3]; 
            password = args[5];
            ssl = true;
            assunto = args[7];
            corpo = args[8].Replace(" LN ", "\n");
            recibo_leitura = args[9];
            if (!string.IsNullOrEmpty(args[10]))
            {
                template = File.ReadAllText(args[10]);
            }
            template_D = @"<!DOCTYPE html><html><body><br><font style=""font-size: 9pt"" face=""Arial""> @@@MSG_BODY </font><br><br></body></html> ";
            anexo1 = args[11];
            anexo2 = args[12];
            anexo3 = args[13];
            mail_origem_mostrar = args[14];
            bcc = args[15];
            if (bcc == "NULL")
                bcc = "";
            anexo4 = args[16];
            anexo5 = args[17];
            anexo6 = args[18];
            login_user = args[19];
            abre_janela = args[20];
            outlook = args[21];
            clientID = args[22]; // ClientID
            clientSecret = args[23]; // clientSecret
            Google = "S";
            //

            // Passar os argumentos para as textbox
            textBox1.Text = mail_origem; // Email Remetente
            textBox2.Text = mail_destino; // Email Receptor
            textBox3.Text = bcc; // BCC
            if (!string.IsNullOrEmpty(args[11]))
            {
                textBox4.Text = args[11]; // Anexos Base de dados
            }
            if (!string.IsNullOrEmpty(args[12]))
            {
                textBox4.Text += ";" + args[12]; // Anexos Base de dados
            }
            if (!string.IsNullOrEmpty(args[13]))
            {
                textBox4.Text += ";" + args[13]; // Anexos Base de dados
            }
            if (!string.IsNullOrEmpty(args[16]))
            {
                textBox4.Text += ";" + args[16]; // Anexos Base de dados
            }
            if (!string.IsNullOrEmpty(args[17]))
            {
                textBox4.Text += ";" + args[17]; // Anexos Base de dados
            }
            if (!string.IsNullOrEmpty(args[18]))
            {
                textBox4.Text += ";" + args[18]; // Anexos Base de dados
            }
            if (!string.IsNullOrEmpty(args[7]))
            {
                textBox5.Text = assunto; // Assunto
            }
            if (!string.IsNullOrEmpty(args[8]))
            {
                richTextBox1.Text = corpo; // Mensagem
            }

            if (abre_janela == "N")
            {
                if (outlook == "MICROSOFT")
                {
                    this.Shown += (sender, e) => this.Hide();

                    // Executar a função
                    //Fazer_Login();
                    Fazer_Login_GOOGLE_Microsoft("Token.txt");
                    // Fechar o programa após executar a função
                    Environment.Exit(0);
                    Application.Exit();
                }else if (outlook == "GOOGLE")
                {
                    this.Shown += (sender, e) => this.Hide();
                    Fazer_Login_GOOGLE_Microsoft("Token_Google.txt");
                    Environment.Exit(0);
                    Application.Exit();
                }
                else
                {
                    this.Shown += (sender, e) => this.Hide();
                    var Destino = textBox2.Text.Split(';');
                    foreach (var bcc_ in Destino)
                    {
                        var trimmedAddress = bcc_.Trim(); // Remove espaços extras
                        if (!string.IsNullOrEmpty(trimmedAddress))
                        {
                            SendEmail_TLS(smtp, porta, trimmedAddress, mail_origem, password, assunto, corpo, ssl, mail_origem_mostrar, bcc, login_user);

                        }
                    }
                    // Fechar o programa após executar a função
                    Environment.Exit(0);
                    Application.Exit();
                }
                //Environment.Exit(0); // Encerra o programa com o código de saída 0 (sucesso)
            }
            // Preencher a ComboBox com os nomes das fontes disponíveis
            comboBox1.Items.AddRange(FontFamily.Families.Select(f => f.Name).ToArray());

            // Selecionar a primeira fonte como padrão (opcional)
            if (comboBox1.Items.Count > 0)
            {
                comboBox1.SelectedIndex = 0;
            }
            // Preencher a segunda ComboBox (comboBox2) com tamanhos de fonte
            comboBox2.Items.AddRange(new object[] { 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72 });
            comboBox2.SelectedIndex = 2; // Define 10 como tamanho padrão
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Adicionar fontes à ComboBox
            comboBox1.Items.Add("Arial");
            comboBox1.Items.Add("Times New Roman");
            comboBox1.Items.Add("Verdana");
            comboBox1.Items.Add("Tahoma");

            // Definir "Arial" como a seleção padrão
            comboBox1.SelectedItem = "Arial";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionLength > 0)
            {
                // Alinha o texto selecionado para a esquerda
                richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
            }
            else
            {
                MessageBox.Show("Selecione o texto que deseja alinhar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionLength > 0)
            {
                // Alinha o texto selecionado para a esquerda
                richTextBox1.SelectionAlignment = HorizontalAlignment.Right;
            }
            else
            {
                MessageBox.Show("Selecione o texto que deseja alinhar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionLength > 0)
            {
                // Alinha o texto selecionado para a esquerda
                richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
            }
            else
            {
                MessageBox.Show("Selecione o texto que deseja alinhar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (richTextBox1.SelectedText.Length > 0)
            {
                // Abrir o seletor de cores
                using (ColorDialog colorDialog = new ColorDialog())
                {
                    if (colorDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Alterar a cor de fundo do texto selecionado
                        richTextBox1.SelectionBackColor = colorDialog.Color;
                    }
                }
            }
            else
            {
                MessageBox.Show("Selecione um texto para aplicar o highlight.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            

        }

        private void button12_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionBullet = !richTextBox1.SelectionBullet;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(abre_janela != "N") { 
            // Verificar se há texto selecionado no RichTextBox
            if (richTextBox1.SelectedText.Length > 0)
            {
                // Obter o nome da fonte selecionada na ComboBox
                string fontName = comboBox1.SelectedItem.ToString();

                // Obter o tamanho atual da fonte ou um tamanho padrão
                float fontSize = richTextBox1.SelectionFont?.Size ?? richTextBox1.Font.Size;

                // Aplicar a nova fonte ao texto selecionado
                richTextBox1.SelectionFont = new Font(fontName, fontSize);
            }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(abre_janela != "N") { 
            // Alterar o tamanho da fonte com base na seleção na comboBox2
            if (richTextBox1.SelectedText.Length > 0)
            {
                string fontName = richTextBox1.SelectionFont?.FontFamily.Name ?? richTextBox1.Font.FontFamily.Name;
                float fontSize = Convert.ToSingle(comboBox2.SelectedItem);

                richTextBox1.SelectionFont = new Font(fontName, fontSize);
            }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionFont != null)
            {
                // Obtém a fonte atual
                Font currentFont = richTextBox1.SelectionFont;
                // Alterna entre negrito e estilo normal
                FontStyle newFontStyle = currentFont.Style ^ FontStyle.Bold;

                // Aplica a nova fonte ao texto selecionado
                richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
            }
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionFont != null)
            {
                Font currentFont = richTextBox1.SelectionFont;
                FontStyle newFontStyle = currentFont.Style ^ FontStyle.Underline;

                richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
            }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionLength > 0)
            {
                Font currentFont = richTextBox1.SelectionFont;

                // Verifica se o texto já está em itálico
                if (currentFont != null)
                {
                    FontStyle newStyle = currentFont.Style;

                    if (currentFont.Italic)
                    {
                        // Remove o estilo itálico
                        newStyle &= ~FontStyle.Italic;
                    }
                    else
                    {
                        // Adiciona o estilo itálico
                        newStyle |= FontStyle.Italic;
                    }

                    // Aplica a nova fonte ao texto selecionado
                    richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newStyle);
                }
                else
                {
                    MessageBox.Show("O texto selecionado tem múltiplas fontes diferentes e não pode ser formatado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Selecione o texto que deseja formatar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os Arquivos (*.*)|*.*"; // Opcional: filtro para tipos de arquivo
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    caminhoDoArquivoAnexo = openFileDialog.FileName;
                    MessageBox.Show("Arquivo anexado: " + caminhoDoArquivoAnexo);
                    anexos.Add(caminhoDoArquivoAnexo);
                    richTextBox2.Text += caminhoDoArquivoAnexo + "\n";
                }
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionLength > 0)
            {
                // Alinha o texto selecionado para a esquerda
                richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
            }
            else
            {
                MessageBox.Show("Selecione o texto que deseja alinhar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionLength > 0)
            {
                // Alinha o texto selecionado para a esquerda
                richTextBox1.SelectionAlignment = HorizontalAlignment.Right;
            }
            else
            {
                MessageBox.Show("Selecione o texto que deseja alinhar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionLength > 0)
            {
                // Alinha o texto selecionado para a esquerda
                richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
            }
            else
            {
                MessageBox.Show("Selecione o texto que deseja alinhar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            richTextBox2.Clear();
            anexos.Clear();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            if (richTextBox1.SelectedText.Length > 0)
            {
                // Abrir o seletor de cores
                using (ColorDialog colorDialog = new ColorDialog())
                {
                    if (colorDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Alterar a cor de fundo do texto selecionado
                        richTextBox1.SelectionBackColor = colorDialog.Color;
                    }
                }
            }
            else
            {
                MessageBox.Show("Selecione um texto para aplicar o highlight.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            if (richTextBox1.SelectedText.Length > 0)
            {
                // Abrir o seletor de cores
                using (ColorDialog colorDialog = new ColorDialog())
                {
                    if (colorDialog.ShowDialog() == DialogResult.OK)
                    {
                        // Alterar a cor da fonte do texto selecionado
                        richTextBox1.SelectionColor = colorDialog.Color;
                    }
                }
            }
            else
            {
                MessageBox.Show("Selecione um texto para alterar a cor da fonte.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionLength > 0)
            {
                Font currentFont = richTextBox1.SelectionFont;

                // Verifica se o texto já está em itálico
                if (currentFont != null)
                {
                    FontStyle newStyle = currentFont.Style;

                    if (currentFont.Italic)
                    {
                        // Remove o estilo itálico
                        newStyle &= ~FontStyle.Italic;
                    }
                    else
                    {
                        // Adiciona o estilo itálico
                        newStyle |= FontStyle.Italic;
                    }

                    // Aplica a nova fonte ao texto selecionado
                    richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newStyle);
                }
                else
                {
                    MessageBox.Show("O texto selecionado tem múltiplas fontes diferentes e não pode ser formatado.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Selecione o texto que deseja formatar!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Btn_Contact_Click(object sender, EventArgs e)
        {
            contactos contactos = new contactos ();
            contactos.ShowDialog();
        }

        static int GetRandomUnusedPort()
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            return port;
        }

        public void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        static string carregar_R_token(string nome_ficheiro)
        {
            string Refresh_token = "";
            // Obtém o caminho da pasta onde o projeto está sendo executado
            //string caminhoPastaProjeto_ = AppDomain.CurrentDomain.BaseDirectory;
            string caminhoPastaProjeto_ = "c:/temp/emails";
            // Combina o caminho da pasta do projeto com o nome do arquivo
            string caminhoArquivo_ = Path.Combine(caminhoPastaProjeto_, nome_ficheiro);
            try
            {
                // Carrega o conteúdo do arquivo e armazena na variável "conteudo"
                Refresh_token = File.ReadAllText(caminhoArquivo_);
                // Exibe o conteúdo carregado
                Console.WriteLine("Conteúdo do arquivo:");
                Console.WriteLine(Refresh_token);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocorreu um erro ao carregar o arquivo: " + ex.Message);
            }
            return Refresh_token;
        }

        public async Task Enviar_Email(string accessToken, string email_r, string email_d, string bbc_, string assunto_, string corpo)
        {
            this.Cursor = Cursors.WaitCursor;
            // Cria a mensagem de e-mail
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(mail_origem_mostrar, email_r));
            message.To.Add(new MailboxAddress(mail_origem_mostrar, email_d));
            if (checkBox1.Checked || recibo_leitura == "S")
                message.Bcc.Add(new MailboxAddress(mail_origem_mostrar, email_r));

            if (!string.IsNullOrEmpty(bbc_))
            {
                // Supondo que os endereços estejam separados por vírgulas, por exemplo: "email1@example.com,email2@example.com"
                var bccAddresses = bbc_.Split(';');
                foreach (var bcc in bccAddresses)
                {
                    var trimmedAddress = bcc.Trim(); // Remove espaços extras
                    if (!string.IsNullOrEmpty(trimmedAddress))
                    {
                        message.Bcc.Add(new MailboxAddress(trimmedAddress, trimmedAddress));
                    }
                }
            }
            string htmlContent;
            message.Subject = assunto_;
            // Converte o RTF para HTML usando a biblioteca RtfPipe


            if (abre_janela == "N")
            {
                if (!string.IsNullOrEmpty(template))
                {
                    // Substituir placeholders
                    htmlContent = template
                       .Replace("@@@MSG_BODY", Rtf.ToHtml(richTextBox1.Rtf));
                }
                else
                {
                    htmlContent = Rtf.ToHtml(richTextBox1.Rtf);
                }
                // Configura o corpo do e-mail como HTML
                message.Body = new TextPart("html")
                {
                    Text = htmlContent
                };
            }
            else
            {
                if (!string.IsNullOrEmpty(template))
                {
                    // Substituir placeholders
                    htmlContent = template
                       .Replace("@@@MSG_BODY", Rtf.ToHtml(richTextBox1.Rtf));
                }
                else
                {
                    htmlContent = Rtf.ToHtml(richTextBox1.Rtf);
                }
                //string htmlContent = Rtf.ToHtml(richTextBox1.Rtf);
                // Configura o corpo do e-mail como HTML
                message.Body = new TextPart("html")
                {
                    Text = htmlContent
                };
            }

            if (anexo1 != "")
            {
                anexos.Add(anexo1);
            }
            if (anexo2 != "")
            {
                anexos.Add(anexo2);
            }
            if (anexo3 != "")
            {
                anexos.Add(anexo3);
            }
            if (anexo4 != "")
            {
                anexos.Add(anexo4);
            }
            if (anexo5 != "")
            {
                anexos.Add(anexo5);
            }
            if (anexo6 != "")
            {
                anexos.Add(anexo6);
            }

            // Se há anexos, adiciona ao corpo do e-mail
            if (anexos.Count() != 0)
            {
                var multipart = new Multipart("mixed");
                multipart.Add(message.Body);

                foreach (var filePath in anexos)
                {
                    if (File.Exists(filePath))
                    {
                        var mimeType = MimeTypes.GetMimeType(filePath);
                        var attachment = new MimePart(mimeType)
                        {
                            Content = new MimeContent(File.OpenRead(filePath)),
                            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                            ContentTransferEncoding = ContentEncoding.Base64,
                            FileName = Path.GetFileName(filePath)
                        };
                        multipart.Add(attachment);
                    }
                }
                message.Body = multipart;
            }

            // Configuração do cliente SMTP usando MailKit
            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                try
                {
                    client.Connect(smtp, porta, SecureSocketOptions.StartTls);
                    client.Authenticate(new SaslMechanismOAuth2(email_r, accessToken));

                    // Envia o e-mail
                    client.Send(message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Erro ao enviar o e-mail: " + ex.Message);
                    MessageBox.Show("Erro ao enviar o e-mail: " + ex.Message);
                    Console.WriteLine("A sessão expirou, aperte em enviar, e faça o login novamente!");
                    string caminho = "c:/temp/emails/Token.txt";
                    anexos.Clear();
                    File.WriteAllText(caminho, string.Empty);
                    caminho = "c:/temp/emails/Token_Google.txt";
                    File.WriteAllText(caminho, string.Empty);
                }
                finally
                {
                    client.Disconnect(true);
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private static readonly HttpClient client = new HttpClient();

                                                        //OAuth//
        async void Fazer_Login_GOOGLE_Microsoft(string ficheiro_token)
        {
            string Refresh_token = carregar_R_token(ficheiro_token);
            // Creates a redirect URI using an available port on the loopback address.
            string redirectUri = string.Format("http://127.0.0.1:{0}/", GetRandomUnusedPort());
            Console.WriteLine("redirect URI: " + redirectUri);

            if (Refresh_token == "")
            {
                string authorizationRequest;
                // Creates an HttpListener to listen for requests on that redirect URI.
                var http = new HttpListener();
                http.Prefixes.Add(redirectUri);
                Console.WriteLine("Listening ...");
                http.Start();
                if (ficheiro_token == "Token.txt") { 
                // Creates the OAuth 2.0 authorization request.
                authorizationRequest = string.Format("{0}?response_type=code&scope={1}&redirect_uri={2}&client_id={3}&prompt=login",
                    authUri,
                    scope,
                    Uri.EscapeDataString(redirectUri),
                    clientID
                );
                }
                else
                {
                    authorizationRequest = string.Format("{0}?response_type=code&scope={1}&redirect_uri={2}&client_id={3}&prompt=login",
                        authUri_Google,
                        scope_Google,
                        Uri.EscapeDataString("http://localhost:5000/callback"),
                        clientID
                    );
                }
                BrowserForm browserForm = new BrowserForm(authorizationRequest);
                browserForm.ShowDialog();

                // extracts the code
                var code = Convert.ToString(browserForm.AuthorizationCode);
                Console.WriteLine("Authorization code: " + code);

                // MARISA - chama funçao TOKEN
                string responseText = await RequestAccessToken(code, redirectUri, ficheiro_token);
                Console.WriteLine(responseText);
                OAuthResponseParser parser = new OAuthResponseParser();
                parser.Load(responseText);
                //MessageBox.Show(browserForm.AuthorizationCode);

                var user = parser.EmailInIdToken;
                var accessToken = parser.AccessToken;

                // Obtém o caminho da pasta onde o projeto está sendo executado
                string caminhoPastaProjeto = "c:/temp/emails";

                // Nome do arquivo que deseja salvar
                string nomeArquivo = ficheiro_token;

                // Combina o caminho da pasta do projeto com o nome do arquivo
                string caminhoArquivo = Path.Combine(caminhoPastaProjeto, nomeArquivo);
                try
                {
                    // Salva o conteúdo no arquivo
                    File.WriteAllText(caminhoArquivo, parser.RefreshToken);
                    Console.WriteLine("O texto foi salvo com sucesso no arquivo!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ocorreu um erro ao salvar o arquivo: " + ex.Message);
                }
                Console.WriteLine("AccessToken: {0}", accessToken);
                try
                {
                    if (abre_janela == "N")
                    {
                        var Destino = args[4].Split(';');
                        foreach (var bcc in Destino)
                        {
                            var trimmedAddress = bcc.Trim(); // Remove espaços extras
                            if (!string.IsNullOrEmpty(trimmedAddress))
                            {
                                await Enviar_Email(accessToken, mail_origem, trimmedAddress, bcc, assunto, corpo);

                            }
                        }
                        Environment.Exit(0);
                        this.Close();
                    }
                    else
                    {
                        var Destino = textBox2.Text.Split(';');
                        foreach (var bcc in Destino)
                        {
                            var trimmedAddress = bcc.Trim(); // Remove espaços extras
                            if (!string.IsNullOrEmpty(trimmedAddress))
                            {
                                await Enviar_Email(accessToken, textBox1.Text, trimmedAddress, textBox3.Text, textBox5.Text, richTextBox1.Text);
                            }
                        }
                        if (abre_janela == "N")
                        {
                            Console.WriteLine("E-mail enviado com sucesso!");
                        }
                        else
                        {
                            MessageBox.Show("E-mail enviado com sucesso!");
                        }
                        Environment.Exit(0);
                        this.Close();
                    }
                }catch(Exception ex)
                {
                    Console.WriteLine(ex);
                    string caminho = "c:/temp/emails/" + ficheiro_token;
                    File.WriteAllText(caminho, string.Empty);
                    MessageBox.Show("Erro ao enviar o email, tente novamente.");

                }
            }
            else
            {
                try { 
                // MARISA - chama funçao TOKEN
                string responseText = await RequestAccessToken("", redirectUri, ficheiro_token);
                Console.WriteLine(responseText);
                OAuthResponseParser parser = new OAuthResponseParser();
                parser.Load(responseText);

                var user = parser.EmailInIdToken;
                var accessToken_ = parser.AccessToken;
                if (abre_janela == "N")
                {
                        var Destino = args[4].Split(';');
                        foreach (var bcc_ in Destino)
                        {
                            var trimmedAddress = bcc_.Trim(); // Remove espaços extras
                            if (!string.IsNullOrEmpty(trimmedAddress))
                            {
                                await Enviar_Email(accessToken_, mail_origem, trimmedAddress, bcc, assunto, corpo);

                            }
                        }
                        Environment.Exit(0);
                        this.Close();
                }
                else
                {
                        var Destino = textBox2.Text.Split(';');
                        foreach (var bcc_ in Destino)
                        {
                            var trimmedAddress = bcc_.Trim(); // Remove espaços extras
                            if (!string.IsNullOrEmpty(trimmedAddress))
                            {
                                await Enviar_Email(accessToken_, textBox1.Text, trimmedAddress, textBox3.Text, textBox5.Text, richTextBox1.Text);

                            }
                        }

                    }
                    if (abre_janela == "N")
                    {
                        Console.WriteLine("E-mail enviado com sucesso!");
                    }
                    else
                    {
                        MessageBox.Show("E-mail enviado com sucesso!");
                    }
                    Environment.Exit(0);
                    this.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    string caminho = "c:/temp/emails/" + ficheiro_token;
                    File.WriteAllText(caminho, string.Empty);
                    MessageBox.Show("Erro ao enviar o email, tente novamente.");
                }
            }
        }
                                                    
        /*
        async void Fazer_Login()
        {
            string Refresh_token = carregar_R_token("Token.txt");
            // Creates a redirect URI using an available port on the loopback address.
            string redirectUri = string.Format("http://127.0.0.1:{0}/", GetRandomUnusedPort());
            Console.WriteLine("redirect URI: " + redirectUri);

            if (Refresh_token == "")
            {
                // Creates an HttpListener to listen for requests on that redirect URI.
                var http = new HttpListener();
                http.Prefixes.Add(redirectUri);
                Console.WriteLine("Listening ...");
                http.Start();

                // Creates the OAuth 2.0 authorization request.
                string authorizationRequest = string.Format("{0}?response_type=code&scope={1}&redirect_uri={2}&client_id={3}&prompt=login",
                    authUri,
                    scope,
                    Uri.EscapeDataString(redirectUri),
                    clientID
                );

                BrowserForm browserForm = new BrowserForm(authorizationRequest);
                browserForm.ShowDialog();
                
                // extracts the code
                var code = Convert.ToString(browserForm.AuthorizationCode);
                Console.WriteLine("Authorization code: " + code);

                // MARISA - chama funçao TOKEN
                string responseText = await RequestAccessToken(code, redirectUri);
                Console.WriteLine(responseText);
                OAuthResponseParser parser = new OAuthResponseParser();
                parser.Load(responseText);
                //MessageBox.Show(browserForm.AuthorizationCode);

                var user = parser.EmailInIdToken;
                var accessToken = parser.AccessToken;

                // Obtém o caminho da pasta onde o projeto está sendo executado
                string caminhoPastaProjeto = "c:/temp/emails";

                // Nome do arquivo que deseja salvar
                string nomeArquivo = "Token.txt";

                // Combina o caminho da pasta do projeto com o nome do arquivo
                string caminhoArquivo = Path.Combine(caminhoPastaProjeto, nomeArquivo);
                try
                {
                    // Salva o conteúdo no arquivo
                    File.WriteAllText(caminhoArquivo, parser.RefreshToken);
                    Console.WriteLine("O texto foi salvo com sucesso no arquivo!");
                }catch (Exception ex)
                {
                    Console.WriteLine("Ocorreu um erro ao salvar o arquivo: " + ex.Message);
                }
                Console.WriteLine("AccessToken: {0}", accessToken);
                if(abre_janela == "N") {
                    await Enviar_Email(accessToken, mail_origem, mail_destino, bcc, assunto, corpo);
                }
                else { 
                await Enviar_Email(accessToken, textBox1.Text, textBox2.Text, textBox3.Text, textBox5.Text, richTextBox1.Text);
                }
            }
            else
            {
                 // MARISA - chama funçao TOKEN
                string responseText = await RequestAccessToken("", redirectUri);
                Console.WriteLine(responseText);
                OAuthResponseParser parser = new OAuthResponseParser();
                parser.Load(responseText);

                var user = parser.EmailInIdToken;
                var accessToken_ = parser.AccessToken;
                if(abre_janela == "N") {
                    await Enviar_Email(accessToken_, mail_origem, mail_destino, bcc, assunto, corpo);
                }
                else { 
                await Enviar_Email(accessToken_, textBox1.Text, textBox2.Text, textBox3.Text, textBox5.Text, richTextBox1.Text);
                }
            }



        }
        */
        async Task<string> RequestAccessToken(string code, string redirectUri, string ficheiro_token)
        {
            string Refresh_token = carregar_R_token(ficheiro_token);
            Console.WriteLine("Exchanging code for tokens...");
            string tokenRequestBody = "";
            HttpWebRequest tokenRequest = (HttpWebRequest)WebRequest.Create(tokenUri);
            
            if (ficheiro_token == "Token.txt") { 
            if (Refresh_token == "")
            {
                tokenRequestBody = string.Format("code={0}&redirect_uri={1}&client_id={2}&grant_type=authorization_code",
                code,
                Uri.EscapeDataString(redirectUri),
                clientID
                );
            }
            else
            {
                //   REFRESH TOKEN
                tokenRequestBody = string.Format("client_id={0}&refresh_token={1}&grant_type=refresh_token",
                    clientID,
                    //clientSecret,
                    Refresh_token
                    );
            }
            }
            else
            {
                tokenRequest = (HttpWebRequest)WebRequest.Create(tokenUri_Google);
                if (Refresh_token == "")
                {
                    tokenRequestBody = string.Format("code={0}&redirect_uri={1}&client_id={2}&client_secret={3}&grant_type=authorization_code",
                    code,
                    Uri.EscapeDataString("http://localhost:5000/callback"),
                    clientID,
                    clientSecret
                    );
                }
                else
                {
                    tokenRequestBody = string.Format("client_id={0}&client_secret={1}&refresh_token={2}&grant_type=refresh_token",
                        clientID,
                        clientSecret,
                        Refresh_token
                );
                }
            }
            // sends the request
            tokenRequest.Method = "POST";
            tokenRequest.ContentType = "application/x-www-form-urlencoded";
            tokenRequest.Accept = "Accept=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";

            byte[] _byteVersion = Encoding.ASCII.GetBytes(tokenRequestBody);
            tokenRequest.ContentLength = _byteVersion.Length;

            Stream stream = tokenRequest.GetRequestStream();
            await stream.WriteAsync(_byteVersion, 0, _byteVersion.Length);
            stream.Close();

            try
            {
                // gets the response
                WebResponse tokenResponse = await tokenRequest.GetResponseAsync();
                using (StreamReader reader = new StreamReader(tokenResponse.GetResponseStream()))
                {
                    // reads response body
                    return await reader.ReadToEndAsync();
                }
            }catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError)
                {
                    var response = ex.Response as HttpWebResponse;
                    if (response != null)
                    {
                        Console.WriteLine("HTTP: " + response.StatusCode);
                        using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                        {
                            // reads response body
                            string responseText = await reader.ReadToEndAsync();
                            Console.WriteLine(responseText);
                        }
                    }
                }
                throw ex;
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        bool IsEmailValid(string email)
        {
            try
            {
                var mailAddress = new MailAddress(email);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(outlook == "MICROSOFT")
            {
                Fazer_Login_GOOGLE_Microsoft("Token.txt");
            }
            else if (outlook == "GOOGLE")
            {
                Fazer_Login_GOOGLE_Microsoft("Token_Google.txt");
            }
            else
            {
                var Destino = textBox2.Text.Split(';');
                try { 
                foreach (var bcc_ in Destino)
                {
                    var trimmedAddress = bcc_.Trim(); // Remove espaços extras
                    if (!string.IsNullOrEmpty(trimmedAddress))
                    {
                        SendEmail_TLS(smtp, porta, trimmedAddress, textBox1.Text, password, textBox5.Text, richTextBox1.Text, ssl, mail_origem_mostrar, textBox3.Text, login_user);
                    }
                }
                MessageBox.Show("E-mail enviado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) {
                    MessageBox.Show("Erro ao enviar e-mail: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                Environment.Exit(0);
                this.Close();
            }
        }

        public void adicionar_bcc_ficheiro(string email)
        {
            if (textBox3.Text == "") { 
            textBox3.Text += "" + email;
            }
            else
            {
               textBox3.Text += ";" + email;
            }
        }

            [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public void BringConsoleToFront()
        {
            SetForegroundWindow(GetConsoleWindow());
        }

        public void SendEmail_TLS(string smtp, int porta, string emailDestino, string emailEmissor, string senha, string assunto, string mensagem, bool enabledSSL, string mailOrigemMostrar, string bccGeral, string login)
        {
            this.Cursor = Cursors.WaitCursor;
            try
            {

                // Inicializa o objeto SmtpMail com a chave de licença
                SmtpMail oMail = new SmtpMail("ES-D1508812687-00760-3DBEA339BE883E9B-2E8D6T51F5BD6C2E");
                SmtpClient_ oSmtp = new SmtpClient_();

                // Define o remetente diretamente como string
                oMail.From = emailEmissor;

                // Configura o destinatário
                oMail.To = emailDestino;

                if (checkBox1.Checked || recibo_leitura == "S")
                    oMail.Bcc += emailEmissor;

                // Adiciona BCC, se fornecido
                if (!string.IsNullOrEmpty(bccGeral))
                {
                    // Supondo que os endereços estejam separados por vírgulas, por exemplo: "email1@example.com,email2@example.com"
                    var bccAddresses = bccGeral.Split(';');
                    foreach (var bcc in bccAddresses)
                    {
                        var trimmedAddress = bcc.Trim(); // Remove espaços extras
                        if (!string.IsNullOrEmpty(trimmedAddress))
                        {
                            oMail.Bcc += ";" + trimmedAddress;
                        }
                    }
                }

                // Define o assunto do e-mail
                oMail.Subject = assunto;

                // Define o corpo do e-mail como HTML
                if (abre_janela == "N")
                {
                    if (!string.IsNullOrEmpty(template))
                    {
                        // Substituir placeholders
                        oMail.HtmlBody = template
                           .Replace("@@@MSG_BODY", corpo);
                    }
                    else
                    {
                        oMail.HtmlBody = corpo;
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(template))
                    {
                        // Substituir placeholders
                        oMail.HtmlBody = template
                           .Replace("@@@MSG_BODY", Rtf.ToHtml(richTextBox1.Rtf));
                    }
                    else
                    {
                        oMail.HtmlBody = Rtf.ToHtml(richTextBox1.Rtf);
                    }
                }

                if (anexo1 != "")
                {
                    anexos.Add(anexo1);
                }
                if (anexo2 != "")
                {
                    anexos.Add(anexo2);
                }
                if (anexo3 != "")
                {
                    anexos.Add(anexo3);
                }
                if (anexo4 != "")
                {
                    anexos.Add(anexo4);
                }
                if (anexo5 != "")
                {
                    anexos.Add(anexo5);
                }
                if (anexo6 != "")
                {
                    anexos.Add(anexo6);
                }
                if (anexos.Count() != 0)
                {
                    var multipart = new Multipart("mixed");

                    foreach (var filePath in anexos)
                    {
                        oMail.AddAttachment(Convert.ToString(filePath));
                    }
                }
                if (login == "")
                    login = emailEmissor;
                // Configura o servidor SMTP
                SmtpServer oServer = new SmtpServer(smtp)
                {
                    User = login,
                    Password = senha,
                    Port = porta,
                    ConnectType = enabledSSL ? SmtpConnectType.ConnectSSLAuto : SmtpConnectType.ConnectDirectSSL,
                    AuthType = SmtpAuthType.AuthAuto
                };

                // Envia o e-mail
                oSmtp.SendMail(oServer, oMail);

            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                // Fecha o formulário de carregamento
                this.Cursor = Cursors.Default;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionFont != null)
            {
                // Obtém a fonte atual
                Font currentFont = richTextBox1.SelectionFont;
                // Alterna entre negrito e estilo normal
                FontStyle newFontStyle = currentFont.Style ^ FontStyle.Bold;

                // Aplica a nova fonte ao texto selecionado
                richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // Verifica se há texto selecionado
            if (richTextBox1.SelectionFont != null)
            {
                Font currentFont = richTextBox1.SelectionFont;
                FontStyle newFontStyle = currentFont.Style ^ FontStyle.Underline;

                richTextBox1.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Todos os Arquivos (*.*)|*.*"; // Opcional: filtro para tipos de arquivo
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    caminhoDoArquivoAnexo = openFileDialog.FileName;
                    MessageBox.Show("Arquivo anexado: " + caminhoDoArquivoAnexo);
                    anexos.Add(caminhoDoArquivoAnexo);
                    richTextBox2.Text += caminhoDoArquivoAnexo + "\n";
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            richTextBox2.Clear();
            anexos.Clear();
        }
    }
}
