using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using System.Web;
using System.Net;
using Microsoft.Web.WebView2.Core;

namespace Softnova_Emails
{
    public partial class BrowserForm : Form
    {
        private WebView2 webView;

        public string AuthorizationCode { get; private set; }

        public BrowserForm()
        {
            InitializeComponent();
        }

        private void BrowserForm_Load(object sender, EventArgs e)
        {

        }

        public BrowserForm(string url)
        {
            // Configura a janela 
            this.Text = "Softnova";
            this.Width = 800;
            this.Height = 600;
            this.Icon = Properties.Resources.NOVA;


            // Inicializa o WebView2
            webView = new WebView2
            {
                Dock = DockStyle.Fill
            };

            this.Controls.Add(webView);

            // Carrega a URL inicial ao abrir o formulário
            this.Load += async (sender, e) =>
            {
                await webView.EnsureCoreWebView2Async(null);
                webView.Source = new Uri(url);

                // Associa o evento de navegação
                webView.NavigationStarting += WebView_NavigationStarting;
            };
        }

        private void WebView_NavigationStarting(object sender, CoreWebView2NavigationStartingEventArgs e)
        {
            var url = e.Uri;

            // Verifica se a URL contém o código de autorização
            if (url.Contains("code="))
            {
                // Extrai o código de autorização
                AuthorizationCode = GetQueryParameter(url, "code");
                // Cancela a navegação e fecha a janela
                e.Cancel = true;
                this.Close();
            }
            else if (url.Contains("error="))
            {
                // Lida com erros de autenticação
                string error = GetQueryParameter(url, "error");
                MessageBox.Show($"Erro na autenticação: {error}");
                // Cancela a navegação e fecha a janela
                e.Cancel = true;
                this.Close();
            }
        }

        // Método para extrair parâmetros de uma URL sem HttpUtility
        private string GetQueryParameter(string url, string parameterName)
        {
            var uri = new Uri(url);
            var queryParams = uri.Query.TrimStart('?')
                .Split('&')
                .Select(q => q.Split('='))
                .ToDictionary(kv => kv[0], kv => kv.Length > 1 ? kv[1] : string.Empty);

            queryParams.TryGetValue(parameterName, out string value);
            return WebUtility.UrlDecode(value); // Decodifica o valor do parâmetro
        }
    }
}
