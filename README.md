### 📧Projeto envio de Emails

  O presente projeto é destinado ao envio de emails, quer sejam eles Outlook ou Gmail, e qualquer outra plataforma. O programa permite adicionar BCC, ficheiros em anexo, assunto, bem como negrito, sublinhado, italico, alinhar para o centro, esquerda e direita, escolher a cor do texto e do sublinhado. O ficheiro executavel também pode ser executado por meio de argumentos de entrada de forma a facilitar o processo.


### Algumas funcionalidades a destacar
Ficheiro-> Emails C#-Tokens-> Form1.cs <br/>

***Funções***

- Carregar_R_Token: A função destina-se a buscar o refresh token armazenado na memoria do computador, de forma a ser possivel atualizar o token. <br/>
- Fazer_Login_GOOGLE_Microsoft: A função tem como a variavel de entrada "ficheiro_token", e destina se a fazer a atualização do token por meio do Refresh token, ou fazer a autenticação dos usuarios por meio de Email e Password. Quando a função é iniciada ela tenta verificar se existe algum ficheiro com token por meio da função "Carregar_R_Token", caso não existir é feito o Login por meio do Email e da Password, e caso existir token é feito a atualização do token.<br/>
- Enviar_Email: A função destina se a enviar os emails do Outlook e da Google, ela conta com variaveis de entrarda accessToken(contem o accessToken), email_r(o email de remetente), email_d(email de destino), bbc_(o bbc do mail), assunto_(assunto do mail), e corpo(a mensagem a ser enviada). No final ao ser enviado o email foi usada a biblioteca MailKit onde foi possivel usar o SmtpClient, primeiro é feito a autenticação com o email e o accesstoken, e no final é enviado o email<br/>


***Variáveis***

- TokenUri: É uma constante que contem o link destinado a obter o token de autenticação do Outlook. <br/>
- Scope: É uma contante que contem o link de scope do Outlook.<br/>
- AuthUri: É uma constante que contem o link destinado a fazer a autenticação de emails Outlook, por meio de Email e Senha. <br/>

- TokenUri_Google: É uma constante que contem o link destinado a obter o token de autenticação de emails Gmail. <br/>
- Scope_Google: É uma contante que contem o link de scope do Gmail.<br/>
- AuthUri_Google: É uma constante que contem o link destinado a fazer a autenticação de emails Gmail, por meio de Email e Senha. <br/>

***Links***
- OAuth 2.0 para API Google <br/>
https://developers.google.com/identity/protocols/oauth2?hl=pt-br

- OAuth 2.0 para API Outlook <br/>
  https://learn.microsoft.com/pt-pt/advertising/guides/authentication-oauth-get-tokens?view=bingads-13
