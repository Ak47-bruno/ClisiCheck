using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Linq;
using System.Net;
using OSVersionExtension;
using WUApiLib;
//TO DO
// Verificar versionanemento MCAfee, SAP
// Outras versões Loja, Gerência, 
// Instalação dos programas.
namespace ClisiCheck
{
    //----------------------------Código para o botao checkar ficar redondo---------------------------------------
    public partial class start : Form
    {
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]

        private static extern IntPtr CreateRoundRectRgn
  (
      int nLeftRect,
      int nTopRect,
      int nRightRect,
      int nBottomRect,
      int nWidthEllipse,
      int nHeightEllipse
  );
        //-------------------------Código para o botao checkar ficar redondo---------------------------------------


        private bool dragging = false;
        private Point startPoint = new Point(0, 0);

        public start()
        {
            InitializeComponent();
            progressBar.Value = 0;
           // Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 25, 25));
            
        }

        public void userLabel(string userName)
        {
            lblUserName.Text = "Logado com: " + userName;
        }

        private void start_Load(object sender, EventArgs e)
        {

        }

        private void btnDash_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnDash.Height;
            pnlNav.Top = btnDash.Top;
            pnlNav.Left = btnDash.Left;
            lblModo.Text = "Escritório";
            //listViewResult.Controls.Clear();
            progressBar.Value = 0;
            progressBar.Text = progressBar.Value.ToString() + "%";
            listBoxResult.Items.Clear();
            btnDash.BackColor = Color.FromArgb(46, 51, 73);
        }
               
        private void btnCaixa_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnCaixa.Height;
            pnlNav.Top = btnCaixa.Top;
            lblModo.Text = "Caixa";
            //listViewResult.Controls.Clear();
            progressBar.Value = 0;
            progressBar.Text = progressBar.Value.ToString() + "%";
            listBoxResult.Items.Clear();
            btnCaixa.BackColor = Color.FromArgb(46, 51, 73);
        }

        private void btnGer_Click(object sender, EventArgs e)
        {
            pnlNav.Height = btnGer.Height;
            pnlNav.Top = btnGer.Top;
            lblModo.Text = "Gerência";
            //listViewResult.Controls.Clear();
            progressBar.Value = 0;
            progressBar.Text = progressBar.Value.ToString() + "%";
            listBoxResult.Items.Clear();
            btnGer.BackColor = Color.FromArgb(46, 51, 73);
        }

        private void btnDash_Leave(object sender, EventArgs e)
        {
            btnDash.BackColor = Color.White;
        }

        private void btnCaixa_Leave(object sender, EventArgs e)
        {
            btnCaixa.BackColor = Color.White;
        }

        private void btnGer_Leave(object sender, EventArgs e)
        {
            btnGer.BackColor = Color.White;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.ExitThread();
        }

        private void btnChek_Click(object sender, EventArgs e)
        {
            btnChek.PointToScreen(panel9.Location);
            listBoxResult.Controls.Clear();
            listBoxResult.Items.Clear();
            logsClear();
            progressBar.Value = 0;
            progressBar.Text = progressBar.Value.ToString() + "%";

            //Lista de Programas
            List<string> listProgram = new List<string>();
            switch (lblModo.Text)
            {
                case "Escritório":
                    listProgram = new List<string> { "7-Zip", "Carsybde", "Google Chrome", "CutePDF", "Google Drive", "Java", "Acrobat Reader DC", "FortiClient VPN", "Dell Command"};
                    escritorioCheck(listProgram);
                    checkPasta();
                    editionWindows();
                    UpdatesAvailable();
                    break;
                case "Caixa":
                    listProgram = new List<string> { "7-Zip", "Carsybde", "Google Chrome", "Google Drive", "Ivanti Endpoint", "Gertec", "Dell Command" };                    
                    escritorioCheck(listProgram);
                    impressoraCheck();
                    checkPasta();
                    atalhosCheck();
                    editionWindows();
                    UpdatesAvailable();
                    break;
                case "Gerência":
                    listProgram = new List<string> { "7-Zip", "Carsybde", "Google Chrome", "CutePDF", "Google Drive", "Java", "Acrobat Reader DC", "FortiClient VPN", "Digifort Enterprise", "Gertec", "Dell Command" };
                    escritorioCheck(listProgram);
                    impressoraCheck();
                    checkPasta();
                    atalhosCheck();
                    editionWindows();
                    UpdatesAvailable();
                    break;
            }
        }

        public void escritorioCheck(List<String> programas)
        {
           
            string displayName;
            string displayVersion;
            //Lista de todos os programas instalados
            List<string> list = new List<string>();
            //Lista de programas que contem versão
            List<string> listVersion = new List<string>();            
            RegistryKey key;
            switch (lblModo.Text)
            {
                case "Escritório":
                   // listProgram = new List<string> { "7-Zip", "Carsybde", "Google Chrome", "CutePDF", "Google Drive", "Java", "Acrobat Reader DC", "FortiClient VPN"};
                    break;
                case "Caixa":
                    //listProgram = new List<string> { "7-Zip", "Carsybde", "Google Chrome", "Google Drive", "Ivanti Endpoint" };
                    break;
                case "Gerência":
                    break;
            }

            //Verifica todos os softwares no computador
            List<String> source = new List<string> { "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" };
            foreach (var s in source)
            {
                key = Registry.LocalMachine.OpenSubKey($@"{s}");
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey subkey = key.OpenSubKey(keyName);
                    displayName = subkey.GetValue("DisplayName") as string;
                    displayVersion = subkey.GetValue("DisplayVersion") as string;
                    if (displayName != null)
                    {
                        if (displayVersion == null)
                        {
                            list.Add(displayName);
                            //listBoxResult.Items.Add(displayName);
                        }
                        else
                        {
                            list.Add(displayName + " Versão " + displayVersion);
                            listVersion.Add(displayName + " Versão " + displayVersion);
                            //listBoxResult.Items.Add(displayName + " Versão " + displayVersion);
                        }

                    }
                }
            }


            //Verifica se está no domínio Bemol
            hostName();

            //Verifica se os softwares ESPECÍFICOS estão instalados
            for (int g = 0; g < programas.Count; g++)
            {
                if (!list.Exists(e => e.Contains(programas[g])))
                {
                    listBoxResult.Items.Add(programas[g]);
                }
            }
            
            if (!list.Exists(e => e.Contains("SAP GUI for Windows")))
            {
                listBoxResult.Items.Add("SAP GUI for Windows");

            }
            else
            {
                if (!list.Exists(e => e.Contains("SAP GUI for Windows 7.70  (Patch 5) Versão 7.70 Compilation 1")))
                {
                    listBoxResult.Items.Add("SAP GUI Desatualizado");
                }
            }

            if (!list.Exists(e => e.Contains("McAfee Agent")))
            {
                listBoxResult.Items.Add("McAfee Agente");
            }
            else
            {
                if (!list.Exists(e => e.Contains("McAfee Agent Versão 5.7.6.251")))
                {
                    listBoxResult.Items.Add("McAfee Agent Desatualizado");
                }
            }          

            
                       
        }

        //Método para checar atalhos
        public void atalhosCheck()
        {

            //path = "C:\\Users\\Public\\Desktop\\SAC - Atalho.lnk";
            //result = File.Exists(path);
            //if (!result)
            //{
            //    listBoxResult.Items.Add("Enviar SAC - Atalho.lnk Para Área De Trabalho Pública");
            //}
            string path = "C:\\Users\\Public\\Desktop\\Caixa_NFCE.exe - Atalho.lnk";
            bool result = File.Exists(path);
            switch (lblModo.Text)
            {                
                case "Caixa":

                    path = "C:\\Users\\Public\\Desktop\\Caixa_NFCE.exe - Atalho.lnk";
                    result = File.Exists(path);
                    if (!result)
                    {
                        listBoxResult.Items.Add("Enviar Caixa_NFCE.exe - Atalho.lnk Para Área De Trabalho Pública");
                    }

                    path = "C:\\Users\\Public\\Desktop\\Prev32.exe - Atalho.lnk";
                    result = File.Exists(path);
                    if (!result)
                    {
                        listBoxResult.Items.Add("Enviar Prev32.exe - Atalho.lnk Para Área De Trabalho Pública");
                    }
                    break;

                case "Gerência":

                    path = "C:\\Users\\Public\\Desktop\\Caixa_NFCE.exe - Atalho.lnk";
                    result = File.Exists(path);
                    if (!result)
                    {
                        listBoxResult.Items.Add("Enviar Caixa_NFCE.exe - Atalho.lnk Para Área De Trabalho Pública");
                    }

                    path = "C:\\Users\\Public\\Desktop\\Prev32.exe - Atalho.lnk";
                    result = File.Exists(path);
                    if (!result)
                    {
                        listBoxResult.Items.Add("Enviar Prev32.exe - Atalho.lnk Para Área De Trabalho Pública");
                    }

                    path = "C:\\Users\\Public\\Desktop\\SAC - Atalho.lnk";
                    result = File.Exists(path);
                    if (!result)
                    {
                        listBoxResult.Items.Add("Enviar SAC - Atalho.lnk Para Área De Trabalho Pública");
                    }
                    break;
            }

        }

        public void impressoraCheck()
        {
            List<String> impressoras = new List<string>();
            foreach (String impressora in PrinterSettings.InstalledPrinters)
            {
                impressoras.Append(impressora);                
            }
            if (!impressoras.Exists(e => e.Contains("CaixaNFCE")))
            {
                listBoxResult.Items.Add("Impressora CaixaNFCE");
            }
        }

        //Método para checar pastas e arquivos
        public void checkPasta()
        {
            string path = "C:\\Program Files\\";
            bool result = Directory.Exists(path);


            path = "C:\\app\\product\\11.2.0";
            result = Directory.Exists(path);
            progressBar.Value += 5;
            progressBar.Text = progressBar.Value.ToString() + "%";
            if (result != true)
            {

                listBoxResult.Items.Add("Oracle Net Manager");

            }

            path = "C:\\oracle";
            result = Directory.Exists(path);
            progressBar.Value += 90;
            progressBar.Text = progressBar.Value.ToString() + "%";
            if (result != true)
            {
                listBoxResult.Items.Add(@"Pasta C:\Oracle");
            }

            path = "C:\\Windows\\Bemol.ini";
            result = File.Exists(path);
            progressBar.Value += 5;
            progressBar.Text = progressBar.Value.ToString() + "%";
            if (result != true)
            {
                listBoxResult.Items.Add("Bemol.ini");
            }

            path = @"C:\Program Files (x86)\LANDesk\LDClient\LANDeskPortalManager.exe";
            result = File.Exists(path);
            string pathDirectory = @"C:\Program Files (x86)\LANDesk";
            bool resultDirectory = Directory.Exists(pathDirectory);
            if (result != true || resultDirectory != true)
            {
                listBoxResult.Items.Add("Landesk");
            }


            switch (lblModo.Text)
            {                
                case "Caixa":
                    path = "C:\\EFCEBEMOL";
                    result = Directory.Exists(path);
                    if (result != true)
                    {
                        listBoxResult.Items.Add(@"Pasta C:\EFCEBEMOL");
                    }
                    break;
                case "Gerência":
                    path = "C:\\EFCEBEMOL";
                    result = Directory.Exists(path);
                    if (result != true)
                    {
                        listBoxResult.Items.Add(@"Pasta C:\EFCEBEMOL");
                    }
                    break;
            }
            // string command = "/C notepad.exe";
            // Process.Start("cmd.exe", command);
        }

        //método para verificar o nome do computador e o dominio
        public void hostName()
        {   
            
            var nome = Environment.MachineName;
            // var dominio = Environment.UserDomainName;
            var nomeCompleto = Dns.GetHostEntry(nome).HostName;

            if (!nomeCompleto.Equals(nome + ".bemol.local"))
            {
                listBoxResult.Items.Add("Inserir no Domínio bemol.local");
            }


        }

        //Verifica a versão do windows 10.
        public void editionWindows()
        {
            string tipoWindows = (OSVersion.GetOperatingSystem()).ToString();
            if (tipoWindows == "Windows10")
            {
                string versaoWin = OSVersion.MajorVersion10Properties().DisplayVersion;
                if (versaoWin != "21H2")
                {
                    listBoxResult.Items.Add("Atualizar o Windows para 21H2");
                }
            }
               
        }

        //Verifica se tem atualizações instaladas
        public void InstalledUpdates()
        {
            UpdateSession UpdateSession = new UpdateSession();
            IUpdateSearcher UpdateSearchResult = UpdateSession.CreateUpdateSearcher();
            UpdateSearchResult.Online = true;//checks for updates online
            ISearchResult SearchResults = UpdateSearchResult.Search("IsInstalled=1 AND IsHidden=0");
            //for the above search criteria refer to 
            //http://msdn.microsoft.com/en-us/library/windows/desktop/aa386526(v=VS.85).aspx
            //Check the remakrs section

            foreach (IUpdate x in SearchResults.Updates)
            {
                //listBoxResult.Items.Add(x.Title);
            }
        }

        //Verifica se tem atualizações pendentes
        public void UpdatesAvailable()
        {
            txtLog.AppendText(" Verificando Atualizações do Windows");
            try
            {
                UpdateSession UpdateSession = new UpdateSession();
                IUpdateSearcher UpdateSearchResult = UpdateSession.CreateUpdateSearcher();
                UpdateSearchResult.Online = true;//checks for updates online
                ISearchResult SearchResults = UpdateSearchResult.Search("IsInstalled=0 AND IsPresent=0");
                //for the above search criteria refer to 
                //http://msdn.microsoft.com/en-us/library/windows/desktop/aa386526(v=VS.85).aspx
                //Check the remakrs section
                if (SearchResults.Updates.Count == 0)
                {
                    txtLog.Text = "Windows Atualizado";
                    //Verifica se todos os programas estão instalado com sucesso
                    if (listBoxResult.Items.Count == 0)
                    {
                        txtLog.Text = "Todos Os Programas Instalados\nWindows Atualizado";
                    }
                }
                else
                {
                    listBoxResult.Items.Add("Windows Possúi Atualizações");
                    //foreach (IUpdate x in SearchResults.Updates)
                    //{
                    //    listBoxResult.Items.Add(x.Title);
                    //}
                }
            }
            catch
            {
                txtLog.Text = "Erro Ao Procurar Atualizações Do Windows\nVerifique A Conexão Com A Internet.";
            }
          
           
        }

        public void logsClear()
        {
            txtLog.Clear();
        }
        

        //Move a janela 
        private void start_MouseDown(object sender, MouseEventArgs e)
        {
            dragging = true;
            startPoint = new Point(e.X, e.Y);
        }

        private void start_MouseUp(object sender, MouseEventArgs e)
        {
            dragging = false;
        }

        private void start_MouseMove(object sender, MouseEventArgs e)
        {
            if(dragging)
            {
                Point p = PointToScreen(e.Location);
                Location = new Point(p.X - this.startPoint.X, p.Y - this.startPoint.Y);
            }
        }

        private void listBoxResult_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listBoxResult.SelectedItem != null)
            {
                //instancia classe aviso usada para exibir messagem de instalação
                Aviso aviso = new Aviso();
                Programas programas = new Programas();
                string programItem = listBoxResult.SelectedItem.ToString();
                switch (programItem)
                {
                    case "7-Zip":
                        if (Application.OpenForms.Count == 1)
                        {
                            //aviso.alert("7-Zip");
                            //programas.Show();
                            //programas.Imagens("7-zip");
                            //programas.AbrirForm();
                        }
                        break;
                    case "McAfee":
                        //listBoxResult.Items.Add("McAfee pode ser encontrado em \\srvlandeskdb\\App");
                        break;

                }
            }
        }

        
    }
}
