using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.DirectoryServices;
namespace ClisiCheck
{
    public partial class Login : Form
    {
        private bool dragging = false;
        private Point startPoint = new Point(0, 0);

        public Login()
        {
            InitializeComponent();
        }

        private void btnConfirmar_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtUser.Text) && !string.IsNullOrWhiteSpace(txtSenha.Text))
            {
                try
                {
                    DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://bemol.local", txtUser.Text, txtSenha.Text);
                    DirectorySearcher directorySearcher = new DirectorySearcher(directoryEntry);
                    directorySearcher.Filter = "(sAMAccountName=" + txtUser.Text + ")";
                    SearchResult searchResult = directorySearcher.FindOne();
                    if ((Int32)searchResult.Properties["userAccountControl"][0] == 512)
                    {
                        start start = new start();
                        this.Hide();                        
                        start.Show();
                        start.userLabel(txtUser.Text);
                    }
                    else
                    {
                        MessageBox.Show("ERRO: Usuário/Senha Inválido!");
                    }


                }
                catch (Exception)
                {
                    MessageBox.Show("ERRO: Usuário Não encontrado!");
                }
            }
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
            if (dragging)
            {
                Point p = PointToScreen(e.Location);
                Location = new Point(p.X - this.startPoint.X, p.Y - this.startPoint.Y);
            }
        }


    }
}
