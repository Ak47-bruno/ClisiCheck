using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ClisiCheck
{
    public partial class Programas : Form
    {
        private bool dragging = false;
        private Point startPoint = new Point(0, 0);
        int largura = 0;
        int altura = 0;
        string nome = "";
        public Programas()
        {
            InitializeComponent();
        }

        public void AbrirForm()
        {
            this.Size = new Size(largura, 500);
            this.Text = nome;
            pnlComoFazer.Width = largura;
            pnlComoFazer.Height = altura;
            pictureBox1.Width = largura;
            pictureBox1.Height = altura;
            pnlComoFazer.Dock = DockStyle.Top;
            pnlComoFazer.Top = 50;
        }

        public void Imagens(string nomePrograma)
        {
            
            switch (nomePrograma)
            {
                case "7-zip":
                    largura = 1120;
                    altura = 1420;
                    nome = "7-Zip";
                    break;
            }
        }



        //Move a janela               
               

       

        private void panel1_MouseDown_1(object sender, MouseEventArgs e)
        {
            dragging = true;
            startPoint = new Point(e.X, e.Y);
        }

        private void panel1_MouseUp_1(object sender, MouseEventArgs e)
        {
            dragging = false;
        }

        private void panel1_MouseMove_1(object sender, MouseEventArgs e)
        {
            if (dragging)
            {
                Point p = PointToScreen(e.Location);
                Location = new Point(p.X - this.startPoint.X, p.Y - this.startPoint.Y);
            }
        }
    }
}
