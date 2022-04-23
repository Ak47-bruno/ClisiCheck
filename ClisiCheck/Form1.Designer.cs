
namespace ClisiCheck
{
    partial class start
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(start));
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.pnlNav = new System.Windows.Forms.Panel();
            this.btnGer = new System.Windows.Forms.Button();
            this.btnCaixa = new System.Windows.Forms.Button();
            this.btnDash = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lblModo = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnChek = new ClisiCheck.CircularButton();
            this.label2 = new System.Windows.Forms.Label();
            this.panel4 = new System.Windows.Forms.Panel();
            this.progressBar = new CircularProgressBar.CircularProgressBar();
            this.panel5 = new System.Windows.Forms.Panel();
            this.pnlResult = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.listBoxResult = new System.Windows.Forms.ListBox();
            this.btnMinimize = new System.Windows.Forms.Button();
            this.panel6 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel5.SuspendLayout();
            this.pnlResult.SuspendLayout();
            this.panel6.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.panel7);
            this.panel1.Controls.Add(this.pnlNav);
            this.panel1.Controls.Add(this.btnGer);
            this.panel1.Controls.Add(this.btnCaixa);
            this.panel1.Controls.Add(this.btnDash);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(200, 443);
            this.panel1.TabIndex = 0;
            // 
            // panel7
            // 
            this.panel7.BackColor = System.Drawing.Color.Black;
            this.panel7.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel7.Location = new System.Drawing.Point(0, 438);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(196, 1);
            this.panel7.TabIndex = 3;
            // 
            // pnlNav
            // 
            this.pnlNav.BackColor = System.Drawing.Color.RoyalBlue;
            this.pnlNav.Location = new System.Drawing.Point(0, 121);
            this.pnlNav.Name = "pnlNav";
            this.pnlNav.Size = new System.Drawing.Size(3, 100);
            this.pnlNav.TabIndex = 2;
            // 
            // btnGer
            // 
            this.btnGer.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnGer.FlatAppearance.BorderSize = 0;
            this.btnGer.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGer.Font = new System.Drawing.Font("Nirmala UI", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnGer.ForeColor = System.Drawing.Color.RoyalBlue;
            this.btnGer.Image = global::ClisiCheck.Properties.Resources.manager;
            this.btnGer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGer.Location = new System.Drawing.Point(0, 189);
            this.btnGer.Name = "btnGer";
            this.btnGer.Size = new System.Drawing.Size(196, 37);
            this.btnGer.TabIndex = 1;
            this.btnGer.Text = "Gerência";
            this.btnGer.UseVisualStyleBackColor = true;
            this.btnGer.Click += new System.EventHandler(this.btnGer_Click);
            this.btnGer.Leave += new System.EventHandler(this.btnGer_Leave);
            // 
            // btnCaixa
            // 
            this.btnCaixa.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnCaixa.FlatAppearance.BorderSize = 0;
            this.btnCaixa.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCaixa.Font = new System.Drawing.Font("Nirmala UI", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnCaixa.ForeColor = System.Drawing.Color.RoyalBlue;
            this.btnCaixa.Image = global::ClisiCheck.Properties.Resources.payment_method;
            this.btnCaixa.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCaixa.Location = new System.Drawing.Point(0, 152);
            this.btnCaixa.Name = "btnCaixa";
            this.btnCaixa.Size = new System.Drawing.Size(196, 37);
            this.btnCaixa.TabIndex = 1;
            this.btnCaixa.Text = "Caixa";
            this.btnCaixa.UseVisualStyleBackColor = true;
            this.btnCaixa.Click += new System.EventHandler(this.btnCaixa_Click);
            this.btnCaixa.Leave += new System.EventHandler(this.btnCaixa_Leave);
            // 
            // btnDash
            // 
            this.btnDash.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnDash.FlatAppearance.BorderSize = 0;
            this.btnDash.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDash.Font = new System.Drawing.Font("Nirmala UI", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnDash.ForeColor = System.Drawing.Color.RoyalBlue;
            this.btnDash.Image = global::ClisiCheck.Properties.Resources.work;
            this.btnDash.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnDash.Location = new System.Drawing.Point(0, 115);
            this.btnDash.Name = "btnDash";
            this.btnDash.Size = new System.Drawing.Size(196, 37);
            this.btnDash.TabIndex = 1;
            this.btnDash.Text = "Escritório";
            this.btnDash.UseVisualStyleBackColor = true;
            this.btnDash.Click += new System.EventHandler(this.btnDash_Click);
            this.btnDash.Leave += new System.EventHandler(this.btnDash_Leave);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.White;
            this.panel2.Controls.Add(this.pictureBox1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(196, 115);
            this.panel2.TabIndex = 0;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ClisiCheck.Properties.Resources.bedigital;
            this.pictureBox1.Location = new System.Drawing.Point(27, 15);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(146, 85);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // lblModo
            // 
            this.lblModo.AutoSize = true;
            this.lblModo.Font = new System.Drawing.Font("Microsoft Sans Serif", 21F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lblModo.ForeColor = System.Drawing.Color.White;
            this.lblModo.Location = new System.Drawing.Point(221, 15);
            this.lblModo.Name = "lblModo";
            this.lblModo.Size = new System.Drawing.Size(143, 32);
            this.lblModo.TabIndex = 1;
            this.lblModo.Text = "Escritório";
            // 
            // btnClose
            // 
            this.btnClose.FlatAppearance.BorderSize = 0;
            this.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClose.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnClose.ForeColor = System.Drawing.Color.White;
            this.btnClose.Location = new System.Drawing.Point(763, 12);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(25, 25);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "x";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // panel3
            // 
            this.panel3.AutoScroll = true;
            this.panel3.BackColor = System.Drawing.Color.White;
            this.panel3.Controls.Add(this.btnChek);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Location = new System.Drawing.Point(221, 50);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(218, 151);
            this.panel3.TabIndex = 3;
            // 
            // btnChek
            // 
            this.btnChek.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnChek.AutoSize = true;
            this.btnChek.BackColor = System.Drawing.Color.White;
            this.btnChek.BackgroundImage = global::ClisiCheck.Properties.Resources.icons8_play_dentro_de_um_círculo_80;
            this.btnChek.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnChek.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnChek.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnChek.Font = new System.Drawing.Font("Segoe UI", 5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnChek.ForeColor = System.Drawing.Color.White;
            this.btnChek.Location = new System.Drawing.Point(62, 29);
            this.btnChek.Margin = new System.Windows.Forms.Padding(1);
            this.btnChek.Name = "btnChek";
            this.btnChek.Size = new System.Drawing.Size(92, 94);
            this.btnChek.TabIndex = 1;
            this.btnChek.UseVisualStyleBackColor = false;
            this.btnChek.Click += new System.EventHandler(this.btnChek_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Nirmala UI", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label2.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label2.Location = new System.Drawing.Point(26, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(174, 28);
            this.label2.TabIndex = 0;
            this.label2.Text = "Iniciar Checagem";
            // 
            // panel4
            // 
            this.panel4.AutoScroll = true;
            this.panel4.BackColor = System.Drawing.Color.White;
            this.panel4.Controls.Add(this.progressBar);
            this.panel4.Location = new System.Drawing.Point(221, 245);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(218, 186);
            this.panel4.TabIndex = 3;
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.AnimationSpeed = 500;
            this.progressBar.BackColor = System.Drawing.Color.White;
            this.progressBar.Font = new System.Drawing.Font("Nirmala UI", 39.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.progressBar.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.progressBar.InnerColor = System.Drawing.Color.White;
            this.progressBar.InnerMargin = 2;
            this.progressBar.InnerWidth = -1;
            this.progressBar.Location = new System.Drawing.Point(22, 3);
            this.progressBar.MarqueeAnimationSpeed = 500;
            this.progressBar.Name = "progressBar";
            this.progressBar.OuterColor = System.Drawing.Color.Gray;
            this.progressBar.OuterMargin = -25;
            this.progressBar.OuterWidth = 26;
            this.progressBar.ProgressColor = System.Drawing.Color.RoyalBlue;
            this.progressBar.ProgressWidth = 6;
            this.progressBar.SecondaryFont = new System.Drawing.Font("Nirmala UI", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.progressBar.Size = new System.Drawing.Size(168, 168);
            this.progressBar.StartAngle = 270;
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.SubscriptColor = System.Drawing.Color.FromArgb(((int)(((byte)(166)))), ((int)(((byte)(166)))), ((int)(((byte)(166)))));
            this.progressBar.SubscriptMargin = new System.Windows.Forms.Padding(10, -35, 0, 0);
            this.progressBar.SubscriptText = "";
            this.progressBar.SuperscriptColor = System.Drawing.Color.FromArgb(((int)(((byte)(166)))), ((int)(((byte)(166)))), ((int)(((byte)(166)))));
            this.progressBar.SuperscriptMargin = new System.Windows.Forms.Padding(10, 35, 0, 0);
            this.progressBar.SuperscriptText = "";
            this.progressBar.TabIndex = 1;
            this.progressBar.TextMargin = new System.Windows.Forms.Padding(8, 8, 0, 0);
            this.progressBar.Value = 68;
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.Color.White;
            this.panel5.Controls.Add(this.pnlResult);
            this.panel5.Location = new System.Drawing.Point(445, 50);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(343, 381);
            this.panel5.TabIndex = 3;
            // 
            // pnlResult
            // 
            this.pnlResult.AutoScroll = true;
            this.pnlResult.Controls.Add(this.label4);
            this.pnlResult.Controls.Add(this.listBoxResult);
            this.pnlResult.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlResult.Location = new System.Drawing.Point(0, 0);
            this.pnlResult.Name = "pnlResult";
            this.pnlResult.Size = new System.Drawing.Size(343, 381);
            this.pnlResult.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Nirmala UI", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label4.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label4.Location = new System.Drawing.Point(122, 10);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(93, 28);
            this.label4.TabIndex = 2;
            this.label4.Text = "Faltando";
            // 
            // listBoxResult
            // 
            this.listBoxResult.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.listBoxResult.Font = new System.Drawing.Font("Nirmala UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.listBoxResult.ForeColor = System.Drawing.Color.RoyalBlue;
            this.listBoxResult.FormattingEnabled = true;
            this.listBoxResult.HorizontalScrollbar = true;
            this.listBoxResult.ItemHeight = 21;
            this.listBoxResult.Location = new System.Drawing.Point(0, 41);
            this.listBoxResult.Name = "listBoxResult";
            this.listBoxResult.Size = new System.Drawing.Size(343, 315);
            this.listBoxResult.TabIndex = 0;
            this.listBoxResult.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listBoxResult_MouseDoubleClick);
            // 
            // btnMinimize
            // 
            this.btnMinimize.FlatAppearance.BorderSize = 0;
            this.btnMinimize.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMinimize.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.btnMinimize.ForeColor = System.Drawing.Color.White;
            this.btnMinimize.Location = new System.Drawing.Point(732, 12);
            this.btnMinimize.Name = "btnMinimize";
            this.btnMinimize.Size = new System.Drawing.Size(25, 25);
            this.btnMinimize.TabIndex = 2;
            this.btnMinimize.Text = "-";
            this.btnMinimize.UseVisualStyleBackColor = true;
            this.btnMinimize.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel6
            // 
            this.panel6.AutoScroll = true;
            this.panel6.BackColor = System.Drawing.Color.White;
            this.panel6.Controls.Add(this.label3);
            this.panel6.Location = new System.Drawing.Point(221, 209);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(218, 38);
            this.panel6.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Nirmala UI", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label3.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label3.Location = new System.Drawing.Point(74, 5);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 28);
            this.label3.TabIndex = 1;
            this.label3.Text = "Status";
            // 
            // start
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.Color.SteelBlue;
            this.ClientSize = new System.Drawing.Size(792, 443);
            this.Controls.Add(this.panel6);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.btnMinimize);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.lblModo);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "start";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CliSiCheck";
            this.Load += new System.EventHandler(this.start_Load);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.start_MouseDown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.start_MouseMove);
            this.MouseUp += new System.Windows.Forms.MouseEventHandler(this.start_MouseUp);
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.pnlResult.ResumeLayout(false);
            this.pnlResult.PerformLayout();
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnGer;
        private System.Windows.Forms.Button btnCaixa;
        private System.Windows.Forms.Button btnDash;
        private System.Windows.Forms.Panel pnlNav;
        private System.Windows.Forms.Label lblModo;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label2;
        private CircularButton btnChek;
        private System.Windows.Forms.Panel panel4;
        private CircularProgressBar.CircularProgressBar progressBar;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Button btnMinimize;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.Panel pnlResult;
        private System.Windows.Forms.ListBox listBoxResult;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Label label3;
    }
}