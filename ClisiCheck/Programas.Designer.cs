
namespace ClisiCheck
{
    partial class Programas
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Programas));
            this.pnlComoFazer = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.imageFazer = new System.Windows.Forms.ImageList(this.components);
            this.pnlComoFazer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlComoFazer
            // 
            this.pnlComoFazer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlComoFazer.AutoScroll = true;
            this.pnlComoFazer.BackColor = System.Drawing.Color.White;
            this.pnlComoFazer.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pnlComoFazer.Controls.Add(this.pictureBox1);
            this.pnlComoFazer.Location = new System.Drawing.Point(12, 12);
            this.pnlComoFazer.Name = "pnlComoFazer";
            this.pnlComoFazer.Size = new System.Drawing.Size(776, 419);
            this.pnlComoFazer.TabIndex = 4;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.SteelBlue;
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Image = global::ClisiCheck.Properties.Resources._7zip1;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(776, 419);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // imageFazer
            // 
            this.imageFazer.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageFazer.ImageSize = new System.Drawing.Size(16, 16);
            this.imageFazer.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // Programas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.Color.SteelBlue;
            this.ClientSize = new System.Drawing.Size(800, 443);
            this.Controls.Add(this.pnlComoFazer);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Programas";
            this.pnlComoFazer.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel pnlComoFazer;
        private System.Windows.Forms.ImageList imageFazer;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}