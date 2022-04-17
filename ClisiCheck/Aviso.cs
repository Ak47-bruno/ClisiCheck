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
namespace ClisiCheck
{
    class Aviso {

        public void alert(string program)
        {
            string message = "Deseja instalar "+ program+"?";
            string title = "Aviso!";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Yes)
            {
                string command = "/C notepad.exe";
                Process.Start("cmd.exe", command);
            }
            else
            {
                
            }
        }

    }

    
    
}

