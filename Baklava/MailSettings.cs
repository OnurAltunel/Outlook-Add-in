using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Baklava;

namespace Baklava
{
    public partial class MailSettings : Form
    {
        public void addmaildelete(string mail)
        {
            if (File.Exists("tomail.txt"))
            {
                File.Delete("tomail.txt");
                File.AppendAllText("tomail.txt", mail);
            }
            else
            {
                File.AppendAllText("tomail.txt", mail);
            }
        }
        public void addmailcreate(string mail)
        {
            File.Create("tomail.txt");           
        }
        public void addadmindelete(string admin)
        {
            if (File.Exists("admin.txt"))
            {
                File.Delete("admin.txt");
                File.AppendAllText("admin.txt", admin);
            }
            else
            {
                File.AppendAllText("admin.txt", admin);
            }
        }        
        public MailSettings()
        {
            InitializeComponent();
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            addmaildelete(textBox1.Text);         
            addadmindelete(textBox2.Text);         
        }

        
    }
}
