using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Baklava
{
    public partial class baklavaformu : Form
    {
        public baklavaformu()
        {
            InitializeComponent();
        }
        string anaMesaj = "Herkese benden";
        string ve = " ve";
        string mesajBaklava = " Baklava";
        string mesajCay = " Cay";
        string mesajKahve = " Kahve";
        string mesajLatte = " Latte";
       
        private void label1_Click_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                if (anaMesaj != "Herkese benden") { anaMesaj += ve + mesajBaklava; }
                else anaMesaj += mesajBaklava;
                Globals.ThisAddIn.fromRibbontostring(anaMesaj, anaMesaj);
            }
            if (checkBox2.Checked)
            {
                if (anaMesaj != "Herkese benden") { anaMesaj += ve + mesajCay; }
                else anaMesaj += mesajCay;
                Globals.ThisAddIn.fromRibbontostring(anaMesaj, anaMesaj);
            }
            if (checkBox3.Checked)
            {
                if (anaMesaj != "Herkese benden") { anaMesaj += ve + mesajKahve; }
                else anaMesaj += mesajKahve;
                Globals.ThisAddIn.fromRibbontostring(anaMesaj, anaMesaj);
            }
            if (checkBox4.Checked)
            {
                if (anaMesaj != "Herkese benden") { anaMesaj += ve + mesajLatte; }
                else anaMesaj += mesajLatte;
                Globals.ThisAddIn.fromRibbontostring(anaMesaj, anaMesaj);
            }
            //buraya gonderme fonksiyonu gelicek
        
            anaMesaj = "Herkese benden" ;
           
            
        }
      
    }
}
