using Microsoft.Office.Core;
using System;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Globalization;

namespace Baklava
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        public void BaklavaButonu(Office.IRibbonControl control)
        {
            if (control.Id == "baklavabuton") 
            {              
                string Konu = "Hediye";
                string Mesaj = "BAKLAVA";
                stringtohtml(Konu, Mesaj);            
            }
        }
        public void CayButonu(Office.IRibbonControl control)
        {
            if (control.Id == "caybuton")
            {
                string Konu = "Hediye";
                string Mesaj = "CAY";
                stringtohtml(Konu, Mesaj);
            }
        }
        public void KahveButonu(Office.IRibbonControl control)
        {
            if (control.Id == "kahvebuton")
            {
                string Konu = "Hediye";
                string Mesaj = "KAHVE";
                stringtohtml(Konu, Mesaj);
            }
        }
        public void LatteButonu(Office.IRibbonControl control)
        {
            if (control.Id == "lattebuton")
            {
                string Konu = "Hediye";
                string Mesaj = " LATTE";
                stringtohtml(Konu, Mesaj);
            }
        }      
        public void stringtohtml(string konu, string mesaj)
        {
            CultureInfo tr = new CultureInfo("tr-TR");
            string tarih = DateTime.Now.ToString("dddd, dd MMMM yyyy", tr);
            string Konu = "Hediye";          
            string Mesaj = "<!DOCTYPE html><html><head><title></title> </head>" +
                "<body><center><h1 color='red'>HERKESE BENDEN " + mesaj + "</h1>" +             
                "<p1> "+ tarih + "</p1>" +
                "<p1> Afiyet Olsun </p1></center></body></html>";           
            Globals.ThisAddIn.PermissionMessage(Konu,Mesaj);
        }
        public void AyarButonu(Office.IRibbonControl control)
        {
            if (control.Id == "AyarButonu")
            {
                MailSettings mailayari = new MailSettings();
                mailayari.Show();
            }
        }


        public System.Drawing.Bitmap Hediyesentimage(IRibbonControl control)
        {
            return Properties.Resources.Hediyesentimage;
        }
        public System.Drawing.Bitmap Baklavaimage(IRibbonControl control)
        {
            return Properties.Resources.Baklavaimage;
        }
        public System.Drawing.Bitmap Cayimage(IRibbonControl control)
        {
            return Properties.Resources.Cayimage;
        }
        public System.Drawing.Bitmap Kahveimage(IRibbonControl control)
        {
            return Properties.Resources.Kahveimage;
        }
        public System.Drawing.Bitmap Latteimage(IRibbonControl control)
        {
            return Properties.Resources.Latteimage;
        }
        public System.Drawing.Bitmap settings(IRibbonControl control)
        {
            return Properties.Resources.settings;
        }


        private Office.IRibbonUI ribbon;
        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Baklava.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks      

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
