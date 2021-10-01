using System;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Globalization;


// TO DO LIST //

// Acilmis outlook hesabina girabilmeliyim +
// Outlook hesabindan kisilere ulasabilmeliyim + 
// Basit mail gonderebilmeliyim +
// istege bagli ek gonderebilmeliyim +
// Birden fazla kisiye mail gonderebilmeliyim +
// Outlook Acıldıgı anda bu ıslemı yapıyor bunu tetiklemelıyız +
// Tus eklemelı ya da dısardan mudehale edebılmelıyım +
// Gonderılen maıl kısılerımden yakın olanları bulmalı onceden koda yazılanları degıl -
// Ayri ayri kisinin yoneticisini ve calisma arkadaslarini bulmali ve maili sadece bunlara gondermeli -
// Mail secenekleri olusturulmali +
// Kimden kime ne zaman mail gittiği bilgilendirilmeli +
// Butonlara resim eklenmeli + 
// Disardan mudahale icin ayri bir projedeki form ile birlestirilmeli -



namespace Baklava
{

    class mails
    {
        public string Admin;
        public string Tomail;

        public string admin    // the Name property
        {
            get => Admin;
            set => Admin = value;
        }
        public string tomail    // the Name property
        {
            get => Tomail;
            set => Tomail = value;
        }

    }   
    public partial class ThisAddIn
    {
        mails yenimailler = new mails();
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            mailbelirleme();
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility
        CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
        public void mailbelirleme()
        {
            var path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) ;
            var subFolderPath = Path.Combine(path, "tomail.txt");
            var subFolderPath2 = Path.Combine(path, "admin.txt");
            string tomail = File.ReadAllText(subFolderPath);  
            string adminmail = File.ReadAllText(subFolderPath2);
                yenimailler.Admin  = adminmail;
                yenimailler.Tomail = tomail;            
        }
        public void fromRibbontostring(string konu, string mesaj)
        {
            PermissionMessage(konu, mesaj);
        }
        public void PermissionMessage(string konu, string mesaj)
        {
            DialogResult secenek = MessageBox.Show("Gondermek istedigine emin misin ?", "",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (secenek == DialogResult.Yes)
            { FindContacts(konu, mesaj); }
            else if (secenek == DialogResult.No)
            { MessageBox.Show(" !!! Hadi yine iyisin !!! "); }
        }
        public void FindContacts(string konu, string mesaj)
        {          
            Outlook.MAPIFolder sentContacts = (Outlook.MAPIFolder)
            this.Application.ActiveExplorer().Session.GetDefaultFolder
            (Outlook.OlDefaultFolders.olFolderContacts);          
            foreach (Outlook.ContactItem contact in sentContacts.Items)
            {
                if (contact.Email1Address.Contains(yenimailler.Tomail))
                {
                    this.CreateEmailItem(konu,contact.Email1Address,mesaj);                   
                }               
            }            
            AdressCounter(yenimailler.Tomail,mesaj);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        public void CreateEmailItem(string konu, string toEmail, string mesaj)
        {
            Outlook.MailItem eMail = (Outlook.MailItem)
            this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            eMail.Subject = konu;
            eMail.To = toEmail;
            eMail.HTMLBody = mesaj;
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
            //eMail.Attachments.Add(@"c:\pngs\hb.png",
            //Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);           
            ((Outlook._MailItem)eMail).Send();
            mailbelirleme();
        }

        public void SendInformtoAdmin(string kisi,string sayi,string hediye)
        {            
            Outlook.MAPIFolder sentAdmin = (Outlook.MAPIFolder)
            this.Application.ActiveExplorer().Session.GetDefaultFolder
            (Outlook.OlDefaultFolders.olFolderContacts);

            foreach (Outlook.ContactItem contact in sentAdmin.Items)
            {
                if (contact.Email1Address.Contains(yenimailler.Admin))
                {
            Outlook.MailItem eMail = (Outlook.MailItem)
            this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            eMail.Subject = "Bilgilendirme";
            eMail.To = yenimailler.Admin;
                    eMail.HTMLBody = kisi + " " + sayi + 
                        " kisiye Sunu gonderdi " + hediye + "gonderdi";
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
                    try
                    {
                        ((Outlook._MailItem)eMail).Send();  
                        this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show("error");
                                   
                    }
            
                }
            }          
          
        }   
        public void InformMessage(int sayi,string isim)
        {
            CultureInfo tr = new CultureInfo("tr-TR");
            string tarih = DateTime.Now.ToString("dddd, dd MMMM yyyy", tr);
            string text = isim + "  ismini iceren ("+ sayi +") kisilere "+
                tarih +" tarihinde gonderildi";
            MessageBox.Show(text);
        }      
        public void AdressCounter(string AdresMaili,string mesaj)
        {   Outlook.MAPIFolder folderContacts = this.Application.ActiveExplorer().Session.
            GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            Outlook.Items searchFolder = folderContacts.Items;
            int sayac = 0;
            foreach (Outlook.ContactItem bulunankisi in searchFolder)
            {
                if (bulunankisi.Email1Address.Contains(AdresMaili))
                {                   
                    sayac = sayac + 1;
                }
            }                       
            InformMessage(sayac, AdresMaili);
            SendInformtoAdmin(AdresMaili,sayac.ToString(),mesaj);
        }
        public void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.OnShutdown();
        }
       #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);    
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);      
        }           
        #endregion
    }
    
}
