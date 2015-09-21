using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Collections;
using Outlook = Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading;

namespace ViadeoOutlookAddIn
{
    public partial class Ribbon1
    {
        public string dataPath = string.Empty;
        public ArrayList contactList = new ArrayList();
        ProgressBar pbr = new ProgressBar();        
        l1 p = new l1();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        public void myProgressBar()
        {            
            pbr.Width = 250;
            p.Controls.Add(pbr);            
            p.Show();
            pbr.Style = ProgressBarStyle.Marquee;
            
        }
        public void SearchInCurrentBox()
        {            
            
            Outlook.MAPIFolder inbox = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder;
            Outlook.Items items = inbox.Items;
            Outlook.MailItem mailItem = null;
            object folderItem;
            string bodyContent = string.Empty;
            folderItem = items;
            folderItem = items.GetFirst();
            while (folderItem != null)
            {
                mailItem = folderItem as Outlook.MailItem;
                if (mailItem != null)
                {
                                        
                    bodyContent = mailItem.Body;
                }
                findContactViadeo(bodyContent);
                folderItem = items.GetNext();
            }

            
            MessageBox.Show("fin d'extraction, Clicker sur OK pour l'ancer l'exportation");
            
            extractContact(contactList);           
            contactList.Clear();

        }
        //=========================================================
        public string theHeadline(string data)
        {
            string val = string.Empty;
            String[] substrings = Regex.Split(data, ",");
            if ((substrings.Length > 2))
            {
                for (int i = 0; i < (substrings.Length - 1); i++)
                {
                    val += substrings.GetValue(i).ToString() + ",";
                }
            }
            else { val = substrings.GetValue(0).ToString(); }

            return val;
        }
        public string theCompany(string data)
        {
            string val = string.Empty;
            String[] substrings = Regex.Split(data, ",");

            val = substrings.GetValue(substrings.Length - 1).ToString();

            return val;
        }
        public string thefirstName(string data)
        {
            string val = string.Empty;
            String[] substrings = Regex.Split(data, " ");
            if ((substrings.Length > 2))
            {
                for (int i = 0; i < (substrings.Length - 1); i++)
                {
                    val += " " + substrings.GetValue(i).ToString();
                }
            }
            else { val = substrings.GetValue(0).ToString(); }

            return val;
        }
        public string theLastName(string data)
        {
            string val = string.Empty;
            String[] substrings = Regex.Split(data, " ");

            val = substrings.GetValue(substrings.Length - 1).ToString();

            return val;
        }
        public string cleanContact(string findword)
        {
            string pattern = "\"";
            string replacement = "";
            string result = Regex.Replace(findword, pattern, replacement);

            return result;
        }
        //=========================================================
        public void findContactViadeo(string content)
        {
            string pattern;
            string fonction = string.Empty;
            string val = string.Empty;
            string val1 = string.Empty;

            pattern = "[\"]{1}[\\s]*[\\w]*[\\s]*[\\w]*\\s[\\w]*|[\"]{1}[\\s]*[\\w]*[\\s]*[\\w]*[-]{1}[\\w]*[\\r]";
            Regex contactViadeo = new Regex(pattern, RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline);

            string contactName = cleanContact(findWord(content));

            ArrayList theContacts = new ArrayList();
            theContacts.Clear();
            String[] substrings = Regex.Split(contactName, "\n");
            foreach (string word in substrings)
            {
                if (word.Trim() != "") theContacts.Add(word.Trim());
            }

            TextBox Box = new TextBox();
            Box.Clear();
            Box.Multiline = true;
            Box.Text = content;

            for (int i = 0; i < Box.Lines.Count(); i++)
            {
                val = Box.Lines[i];

                if (val.Trim() != "")
                {
                    Match m = contactViadeo.Match(val);
                    string v = cleanContact(m.Value);
                    string theDate = findDate(content);                    
                    string firstName;
                    string lastName;
                    string headLine;
                    string companyName;
                    string joinDate;                                 
                    if (theContacts.Contains(v.Trim()))
                    {
                        fonction = Box.Lines[i + 1];                        
                        firstName = thefirstName(v.Trim());
                        lastName = theLastName(v.Trim());
                        headLine = theHeadline(fonction);
                        companyName = theCompany(fonction);
                        joinDate = theDate;
                        l0 viadeoContact = new l0 { firstName = firstName, lastName = lastName, headLine = headLine, companyName = companyName, joinDate = joinDate };
                        contactList.Add(viadeoContact);                        
                    }                    

                }
            }
            

        }        
        public string findWord(string Content)
        {            
            string contactPattern;                                    
            contactPattern = "[\"]{1}[\\s]*[\\w]*[\\w]*\\s[\\w]*[\\r]|[\"][\\s]*[\\w]*[\\s]*[\\w]*[-]{1}[\\w]*[\\r]|[\"]{1}[\\s]*[\\w]*[\\s]*[\\w]{1,}{1}[\\r]";            

            Regex entrepriseViadeo = new Regex(contactPattern, RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            string val = string.Empty;
            MatchCollection matches = entrepriseViadeo.Matches(Content);
            foreach (Match match in matches)
            {
                val +=match.Value;                                     
            }            
            return val;
        }        
        public string findDate(string Content)
        {
            string datePattern;
            datePattern = @"[\w]*\s[\d][\d]\s[\w]*\s[\d][\d][\d][\d]|[\w]*\s[\d]\s\s[\d][\d][\d][\d]";                                

            Regex entrepriseViadeo = new Regex(datePattern, RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            string theDate = string.Empty;
            MatchCollection matches = entrepriseViadeo.Matches(Content);
            foreach (Match match in matches)
            {
                theDate = match.Value.Trim();
            }

            return theDate;
        }              
        public void extractContact(ArrayList data)
        {                        
            var excelApp = new Excel.Application();            
            excelApp.Visible = true;
            excelApp.Workbooks.Add();

            try
            {                
                Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                workSheet.Name = "Contacts Viadeo";                
                workSheet.Cells[1, "A"] = "Prénom";
                workSheet.Cells[1, "B"] = "Nom";
                workSheet.Cells[1, "C"] = "Fonction";
                workSheet.Cells[1, "D"] = "Société";
                workSheet.Cells[1, "E"] = "Date Inscription";                
                var row = 1;
                foreach (l0 objCourriel in data)
                {
                    row++;
                    workSheet.Cells[row, "A"] = objCourriel.firstName.Trim();
                    workSheet.Cells[row, "B"] = objCourriel.lastName.Trim();
                    workSheet.Cells[row, "C"] = objCourriel.headLine;
                    workSheet.Cells[row, "D"] = objCourriel.companyName.Trim();
                    workSheet.Cells[row, "E"] = objCourriel.joinDate;                    
                }
             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }            

        }
        
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            SearchInCurrentBox();            
        }
    }
}
