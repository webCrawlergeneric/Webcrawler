
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Menu;
using Excel = Microsoft.Office.Interop.Excel;

namespace Documents
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox4.Enabled = false;

            List<string> items = new List<string>() { "2", "4", "6" };
            comboBox1.DataSource = items;
        }


        OpenFileDialog ofd = new OpenFileDialog();
        FirefoxProfile prof = new FirefoxProfile();
        List<Document> listDocuments = new List<Document>();
        List<Document> listSavedPdf = new List<Document>();
        private void button1_Click(object sender, EventArgs e)
        {
            ofd.Filter = "xlsx|*.xlsx";
            listDocuments = new List<Document>();
            if (ofd.ShowDialog() == DialogResult.OK)
            {

                //                myCombo.DataSource = new ComboItem[] {
                //    new ComboItem{ ID = 1, Text = "One" },
                //    new ComboItem{ ID = 2, Text = "Two" },
                //    new ComboItem{ ID = 3, Text = "Three" }
                //};

                textBox2.Text = ofd.FileName;
                //  textBox3.Text = ofd.SafeFileName;

                Excel.Application xlApp1 = new Excel.Application();
                Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(ofd.FileName);
                Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
                Excel.Range xlRange1 = xlWorksheet1.UsedRange;
                int rowCount1 = xlRange1.Rows.Count;
                int columnCount1 = xlRange1.Columns.Count;


                for (int i = 2; i <= rowCount1; i++)
                {
                    Document document = Document.getDocumentObj();
                    document.Legalname = xlRange1.Cells[i, 1].Value2;
                    document.AccountNumber = xlRange1.Cells[i, 2].Value2;
                    document.Country = xlRange1.Cells[i, 3].Value2;
                    document.Address = xlRange1.Cells[i, 4].Value2;
                    document.City = xlRange1.Cells[i, 5].Value2;
                    document.Province = xlRange1.Cells[i, 6].Value2;
                    document.Postalcode = xlRange1.Cells[i, 7].Value2;
                    document.Phonenumber = xlRange1.Cells[i, 8].Value2;
                    document.Ext = xlRange1.Cells[i, 9].Value2;
                    document.Emailaddress = xlRange1.Cells[i, 10].Value2;
                    //document.WorkSafeBC_Legalname = xlRange1.Cells[i, 11].Value2;
                    document.Account = xlRange1.Cells[i, 11].Value2;
                    document.Legalname_Tradename = xlRange1.Cells[i, 12].Value2;
                    document.ClientCode = xlRange1.Cells[i, 13].Value2;
                    listDocuments.Add(document);
                }
                // myList.GroupBy(x => x.prop1).Select(y => y.First());
                listDocuments = listDocuments.GroupBy(x => x.Legalname_Tradename).Select(y => y.First()).ToList();
            }
            textBox4.Text = listDocuments.Count().ToString();
            listDocuments.RemoveAll(x => listSavedPdf.Contains(x));
        }



        private void button2_Click(object sender, EventArgs e)
        {

            //  SaveFileDialog objSav = new SaveFileDialog();
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            folderDialog.Description = "Select the Folder where do you want to save documents";
            folderDialog.ShowNewFolderButton = false;

            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = folderDialog.SelectedPath;
                DirectoryInfo d = new DirectoryInfo(textBox3.Text);//Assuming Test is your Folder
                FileInfo[] Files = d.GetFiles("*.pdf"); //Getting Text files


                foreach (FileInfo file in Files)
                {
                    Document obj = Document.getDocumentObj();
                    obj.Legalname = file.Name.Replace(".pdf", "");
                    obj.Status = "Success";
                    listSavedPdf.Add(obj);
                }
            }
        }

        private Object thisLock = new Object();


        private void startingThread(List<Document> listThreadDocument)
        {
            if (listThreadDocument.Count > 0)
            {

                // List<Document> listThreadDocument = new List<Document>();

                var specialcharacters = ConfigurationManager.AppSettings["SpecialChar"];
                DirectoryInfo d = new DirectoryInfo(textBox3.Text);//Assuming Test is your Folder
                FileInfo[] Files = d.GetFiles("*.pdf"); //Getting Text files
                                                        //lock (thisLock)
                                                        //{
                for (int i = 0; i < listThreadDocument.Count(); i++)
                {
                    if (listThreadDocument[i].Legalname_Tradename != null && listThreadDocument[i].Legalname_Tradename != "")
                    {
                        listThreadDocument[i].Legalname = listThreadDocument[i].Legalname.Trim();
                        listThreadDocument[i].Legalname_Tradename = Regex.Replace(listThreadDocument[i].Legalname_Tradename, @"[^0-9a-zA-Z]+", specialcharacters);
                        var chromeOptions = new ChromeOptions();
                        chromeOptions.AddUserProfilePreference("download.default_directory", textBox3.Text);
                        chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
                        chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                        chromeOptions.AddArgument("sujeeth" + i);
                        IWebDriver driver = new ChromeDriver(chromeOptions);

                        driver.Navigate().GoToUrl("https://online.worksafebc.com/Anonymous/EmployerClearanceLetter/Default.aspx");
                        IWebElement element = driver.FindElement(By.Id("ctl00_middle_btnAddEmployer"));
                        Thread.Sleep(2330);
                        element.Click();
                        Thread.Sleep(2330);
                        driver.FindElement(By.Id("ctl00_middle_radSearchTypeName")).Click();
                        Thread.Sleep(2330);
                        element = driver.FindElement(By.Id("ctl00_middle_txtLegalName"));
                        element.SendKeys(listThreadDocument[i].Legalname_Tradename);
                        Thread.Sleep(2330);
                        driver.FindElement(By.Id("ctl00_middle_btnNameSearch")).Click();
                        Thread.Sleep(2330);
                        element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults"));

                        string clientName = listThreadDocument[i].ClientCode + '_' + listThreadDocument[i].Legalname_Tradename;
                        var rowcount = element.FindElements(By.TagName("tr")).Count();
                        var noRecords = driver.FindElements(By.CssSelector("table tr"));
                        if (noRecords[5].Text.ToString() == "No firms found.")
                        {

                            Document obj = Document.getDocumentObj();
                            obj.Legalname = clientName;
                            obj.Status = "No firms found.";
                            saveExcel(obj);
                            //Document obj = Document.getDocumentObj();
                            //obj.Legalname = clientName;
                            //obj.Status = "No firms found.";
                            //listSavedPdf.Add(obj);


                            //listSavedPdf = listSavedPdf.GroupBy(x => x.Legalname).Select(y => y.First()).ToList();

                            //xlWorkSheet.Cells[1, 1] = "Legalname";
                            //xlWorkSheet.Cells[1, 2] = "Status";
                            //for (int j = 0; j < listSavedPdf.Count; j++)
                            //{
                            //    xlWorkSheet.Cells[j + 2, 1] = clientName;
                            //    xlWorkSheet.Cells[j + 2, 2] = listSavedPdf[j].Status;
                            //}
                            //xlApp.DisplayAlerts = false;
                            //xlWorkBook.SaveAs(textBox3.Text + "\\Tracking.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                            //xlWorkBook.Close(true, misValue, misValue);
                            //xlApp.Quit();
                            driver.Quit();
                        }
                        else if (rowcount >= 2)
                        {
                            //table ctl00_middle_grvSearchResults
                            element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults_ctl03_chkbxSelect"));
                            Thread.Sleep(2330);
                            element.Click();
                            //clicking on Done Button
                            driver.FindElement(By.Id("ctl00_middle_btnDoneBottom")).Click();


                            driver.FindElement(By.Id("ctl00_middle_txtLegalName")).SendKeys(listThreadDocument[i].Legalname);
                            if (listThreadDocument[i].AccountNumber != null)
                                driver.FindElement(By.Id("ctl00_middle_txtEmployerID")).SendKeys(listThreadDocument[i].AccountNumber);
                            driver.FindElement(By.Id("ctl00_middle_addAddress__ddlCountry")).SendKeys(listThreadDocument[i].Country);
                            driver.FindElement(By.Id("ctl00_middle_addAddress_txtAddress1")).SendKeys(listThreadDocument[i].Address);
                            driver.FindElement(By.Id("ctl00_middle_addAddress_txtCity")).SendKeys(listThreadDocument[i].City);
                            driver.FindElement(By.Id("ctl00_middle_addAddress_ddlProvince")).SendKeys(listThreadDocument[i].Province);
                            driver.FindElement(By.Id("ctl00_middle_addAddress_txtPostalCode")).SendKeys(listThreadDocument[i].Postalcode);
                            driver.FindElement(By.Id("ctl00_middle_phnPhone_txtPhone")).SendKeys(listThreadDocument[i].Phonenumber);
                            if (listThreadDocument[i].Ext != null)
                                driver.FindElement(By.Id("ctl00_middle_phnPhone_txtext")).SendKeys(listThreadDocument[i].Ext);
                            if (listThreadDocument[i].Emailaddress != null)
                                driver.FindElement(By.Id("ctl00_middle_emlEmail__txtEmail1")).SendKeys(listThreadDocument[i].Emailaddress);


                            Thread.Sleep(2330);
                            driver.FindElement(By.Id("ctl00_middle_btnCreateBottom")).Click();
                            Thread.Sleep(2330);
                            driver.FindElement(By.Id("ctl00_middle_btnView")).Click();
                            Thread.Sleep(2330);

                            SendKeys.SendWait("^s");  // send control+s
                            Thread.Sleep(2330);

                            // string clientName = listThreadDocument[i].ClientCode + '_' + listThreadDocument[i].Legalname_Tradename;
                            SendKeys.SendWait(clientName + ".pdf{ENTER}"); // sends "fileName then enter

                            string fileName = clientName + ".pdf";
                            string sourcePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                            string targetPath = textBox3.Text;

                            // Use Path class to manipulate file and directory paths.
                            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                            string destFile = System.IO.Path.Combine(targetPath, fileName);
                            Thread.Sleep(2330);

                            // To copy a file to another location and 
                            // overwrite the destination file if it already exists.
                            if (!File.Exists(destFile))
                            {
                                System.IO.File.Move(sourceFile, destFile);
                            }


                            Document obj = Document.getDocumentObj();
                            obj.Legalname = clientName;
                            obj.Status = "Success";
                            saveExcel(obj);
                            //listSavedPdf.Add(obj);

                            //listSavedPdf = listSavedPdf.GroupBy(x => x.Legalname).Select(y => y.First()).ToList();


                            //xlWorkSheet.Cells[1, 1] = "Legalname";
                            //xlWorkSheet.Cells[1, 2] = "Status";
                            //for (int j = 0; j < listSavedPdf.Count; j++)
                            //{
                            //    xlWorkSheet.Cells[j + 2, 1] = clientName;
                            //    xlWorkSheet.Cells[j + 2, 2] = listSavedPdf[j].Status;
                            //}
                            //xlApp.DisplayAlerts = false;
                            //xlWorkBook.SaveAs(textBox3.Text + "\\Tracking.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                            //xlWorkBook.Close(true, misValue, misValue);
                            //xlApp.Quit();
                            driver.Quit();
                        }
                    }
                }
            }
            //}
        }

        private void button3_Click(object sender, EventArgs e)
        {

            int numberOFTurns = int.Parse(comboBox1.SelectedItem.ToString());
            int numberOfDocuments = listDocuments.Count();
            int numberofDocumentsSelected = 0;

            List<Thread> listThread = new List<Thread>();

           
            //Checking conditions
            if (numberOFTurns >= numberOfDocuments)
            {

                for (int i = 0; i < numberOFTurns; i++)
                {

                    List<Document> listSelected = listDocuments.Take(i + 1).ToList();
                    Thread obj = new Thread(() => startingThread(listSelected));
                    listDocuments.RemoveAll(x => listSelected.Contains(x));
                    obj.Start();
                    // obj.Join();
                }
            }
            else
            {
                numberofDocumentsSelected = int.Parse(Math.Round(Convert.ToDecimal(numberOfDocuments / numberOFTurns)).ToString());

                for (int i = 0; i < numberOFTurns; i++)
                {
                    List<Document> listSelected = listDocuments.Take(numberofDocumentsSelected).ToList();
                    Thread obj = new Thread(() => startingThread(listSelected));
                    listDocuments.RemoveAll(x => listSelected.Contains(x));
                    obj.Start();
                    //  obj.Join();
                    //  obj.Abort();
                }
            }

            //Thread finalThread = new Thread(saveExcel);
            //Thread.Sleep(1000);
            //finalThread.Start();
            //finalThread.Join();
            // KillSpecificExcelFileProcess(ofd.SafeFileName);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            var processess = from p in Process.GetProcessesByName("chrome")
                             select p;

            foreach (var process in processes)
            {
                process.Kill();
            }
            foreach (var process in Process.GetProcessesByName("chromedriver"))
            {
                process.Kill();
            }
            //svchost
            //conhost
            //access denied

            foreach (var process in Process.GetProcessesByName("chrome"))
            {
                process.Kill();
            }
        }


        private void saveExcel(Document obj)
        {
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            listSavedPdf.Add(obj);
            listSavedPdf = listSavedPdf.GroupBy(x => x.Legalname).Select(y => y.First()).ToList();
            xlWorkSheet.Cells[1, 1] = "Legalname";
            xlWorkSheet.Cells[1, 2] = "Status";
            for (int j = 0; j < listSavedPdf.Count; j++)
            {
                xlWorkSheet.Cells[j + 2, 1] = listSavedPdf[j].Legalname;
                xlWorkSheet.Cells[j + 2, 2] = listSavedPdf[j].Status;
            }
            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(textBox3.Text + "\\Tracking.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }

        private void KillSpecificExcelFileProcess(string excelFileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                process.Kill();
            }
            foreach (var process in Process.GetProcessesByName("chromedriver"))
            {
                process.Kill();
            }
            foreach (var process in Process.GetProcessesByName("chrome"))
            {
                process.Kill();
            }
            foreach (var process in Process.GetProcessesByName("conhost"))
            {
                process.Kill();
            }

        }
    }
}
