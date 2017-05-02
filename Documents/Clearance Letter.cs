using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
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
using System.Net.Http;
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
        private List<Document> list = new List<Document>();

        public Form1()
        {
            InitializeComponent();
            RecordsCount.Enabled = false;

            List<string> iterationsList = new List<string>() { "2", "4", "6" };
            comboBox1.DataSource = iterationsList;
            Loading.Hide();
            threadCount = 0;
            this.Text = "Web Crawler-BC form name";
        }

        OpenFileDialog ofd = new OpenFileDialog();
        FirefoxProfile prof = new FirefoxProfile();
        List<Document> listDocuments = new List<Document>();
        List<Document> listSavedPdf = new List<Document>();
        List<Document> listTracking = new List<Document>();
        private int threadCount;
        private Object thisLock = new Object();

        private void ExcelBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                Loading.Minimum = 0;
                ofd.Filter = "Excel(*.xls,*.xlsx)|*.xls;*.xlsx";
              //  string currentDir = Environment.CurrentDirectory;
              //  DirectoryInfo directory = new DirectoryInfo(currentDir);
                listDocuments = new List<Document>();
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    Loading.Minimum = 0;
                    Loading.Maximum = 100;
                    Loading.Show();

                    ExcelTextbox.Text = ofd.FileName;
                    //ExcelTextbox.Text = ofd.SafeFileName;
                    Excel.Application xlApp1 = new Excel.Application();
                    Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(ofd.FileName);
                    Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
                    Excel.Range xlRange1 = xlWorksheet1.UsedRange;
                    int rowCount1 = xlRange1.Rows.Count;
                    int columnCount1 = xlRange1.Columns.Count;


                    for (int i = 2; i <= rowCount1; i++)
                    {
                        if (xlRange1.Cells[i, 2].Value2 != "" && xlRange1.Cells[i, 2].Value2 != null)
                        {
                            Document document = Document.getDocumentObj();
                            document.Legalname = xlRange1.Cells[i, 2].Value2;
                            document.AccountNumber = xlRange1.Cells[i, 1].Value2.ToString();
                            document.Country = xlRange1.Cells[i, 3].Value2;
                            document.Address = xlRange1.Cells[i, 4].Value2;
                            document.City = xlRange1.Cells[i, 5].Value2;
                            document.Province = xlRange1.Cells[i, 6].Value2;
                            document.Postalcode = xlRange1.Cells[i, 7].Value2;
                            document.Phonenumber = xlRange1.Cells[i, 8].Value2;
                            document.Ext = xlRange1.Cells[i, 9].Value2;
                            document.Emailaddress = xlRange1.Cells[i, 10].Value2;
                            document.ClientAccoutnumber = xlRange1.Cells[i, 11].Value2.ToString();
                            document.Legalname_Tradename = xlRange1.Cells[i, 12].Value2;
                            document.ClientCode = xlRange1.Cells[i, 13].Value2;
                            document.ClientCode_ClientName = xlRange1.Cells[i, 13].Value2 + '_' + xlRange1.Cells[i, 12].Value2;
                            document.ClientCode_ClientName = document.ClientCode_ClientName.ToLower();
                            listDocuments.Add(document);
                            list.Add(document);
                            var percentage = (i * 100 / rowCount1);
                            Loading.Value = percentage;
                            Loading.Update();
                        }
                    }
                    listDocuments = listDocuments.GroupBy(x => x.Legalname_Tradename + "_" + x.ClientCode).Select(y => y.First()).ToList();
                    list = listDocuments.GroupBy(x => x.Legalname_Tradename + "_" + x.ClientCode).Select(y => y.First()).ToList();

                    //listDocuments = listDocuments.GroupBy(x => x.Legalname_Tradename).Select(y => y.First()).ToList();
                    //list = listDocuments.GroupBy(x => x.Legalname_Tradename).Select(y => y.First()).ToList();
                }
                RecordsCount.Text = listDocuments.Count().ToString();
                Loading.Hide();
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DownloadPath_Click(object sender, EventArgs e)
        {
            try
            {
                //  Loading.Minimum = 0;
                FolderBrowserDialog folderDialog = new FolderBrowserDialog();
                folderDialog.Description = "Select the Folder where do you want to save documents";
                folderDialog.ShowNewFolderButton = false;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    //    Loading.Minimum = 0;
                    //  Loading.Maximum = 100;
                    //Loading.Show();



                    DownloadPath.Text = folderDialog.SelectedPath;
                    bool currentDateFolder = System.IO.Directory.Exists(folderDialog.SelectedPath + "/" + DateTime.Now.ToString("MM-dd-yyyy"));

                    if (currentDateFolder)
                    {
                        DirectoryInfo d = new DirectoryInfo(DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy"));//Assuming Test is your Folder
                        FileInfo[] Files = d.GetFiles("*.pdf"); //Getting Text files

                        foreach (FileInfo file in Files)
                        {
                            string code_clientName = file.Name.Replace(".pdf", "").ToLower();
                            Document objDocument = listDocuments.Where(x => x.ClientCode_ClientName == code_clientName).FirstOrDefault();
                            // removing Pdf Generated
                            listDocuments.Remove(objDocument);
                            list.Remove(objDocument);
                        }
                        DirectoryInfo downloadDirectory = new DirectoryInfo(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"));

                        FileInfo[] downloadFiles = downloadDirectory.GetFiles("*.pdf");
                        foreach (FileInfo file in downloadFiles)
                        {
                            string code_clientName = file.Name.Replace(".pdf", "").ToLower();
                            Document objDocument = listDocuments.Where(x => x.ClientCode_ClientName == code_clientName).FirstOrDefault();
                            // removing Pdf Generated
                            listDocuments.Remove(objDocument);
                            list.Remove(objDocument);
                        }



                        if (File.Exists(DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy") + "\\Tracking.xls"))
                        {
                            bool ExpiredPathExists = System.IO.Directory.Exists(DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy") + "/Expired");

                            if (!ExpiredPathExists)
                                System.IO.Directory.CreateDirectory(DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy") + "/Expired");


                            Excel.Application xlApp1 = new Excel.Application();
                            Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy") + "\\Tracking.xls");
                            Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
                            Excel.Range xlRange1 = xlWorksheet1.UsedRange;
                            int rowCount1 = xlRange1.Rows.Count;
                            int columnCount1 = xlRange1.Columns.Count;



                            for (int i = 2; i <= rowCount1; i++)
                            {
                                if (xlRange1.Cells[i, 2].Value2 != "" && xlRange1.Cells[i, 2].Value2 != null)
                                {
                                    if (xlRange1.Cells[i, 14].Value2 != "Successfully Downloaded" && xlRange1.Cells[i, 14].Value2 != null)
                                    {
                                        Document document = Document.getDocumentObj();
                                        document.Legalname = xlRange1.Cells[i, 2].Value2;
                                        document.AccountNumber = xlRange1.Cells[i, 1].Value2.ToString();
                                        document.Country = xlRange1.Cells[i, 3].Value2;
                                        document.Address = xlRange1.Cells[i, 4].Value2;
                                        document.City = xlRange1.Cells[i, 5].Value2;
                                        document.Province = xlRange1.Cells[i, 6].Value2;
                                        document.Postalcode = xlRange1.Cells[i, 7].Value2;
                                        document.Phonenumber = xlRange1.Cells[i, 8].Value2;
                                        document.Ext = xlRange1.Cells[i, 9].Value2;
                                        document.Emailaddress = xlRange1.Cells[i, 10].Value2;
                                        document.ClientAccoutnumber = xlRange1.Cells[i, 11].Value2.ToString();
                                        document.Legalname_Tradename = xlRange1.Cells[i, 12].Value2;
                                        document.ClientCode = xlRange1.Cells[i, 13].Value2;
                                        document.ClientCode_ClientName = xlRange1.Cells[i, 13].Value2 + '_' + xlRange1.Cells[i, 12].Value2;
                                        listTracking.Add(document);
                                        //    var percentage = (i * 100 / rowCount1);
                                        //  Loading.Value = percentage;
                                        //Loading.Update();
                                    }
                                }
                            }
                            //Loading.Hide();
                            foreach (Document document in listTracking)
                            {
                                var smallDocument = document.ClientCode_ClientName.ToLower();
                                var mainDocument = listDocuments.Where(x => x.ClientCode_ClientName == smallDocument).FirstOrDefault();
                                listDocuments.Remove(mainDocument);
                                list.Remove(mainDocument);
                            }
                            //Loading.Hide();
                            RecordsCount.Text = listDocuments.Count().ToString();
                            foreach (var process in Process.GetProcessesByName("EXCEL"))
                            {
                                process.Kill();
                            }
                        }

                    }
                    else
                    {
                        System.IO.Directory.CreateDirectory(folderDialog.SelectedPath + "/" + DateTime.Now.ToString("MM-dd-yyyy"));
                        bool ExpiredPathExists = System.IO.Directory.Exists(DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy") + "/Expired");

                        if (!ExpiredPathExists)
                            System.IO.Directory.CreateDirectory(DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy") + "/Expired");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool IsElementPresent(IWebDriver driver, By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        private void startingThread(List<Document> listThreadDocument)
        {
            try
            {
                if (listThreadDocument.Count > 0)
                {
                    threadCount = threadCount + 1;
                    var specialcharacters = ConfigurationManager.AppSettings["SpecialChar"];

                    DateTime date = new DateTime();


                    for (int i = 0; i < listThreadDocument.Count(); i++)
                    {
                        Document threadDocument = Document.getDocumentObj();
                        threadDocument = listThreadDocument[i];
                        if (threadDocument.Legalname_Tradename != null && threadDocument.Legalname_Tradename != "")
                        {
                            //listThreadDocument[i].Legalname = listThreadDocument[i].Legalname.Trim();
                            //listThreadDocument[i].Legalname_Tradename = Regex.Replace(listThreadDocument[i].Legalname_Tradename, @"[^0-9a-zA-Z]+", specialcharacters);
                            // listThreadDocument[i].Legalname_Tradename = Regex.Replace(listThreadDocument[i].Legalname_Tradename, @"[^0-9a-zA-Z]+", specialcharacters);
                            var chromeOptions = new ChromeOptions();
                            chromeOptions.AddUserProfilePreference("download.default_directory", DownloadPath.Text);
                            chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
                            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                            IWebDriver driver = new ChromeDriver(chromeOptions);
                            IWebElement element;
                            //Direct Search Page
                            driver.Navigate().GoToUrl("https://online.worksafebc.com/Anonymous/EmployerClearanceLetter/search.aspx");
                            Thread.Sleep(950);

                            //if (threadDocument.ClientAccoutnumber != null && threadDocument.ClientAccoutnumber != "Not Available")
                            //{
                            //    driver.FindElement(By.Id("ctl00_middle_radSearchTypeAccount")).Click();//
                            //    Thread.Sleep(1500);
                            //    element = driver.FindElement(By.Id("ctl00_middle_txtAccountNumber"));
                            //    element.SendKeys(threadDocument.ClientAccoutnumber);
                            //    //ctl00_middle_btnAccountSearch
                            //    Thread.Sleep(1500);
                            //    driver.FindElement(By.Id("ctl00_middle_btnAccountSearch")).Click();
                            //}
                            //else
                            //{

                            driver.FindElement(By.Id("ctl00_middle_radSearchTypeName")).Click();//ctl00_middle_txtLegalName
                            Thread.Sleep(900);
                            element = driver.FindElement(By.Id("ctl00_middle_txtLegalName"));
                            element.SendKeys(threadDocument.Legalname_Tradename);
                            Thread.Sleep(900);
                            driver.FindElement(By.Id("ctl00_middle_btnNameSearch")).Click();
                            //  }
                            Thread.Sleep(900);


                            lock (thisLock)
                            {
                                element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults"));
                                var rowcount = element.FindElements(By.TagName("tr")).Count();
                                var noRecords = driver.FindElements(By.CssSelector("table tr"));
                                if (noRecords[5].Text.ToString() == "No firms found.")
                                {
                                    //lock (thisLock)
                                    //{
                                    Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                    documentObj.Status = "No Records Found";
                                    if (documentObj != null)
                                    {
                                        list.Remove(documentObj);
                                        list.Add(documentObj);
                                        listTracking.Remove(documentObj);
                                        listTracking.Add(documentObj);
                                    }
                                    //  saveExcel(list);
                                    Thread saveThreadexcel = new Thread(() => savingPDF(documentObj.ClientCode_ClientName, documentObj.ClientCode));
                                    saveThreadexcel.Start();
                                    driver.Quit();
                                    //}
                                }
                                else if (rowcount > 2)
                                {
                                    //lock (thisLock)
                                    //{
                                    Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                    documentObj.Status = "Multiple Records";
                                    if (documentObj != null)
                                    {
                                        list.Remove(documentObj);
                                        list.Add(documentObj);
                                        listTracking.Remove(documentObj);
                                        listTracking.Add(documentObj);
                                    }
                                    //   saveExcel(list);
                                    Thread saveThreadexcel = new Thread(() => savingPDF(documentObj.ClientCode_ClientName, documentObj.ClientCode));
                                    saveThreadexcel.Start();
                                    driver.Quit();
                                    // }
                                }
                                else if (rowcount == 2)
                                {
                                    //lock (thisLock)
                                    //{
                                    string clientNamess = threadDocument.ClientCode.ToUpper() + '_' + threadDocument.Legalname_Tradename.ToUpper();
                                    clientNamess = clientNamess.ToUpper();
                                    var accountNumber = noRecords[5].Text.Split(' ')[0];
                                    //if (threadDocument.ClientAccoutnumber != null && threadDocument.ClientAccoutnumber != "Not Available")
                                    //{
                                    //    element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults_ctl03_chkbxSelect"));
                                    //    driver.FindElement(By.Id("ctl00_middle_btnDoneBottom")).Click();
                                    //}
                                    //else
                                    //{
                                    element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults_ctl03_chkbxSelect"));
                                    Thread.Sleep(1000);
                                    element.Click();
                                    //clicking on Done Button
                                    driver.FindElement(By.Id("ctl00_middle_btnDoneBottom")).Click();
                                    //  }

                                    bool isElementDisplayed = IsElementPresent(driver, By.Id("ctl00_middle_txtLegalName"));//driver.FindElements(By.Id("ctl00_middle_txtLegalName"));
                                    if (isElementDisplayed)
                                    {
                                        driver.FindElement(By.Id("ctl00_middle_txtLegalName")).SendKeys(threadDocument.Legalname);
                                        if (threadDocument.AccountNumber != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_txtEmployerID")).SendKeys(threadDocument.AccountNumber);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress__ddlCountry")).SendKeys(threadDocument.Country);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtAddress1")).SendKeys(threadDocument.Address);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtCity")).SendKeys(threadDocument.City);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_ddlProvince")).SendKeys(threadDocument.Province);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtPostalCode")).SendKeys(threadDocument.Postalcode);
                                        driver.FindElement(By.Id("ctl00_middle_phnPhone_txtPhone")).SendKeys(threadDocument.Phonenumber);
                                        if (threadDocument.Ext != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_phnPhone_txtext")).SendKeys(threadDocument.Ext);
                                        if (threadDocument.Emailaddress != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_emlEmail__txtEmail1")).SendKeys(threadDocument.Emailaddress);

                                        Thread.Sleep(900);
                                        string fileName = "";

                                        //wcbProcessScreen
                                        driver.FindElement(By.Id("ctl00_middle_btnCreateBottom")).Click();
                                        Thread.Sleep(1500);
                                        //ctl00_middle_btnView
                                        driver.FindElement(By.Id("ctl00_middle_btnView")).Click();
                                        Thread.Sleep(2300);

                                        SendKeys.SendWait("^s");  // send control+s
                                        Thread.Sleep(1500);
                                        clientNamess = clientNamess.ToUpper();
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait(clientNamess + ".pdf{ENTER}"); // sends "fileName then enter
                                        fileName = clientNamess + ".pdf";
                                        //}
                                        Thread.Sleep(2000);

                                        Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                        documentObj.ClientAccoutnumber = accountNumber;
                                        documentObj.Status = "Successfully Downloaded";
                                        documentObj.ClearanceDate = date.ToString("MM-dd-yyyy");

                                        if (documentObj != null)
                                        {
                                            list.Remove(documentObj);
                                            list.Add(documentObj);
                                            listTracking.Remove(documentObj);
                                            listTracking.Add(documentObj);
                                        }
                                        Thread.Sleep(900);
                                        driver.Quit();


                                        // Thread saveThreadexcel = new Thread(savingPDF);
                                        Thread saveThreadexcel = new Thread(() => savingPDF(clientNamess, documentObj.ClientCode));
                                        saveThreadexcel.Start();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void startingThread2(List<Document> listThreadDocument)
        {
            try
            {
                if (listThreadDocument.Count > 0)
                {
                    threadCount = threadCount + 1;
                    var specialcharacters = ConfigurationManager.AppSettings["SpecialChar"];

                    DateTime date = new DateTime();


                    for (int i = 0; i < listThreadDocument.Count(); i++)
                    {
                        Document threadDocument = Document.getDocumentObj();
                        threadDocument = listThreadDocument[i];
                        if (threadDocument.Legalname_Tradename != null && threadDocument.Legalname_Tradename != "")
                        {
                            //listThreadDocument[i].Legalname = listThreadDocument[i].Legalname.Trim();
                            //listThreadDocument[i].Legalname_Tradename = Regex.Replace(listThreadDocument[i].Legalname_Tradename, @"[^0-9a-zA-Z]+", specialcharacters);
                            // listThreadDocument[i].Legalname_Tradename = Regex.Replace(listThreadDocument[i].Legalname_Tradename, @"[^0-9a-zA-Z]+", specialcharacters);
                            var chromeOptions = new ChromeOptions();
                            chromeOptions.AddUserProfilePreference("download.default_directory", DownloadPath.Text);
                            chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
                            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                            IWebDriver driver = new ChromeDriver(chromeOptions);
                            IWebElement element;
                            //Direct Search Page
                            driver.Navigate().GoToUrl("https://online.worksafebc.com/Anonymous/EmployerClearanceLetter/search.aspx");
                            Thread.Sleep(950);

                            //if (threadDocument.ClientAccoutnumber != null && threadDocument.ClientAccoutnumber != "Not Available")
                            //{
                            //    driver.FindElement(By.Id("ctl00_middle_radSearchTypeAccount")).Click();//
                            //    Thread.Sleep(1500);
                            //    element = driver.FindElement(By.Id("ctl00_middle_txtAccountNumber"));
                            //    element.SendKeys(threadDocument.ClientAccoutnumber);
                            //    //ctl00_middle_btnAccountSearch
                            //    Thread.Sleep(1500);
                            //    driver.FindElement(By.Id("ctl00_middle_btnAccountSearch")).Click();
                            //}
                            //else
                            //{

                            driver.FindElement(By.Id("ctl00_middle_radSearchTypeName")).Click();//ctl00_middle_txtLegalName
                            Thread.Sleep(900);
                            element = driver.FindElement(By.Id("ctl00_middle_txtLegalName"));
                            element.SendKeys(threadDocument.Legalname_Tradename);
                            Thread.Sleep(900);
                            driver.FindElement(By.Id("ctl00_middle_btnNameSearch")).Click();
                            //  }
                            Thread.Sleep(900);


                            lock (thisLock)
                            {
                                element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults"));
                                var rowcount = element.FindElements(By.TagName("tr")).Count();
                                var noRecords = driver.FindElements(By.CssSelector("table tr"));
                                if (noRecords[5].Text.ToString() == "No firms found.")
                                {
                                    //lock (thisLock)
                                    //{
                                    Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                    documentObj.Status = "No Records Found";
                                    if (documentObj != null)
                                    {
                                        list.Remove(documentObj);
                                        list.Add(documentObj);
                                        listTracking.Remove(documentObj);
                                        listTracking.Add(documentObj);
                                    }
                                    //  saveExcel(list);
                                    Thread saveThreadexcel = new Thread(() => savingPDF(documentObj.ClientCode_ClientName, documentObj.ClientCode));
                                    saveThreadexcel.Start();
                                    driver.Quit();
                                    //}
                                }
                                else if (rowcount > 2)
                                {
                                    //lock (thisLock)
                                    //{
                                    Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                    documentObj.Status = "Multiple Records";
                                    if (documentObj != null)
                                    {
                                        list.Remove(documentObj);
                                        list.Add(documentObj);
                                        listTracking.Remove(documentObj);
                                        listTracking.Add(documentObj);
                                    }
                                    //   saveExcel(list);
                                    Thread saveThreadexcel = new Thread(() => savingPDF(documentObj.ClientCode_ClientName, documentObj.ClientCode));
                                    saveThreadexcel.Start();
                                    driver.Quit();
                                    // }
                                }
                                else if (rowcount == 2)
                                {
                                    //lock (thisLock)
                                    //{
                                    string clientNamess = threadDocument.ClientCode.ToUpper() + '_' + threadDocument.Legalname_Tradename.ToUpper();
                                    clientNamess = clientNamess.ToUpper();
                                    var accountNumber = noRecords[5].Text.Split(' ')[0];
                                    //if (threadDocument.ClientAccoutnumber != null && threadDocument.ClientAccoutnumber != "Not Available")
                                    //{
                                    //    element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults_ctl03_chkbxSelect"));
                                    //    driver.FindElement(By.Id("ctl00_middle_btnDoneBottom")).Click();
                                    //}
                                    //else
                                    //{
                                    element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults_ctl03_chkbxSelect"));
                                    Thread.Sleep(1000);
                                    element.Click();
                                    //clicking on Done Button
                                    driver.FindElement(By.Id("ctl00_middle_btnDoneBottom")).Click();
                                    //  }

                                    bool isElementDisplayed = IsElementPresent(driver, By.Id("ctl00_middle_txtLegalName"));//driver.FindElements(By.Id("ctl00_middle_txtLegalName"));
                                    if (isElementDisplayed)
                                    {
                                        driver.FindElement(By.Id("ctl00_middle_txtLegalName")).SendKeys(threadDocument.Legalname);
                                        if (threadDocument.AccountNumber != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_txtEmployerID")).SendKeys(threadDocument.AccountNumber);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress__ddlCountry")).SendKeys(threadDocument.Country);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtAddress1")).SendKeys(threadDocument.Address);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtCity")).SendKeys(threadDocument.City);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_ddlProvince")).SendKeys(threadDocument.Province);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtPostalCode")).SendKeys(threadDocument.Postalcode);
                                        driver.FindElement(By.Id("ctl00_middle_phnPhone_txtPhone")).SendKeys(threadDocument.Phonenumber);
                                        if (threadDocument.Ext != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_phnPhone_txtext")).SendKeys(threadDocument.Ext);
                                        if (threadDocument.Emailaddress != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_emlEmail__txtEmail1")).SendKeys(threadDocument.Emailaddress);

                                        Thread.Sleep(900);
                                        string fileName = "";

                                        //wcbProcessScreen
                                        driver.FindElement(By.Id("ctl00_middle_btnCreateBottom")).Click();
                                        Thread.Sleep(1500);
                                        //ctl00_middle_btnView
                                        driver.FindElement(By.Id("ctl00_middle_btnView")).Click();
                                        Thread.Sleep(2300);

                                        SendKeys.SendWait("^s");  // send control+s
                                        Thread.Sleep(1500);
                                        clientNamess = clientNamess.ToUpper();
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait(clientNamess + ".pdf{ENTER}"); // sends "fileName then enter
                                        fileName = clientNamess + ".pdf";
                                        //}
                                        Thread.Sleep(2000);

                                        Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                        documentObj.ClientAccoutnumber = accountNumber;
                                        documentObj.Status = "Successfully Downloaded";
                                        documentObj.ClearanceDate = date.ToString("MM-dd-yyyy");

                                        if (documentObj != null)
                                        {
                                            list.Remove(documentObj);
                                            list.Add(documentObj);
                                            listTracking.Remove(documentObj);
                                            listTracking.Add(documentObj);
                                        }
                                        Thread.Sleep(900);
                                        driver.Quit();


                                        // Thread saveThreadexcel = new Thread(savingPDF);
                                        Thread saveThreadexcel = new Thread(() => savingPDF(clientNamess, documentObj.ClientCode));
                                        saveThreadexcel.Start();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void startingThread3(List<Document> listThreadDocument)
        {
            try
            {
                if (listThreadDocument.Count > 0)
                {
                    threadCount = threadCount + 1;
                    var specialcharacters = ConfigurationManager.AppSettings["SpecialChar"];

                    DateTime date = new DateTime();
                    int finalThread = 0;

                    for (int i = 0; i < listThreadDocument.Count(); i++)
                    {
                        Document threadDocument = Document.getDocumentObj();
                        threadDocument = listThreadDocument[i];
                        if (threadDocument.Legalname_Tradename != null && threadDocument.Legalname_Tradename != "")
                        {
                            //listThreadDocument[i].Legalname = listThreadDocument[i].Legalname.Trim();
                            //listThreadDocument[i].Legalname_Tradename = Regex.Replace(listThreadDocument[i].Legalname_Tradename, @"[^0-9a-zA-Z]+", specialcharacters);
                            // listThreadDocument[i].Legalname_Tradename = Regex.Replace(listThreadDocument[i].Legalname_Tradename, @"[^0-9a-zA-Z]+", specialcharacters);
                            var chromeOptions = new ChromeOptions();
                            chromeOptions.AddUserProfilePreference("download.default_directory", DownloadPath.Text);
                            chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
                            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                            IWebDriver driver = new ChromeDriver(chromeOptions);
                            IWebElement element;
                            //Direct Search Page
                            driver.Navigate().GoToUrl("https://online.worksafebc.com/Anonymous/EmployerClearanceLetter/search.aspx");
                            Thread.Sleep(950);

                            //if (threadDocument.ClientAccoutnumber != null && threadDocument.ClientAccoutnumber != "Not Available")
                            //{
                            //    driver.FindElement(By.Id("ctl00_middle_radSearchTypeAccount")).Click();//
                            //    Thread.Sleep(1500);
                            //    element = driver.FindElement(By.Id("ctl00_middle_txtAccountNumber"));
                            //    element.SendKeys(threadDocument.ClientAccoutnumber);
                            //    //ctl00_middle_btnAccountSearch
                            //    Thread.Sleep(1500);
                            //    driver.FindElement(By.Id("ctl00_middle_btnAccountSearch")).Click();
                            //}
                            //else
                            //{

                            driver.FindElement(By.Id("ctl00_middle_radSearchTypeName")).Click();//ctl00_middle_txtLegalName
                            Thread.Sleep(900);
                            element = driver.FindElement(By.Id("ctl00_middle_txtLegalName"));
                            element.SendKeys(threadDocument.Legalname_Tradename);
                            Thread.Sleep(900);
                            driver.FindElement(By.Id("ctl00_middle_btnNameSearch")).Click();
                            //  }
                            Thread.Sleep(900);

                            finalThread = finalThread + 1;
                            lock (thisLock)
                            {
                                element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults"));
                                var rowcount = element.FindElements(By.TagName("tr")).Count();
                                var noRecords = driver.FindElements(By.CssSelector("table tr"));
                                if (noRecords[5].Text.ToString() == "No firms found.")
                                {
                                    //lock (thisLock)
                                    //{
                                    Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                    documentObj.Status = "No Records Found";
                                    if (documentObj != null)
                                    {
                                        list.Remove(documentObj);
                                        list.Add(documentObj);
                                        listTracking.Remove(documentObj);
                                        listTracking.Add(documentObj);
                                    }
                                    //  saveExcel(list);
                                    //Thread saveThreadexcel = new Thread(() => savingPDF(documentObj.ClientCode_ClientName, documentObj.ClientCode));
                                    //saveThreadexcel.Start();
                                    driver.Quit();
                                    //}
                                }
                                else if (rowcount > 2)
                                {
                                    //lock (thisLock)
                                    //{
                                    Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                    documentObj.Status = "Multiple Records";
                                    if (documentObj != null)
                                    {
                                        list.Remove(documentObj);
                                        list.Add(documentObj);
                                        listTracking.Remove(documentObj);
                                        listTracking.Add(documentObj);
                                    }
                                    //   saveExcel(list);
                                    //Thread saveThreadexcel = new Thread(() => savingPDF(documentObj.ClientCode_ClientName, documentObj.ClientCode));
                                    //saveThreadexcel.Start();
                                    driver.Quit();
                                    // }
                                }
                                else if (rowcount == 2)
                                {
                                    //lock (thisLock)
                                    //{
                                    string clientNamess = threadDocument.ClientCode.ToUpper() + '_' + threadDocument.Legalname_Tradename.ToUpper();
                                    clientNamess = clientNamess.ToUpper();
                                    var accountNumber = noRecords[5].Text.Split(' ')[0];
                                    //if (threadDocument.ClientAccoutnumber != null && threadDocument.ClientAccoutnumber != "Not Available")
                                    //{
                                    //    element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults_ctl03_chkbxSelect"));
                                    //    driver.FindElement(By.Id("ctl00_middle_btnDoneBottom")).Click();
                                    //}
                                    //else
                                    //{
                                    element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults_ctl03_chkbxSelect"));
                                    Thread.Sleep(1000);
                                    element.Click();
                                    //clicking on Done Button
                                    driver.FindElement(By.Id("ctl00_middle_btnDoneBottom")).Click();
                                    //  }

                                    bool isElementDisplayed = IsElementPresent(driver, By.Id("ctl00_middle_txtLegalName"));//driver.FindElements(By.Id("ctl00_middle_txtLegalName"));
                                    if (isElementDisplayed)
                                    {
                                        driver.FindElement(By.Id("ctl00_middle_txtLegalName")).SendKeys(threadDocument.Legalname);
                                        if (threadDocument.AccountNumber != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_txtEmployerID")).SendKeys(threadDocument.AccountNumber);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress__ddlCountry")).SendKeys(threadDocument.Country);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtAddress1")).SendKeys(threadDocument.Address);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtCity")).SendKeys(threadDocument.City);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_ddlProvince")).SendKeys(threadDocument.Province);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtPostalCode")).SendKeys(threadDocument.Postalcode);
                                        driver.FindElement(By.Id("ctl00_middle_phnPhone_txtPhone")).SendKeys(threadDocument.Phonenumber);
                                        if (threadDocument.Ext != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_phnPhone_txtext")).SendKeys(threadDocument.Ext);
                                        if (threadDocument.Emailaddress != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_emlEmail__txtEmail1")).SendKeys(threadDocument.Emailaddress);

                                        Thread.Sleep(900);
                                        string fileName = "";

                                        //wcbProcessScreen
                                        driver.FindElement(By.Id("ctl00_middle_btnCreateBottom")).Click();
                                        Thread.Sleep(1500);
                                        //ctl00_middle_btnView
                                        driver.FindElement(By.Id("ctl00_middle_btnView")).Click();
                                        Thread.Sleep(2300);

                                        SendKeys.SendWait("^s");  // send control+s
                                        Thread.Sleep(1500);
                                        clientNamess = clientNamess.ToUpper();
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait(clientNamess + ".pdf{ENTER}"); // sends "fileName then enter
                                        fileName = clientNamess + ".pdf";
                                        //}
                                        Thread.Sleep(2000);

                                        Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                        documentObj.ClientAccoutnumber = accountNumber;
                                        documentObj.Status = "Successfully Downloaded";
                                        documentObj.ClearanceDate = date.ToString("MM-dd-yyyy");

                                        if (documentObj != null)
                                        {
                                            list.Remove(documentObj);
                                            list.Add(documentObj);
                                            listTracking.Remove(documentObj);
                                            listTracking.Add(documentObj);
                                        }
                                        Thread.Sleep(900);
                                        driver.Quit();


                                        //// Thread saveThreadexcel = new Thread(savingPDF);
                                        //Thread saveThreadexcel = new Thread(() => savingPDF(clientNamess, documentObj.ClientCode));
                                        //saveThreadexcel.Start();
                                    }
                                }
                            }
                        }
                    }

                    if (finalThread == listThreadDocument.Count())
                    {
                        savingPDFAllOnce();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void startingFinalThread(List<Document> listThreadDocument)
        {
            try
            {
                if (listThreadDocument.Count > 0)
                {
                    threadCount = threadCount + 1;
                    var specialcharacters = ConfigurationManager.AppSettings["SpecialChar"];

                    DateTime date = new DateTime();


                    for (int i = 0; i < listThreadDocument.Count(); i++)
                    {
                        Document threadDocument = Document.getDocumentObj();
                        threadDocument = listThreadDocument[i];
                        if (threadDocument.Legalname_Tradename != null && threadDocument.Legalname_Tradename != "")
                        {
                            //listThreadDocument[i].Legalname = listThreadDocument[i].Legalname.Trim();
                            //listThreadDocument[i].Legalname_Tradename = Regex.Replace(listThreadDocument[i].Legalname_Tradename, @"[^0-9a-zA-Z]+", specialcharacters);
                            // listThreadDocument[i].Legalname_Tradename = Regex.Replace(listThreadDocument[i].Legalname_Tradename, @"[^0-9a-zA-Z]+", specialcharacters);
                            var chromeOptions = new ChromeOptions();
                            chromeOptions.AddUserProfilePreference("download.default_directory", DownloadPath.Text);
                            chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
                            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                            IWebDriver driver = new ChromeDriver(chromeOptions);
                            IWebElement element;
                            //Direct Search Page
                            driver.Navigate().GoToUrl("https://online.worksafebc.com/Anonymous/EmployerClearanceLetter/search.aspx");
                            Thread.Sleep(950);

                            //if (threadDocument.ClientAccoutnumber != null && threadDocument.ClientAccoutnumber != "Not Available")
                            //{
                            //    driver.FindElement(By.Id("ctl00_middle_radSearchTypeAccount")).Click();//
                            //    Thread.Sleep(1500);
                            //    element = driver.FindElement(By.Id("ctl00_middle_txtAccountNumber"));
                            //    element.SendKeys(threadDocument.ClientAccoutnumber);
                            //    //ctl00_middle_btnAccountSearch
                            //    Thread.Sleep(1500);
                            //    driver.FindElement(By.Id("ctl00_middle_btnAccountSearch")).Click();
                            //}
                            //else
                            //{

                            driver.FindElement(By.Id("ctl00_middle_radSearchTypeName")).Click();//ctl00_middle_txtLegalName
                            Thread.Sleep(900);
                            element = driver.FindElement(By.Id("ctl00_middle_txtLegalName"));
                            element.SendKeys(threadDocument.Legalname_Tradename);
                            Thread.Sleep(900);
                            driver.FindElement(By.Id("ctl00_middle_btnNameSearch")).Click();
                            //  }
                            Thread.Sleep(900);


                            lock (thisLock)
                            {
                                element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults"));
                                var rowcount = element.FindElements(By.TagName("tr")).Count();
                                var noRecords = driver.FindElements(By.CssSelector("table tr"));
                                if (noRecords[5].Text.ToString() == "No firms found.")
                                {
                                    //lock (thisLock)
                                    //{
                                    Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                    documentObj.Status = "No Records Found";
                                    if (documentObj != null)
                                    {
                                        list.Remove(documentObj);
                                        list.Add(documentObj);
                                        listTracking.Remove(documentObj);
                                        listTracking.Add(documentObj);
                                    }
                                    //  saveExcel(list);
                                    //Thread saveThreadexcel = new Thread(() => savingPDF(documentObj.ClientCode_ClientName, documentObj.ClientCode));
                                    //saveThreadexcel.Start();
                                    driver.Quit();
                                    //}
                                }
                                else if (rowcount > 2)
                                {
                                    //lock (thisLock)
                                    //{
                                    Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                    documentObj.Status = "Multiple Records";
                                    if (documentObj != null)
                                    {
                                        list.Remove(documentObj);
                                        list.Add(documentObj);
                                        listTracking.Remove(documentObj);
                                        listTracking.Add(documentObj);
                                    }
                                    //   saveExcel(list);
                                    //Thread saveThreadexcel = new Thread(() => savingPDF(documentObj.ClientCode_ClientName, documentObj.ClientCode));
                                    //saveThreadexcel.Start();
                                    driver.Quit();
                                    // }
                                }
                                else if (rowcount == 2)
                                {
                                    //lock (thisLock)
                                    //{
                                    string clientNamess = threadDocument.ClientCode.ToUpper() + '_' + threadDocument.Legalname_Tradename.ToUpper();
                                    clientNamess = clientNamess.ToUpper();
                                    var accountNumber = noRecords[5].Text.Split(' ')[0];
                                    //if (threadDocument.ClientAccoutnumber != null && threadDocument.ClientAccoutnumber != "Not Available")
                                    //{
                                    //    element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults_ctl03_chkbxSelect"));
                                    //    driver.FindElement(By.Id("ctl00_middle_btnDoneBottom")).Click();
                                    //}
                                    //else
                                    //{
                                    element = driver.FindElement(By.Id("ctl00_middle_grvSearchResults_ctl03_chkbxSelect"));
                                    Thread.Sleep(1000);
                                    element.Click();
                                    //clicking on Done Button
                                    driver.FindElement(By.Id("ctl00_middle_btnDoneBottom")).Click();
                                    //  }

                                    bool isElementDisplayed = IsElementPresent(driver, By.Id("ctl00_middle_txtLegalName"));//driver.FindElements(By.Id("ctl00_middle_txtLegalName"));
                                    if (isElementDisplayed)
                                    {
                                        driver.FindElement(By.Id("ctl00_middle_txtLegalName")).SendKeys(threadDocument.Legalname);
                                        if (threadDocument.AccountNumber != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_txtEmployerID")).SendKeys(threadDocument.AccountNumber);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress__ddlCountry")).SendKeys(threadDocument.Country);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtAddress1")).SendKeys(threadDocument.Address);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtCity")).SendKeys(threadDocument.City);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_ddlProvince")).SendKeys(threadDocument.Province);
                                        driver.FindElement(By.Id("ctl00_middle_addAddress_txtPostalCode")).SendKeys(threadDocument.Postalcode);
                                        driver.FindElement(By.Id("ctl00_middle_phnPhone_txtPhone")).SendKeys(threadDocument.Phonenumber);
                                        if (threadDocument.Ext != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_phnPhone_txtext")).SendKeys(threadDocument.Ext);
                                        if (threadDocument.Emailaddress != null && threadDocument.Ext != "Not Available")
                                            driver.FindElement(By.Id("ctl00_middle_emlEmail__txtEmail1")).SendKeys(threadDocument.Emailaddress);

                                        Thread.Sleep(900);
                                        string fileName = "";

                                        //wcbProcessScreen
                                        driver.FindElement(By.Id("ctl00_middle_btnCreateBottom")).Click();
                                        Thread.Sleep(1500);
                                        //ctl00_middle_btnView
                                        driver.FindElement(By.Id("ctl00_middle_btnView")).Click();
                                        Thread.Sleep(2300);

                                        SendKeys.SendWait("^s");  // send control+s
                                        Thread.Sleep(1500);
                                        clientNamess = clientNamess.ToUpper();
                                        Thread.Sleep(1000);
                                        SendKeys.SendWait(clientNamess + ".pdf{ENTER}"); // sends "fileName then enter
                                        fileName = clientNamess + ".pdf";
                                        //}
                                        Thread.Sleep(2000);

                                        Document documentObj = list.Where(x => x.ClientCode == threadDocument.ClientCode).FirstOrDefault();
                                        documentObj.ClientAccoutnumber = accountNumber;
                                        documentObj.Status = "Successfully Downloaded";
                                        documentObj.ClearanceDate = date.ToString("MM-dd-yyyy");

                                        if (documentObj != null)
                                        {
                                            list.Remove(documentObj);
                                            list.Add(documentObj);
                                            listTracking.Remove(documentObj);
                                            listTracking.Add(documentObj);
                                        }
                                        Thread.Sleep(900);
                                        driver.Quit();


                                        // Thread saveThreadexcel = new Thread(savingPDF);
                                        //Thread saveThreadexcel = new Thread(() => savingPDF(clientNamess, documentObj.ClientCode));
                                        //saveThreadexcel.Start();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void savingPDF(string ClientCode_ClientName, string ClientCode)
        {
            ////Saving and moving file code  starts
            string sourcePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            string targetPath = DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy");
            DirectoryInfo d = new DirectoryInfo(sourcePath);//Assuming Test is your Folder

            //Getting Text files
            string sourceFile = "";
            string dateformat = "";
            int clearance_Index = 0;
            string destFile = "";
            string ClientName = ClientCode_ClientName.ToLower();
            string clientCode = ClientCode;

            Document obj = list.Where(x => x.ClientCode_ClientName == ClientName).FirstOrDefault();

            if (obj != null)
            {
                if (obj.Status != "Successfully Downloaded" && obj.Status != "Clearance Expired")
                {
                    saveMainExcel_AccountNumbers(list);
                    saveExcel(listTracking);
                }
                // ClearanceAdded = fileName.Replace(".pdf", "  Clearance Valid Until " + date.ToString("MM-dd-yyyy") + ".pdf");
                sourceFile = System.IO.Path.Combine(sourcePath, ClientName.ToUpper()).Replace("\\", "/") + ".pdf";
                destFile = System.IO.Path.Combine(targetPath, ClientName.ToUpper()).Replace("\\", "/") + ".pdf";
                // To copy a file to another location and 
                // overwrite the destination file if it already exists.

                if (File.Exists(sourceFile))
                {
                    if (!File.Exists(destFile))
                    {
                        System.IO.File.Move(sourceFile, destFile);
                    }

                    //  obj.Legalname_Tradename = filename;
                    obj.ClientCode = clientCode;
                    obj.Status = "Successfully Downloaded";
                    listSavedPdf.Add(obj);

                    ITextExtractionStrategy pdfSharp = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();

                    if (sourceFile != "" && sourceFile != null)
                    {
                        ////Reading Clearence Date from PDF 
                        using (PdfReader reader = new PdfReader(destFile))
                        {
                            StringBuilder text = new StringBuilder();

                            for (int j = 1; j <= reader.NumberOfPages; j++)
                            {
                                string thePage = PdfTextExtractor.GetTextFromPage(reader, j, pdfSharp);
                                string[] theLines = thePage.Split('\n');
                                foreach (var theLine in theLines)
                                {
                                    text.AppendLine(theLine);
                                }
                            }

                            if (text.ToString().Contains("clearance status is due on"))
                            {
                                clearance_Index = text.ToString().LastIndexOf("clearance status is due on");
                                dateformat = text.ToString().Substring(clearance_Index + 27, 18).Split('.')[0];
                            }
                            else if (text.ToString().Contains("above-referenced firm to"))
                            {
                                clearance_Index = text.ToString().LastIndexOf("above-referenced firm to");
                                dateformat = text.ToString().Substring(clearance_Index + 24, 18).Split('.')[0];
                            }
                            else if (text.ToString().Contains("assessment remittance requirements to"))
                            {
                                clearance_Index = text.ToString().LastIndexOf("assessment remittance requirements to");
                                dateformat = text.ToString().Substring(clearance_Index + 37, 18).Split('.')[0];
                            }
                            if (dateformat == null && dateformat == "")
                            {
                                dateformat = DateTime.Now.ToString("MM-dd-yyyy");
                            }
                        }
                    }


                    //       Thread.Sleep(1500);

                    if (DateTime.Now > Convert.ToDateTime(dateformat))
                    {
                        //ClearanceAdded = fileName.Replace(".pdf", "  Clearance Expired " + date.ToString("MM-dd-yyyy") + ".pdf");
                        string expireddestFile = System.IO.Path.Combine(targetPath + "/Expired", obj.ClientCode_ClientName.ToUpper() + ".pdf");
                        obj.Status = "Clearance Expired";
                        obj.ClearanceDate = dateformat;

                        if (obj != null)
                        {
                            list.Remove(obj);
                            list.Add(obj);
                            listTracking.Remove(obj);
                            listTracking.Add(obj);
                        }

                        //     saveExcel(list);
                        if (!File.Exists(expireddestFile))
                        {
                            System.IO.File.Move(destFile, expireddestFile);
                        }
                    }
                    else
                    {
                        obj.Status = "Successfully Downloaded";
                        obj.ClearanceDate = dateformat;

                        if (obj != null)
                        {
                            list.Remove(obj);
                            list.Add(obj);
                            listTracking.Remove(obj);
                            listTracking.Add(obj);
                        }
                        saveMainExcel_AccountNumbers(list);
                        saveExcel(listTracking);
                    }
                    //  }
                }
            }
        }

        //private void savingPDF()
        //{
        //    ////Saving and moving file code  starts
        //    string sourcePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        //    string targetPath = DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy");
        //    //  Thread.Sleep(1500);
        //    DirectoryInfo d = new DirectoryInfo(sourcePath);//Assuming Test is your Folder

        //    lock (thisLock)
        //    {
        //        //Getting Text files
        //        string sourceFile = "";
        //        string dateformat = "";
        //        int clearance_Index = 0;
        //        string destFile = "";
        //        DateTime date = new DateTime();

        //        FileInfo[] Files = d.GetFiles("*.pdf");
        //        foreach (FileInfo file in Files)
        //        {
        //            string filename = file.Name.Replace(".pdf", "").ToLower();
        //            Document obj = list.Where(x => x.ClientCode_ClientName == filename).FirstOrDefault();

        //            if (obj != null)
        //            {
        //                // ClearanceAdded = fileName.Replace(".pdf", "  Clearance Valid Until " + date.ToString("MM-dd-yyyy") + ".pdf");
        //                sourceFile = System.IO.Path.Combine(sourcePath, file.Name);
        //                destFile = System.IO.Path.Combine(targetPath, file.Name);
        //                // To copy a file to another location and 
        //                // overwrite the destination file if it already exists.
        //                if (!File.Exists(destFile))
        //                {
        //                    System.IO.File.Move(sourceFile, destFile);
        //                }

        //                //  obj.Legalname_Tradename = filename;
        //                obj.ClientCode = file.Name.Split('_')[0];
        //                obj.Status = "Success";
        //                listSavedPdf.Add(obj);

        //                ITextExtractionStrategy pdfSharp = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();

        //                if (sourceFile != "" && sourceFile != null)
        //                {
        //                    ////Reading Clearence Date from PDF 
        //                    using (PdfReader reader = new PdfReader(destFile))
        //                    {
        //                        StringBuilder text = new StringBuilder();

        //                        for (int j = 1; j <= reader.NumberOfPages; j++)
        //                        {
        //                            string thePage = PdfTextExtractor.GetTextFromPage(reader, j, pdfSharp);
        //                            string[] theLines = thePage.Split('\n');
        //                            foreach (var theLine in theLines)
        //                            {
        //                                text.AppendLine(theLine);
        //                            }
        //                        }

        //                        if (text.ToString().Contains("clearance status is due on"))
        //                        {
        //                            clearance_Index = text.ToString().LastIndexOf("clearance status is due on");
        //                            dateformat = text.ToString().Substring(clearance_Index + 27, 18).Split('.')[0];
        //                        }
        //                        else if (text.ToString().Contains("above-referenced firm to"))
        //                        {
        //                            clearance_Index = text.ToString().LastIndexOf("above-referenced firm to");
        //                            dateformat = text.ToString().Substring(clearance_Index + 24, 18).Split('.')[0];
        //                        }
        //                        else if (text.ToString().Contains("assessment remittance requirements to"))
        //                        {
        //                            clearance_Index = text.ToString().LastIndexOf("assessment remittance requirements to");
        //                            dateformat = text.ToString().Substring(clearance_Index + 37, 18).Split('.')[0];
        //                        }
        //                        if (dateformat == null && dateformat == "")
        //                        {
        //                            dateformat = DateTime.Now.ToString("MM-dd-yyyy");
        //                        }
        //                    }
        //                }


        //                Thread.Sleep(1500);

        //                if (DateTime.Now > Convert.ToDateTime(dateformat))
        //                {
        //                    //ClearanceAdded = fileName.Replace(".pdf", "  Clearance Expired " + date.ToString("MM-dd-yyyy") + ".pdf");
        //                    string expireddestFile = System.IO.Path.Combine(targetPath + "/Expired", obj.ClientCode_ClientName.ToUpper() + ".pdf");
        //                    obj.Status = "Clearance Expired";
        //                    obj.ClearanceDate = dateformat;

        //                    if (obj != null)
        //                    {
        //                        list.Remove(obj);
        //                        list.Add(obj);
        //                        listTracking.Remove(obj);
        //                        listTracking.Add(obj);
        //                    }

        //                    //     saveExcel(list);
        //                    if (!File.Exists(expireddestFile))
        //                    {
        //                        System.IO.File.Move(destFile, expireddestFile);
        //                    }
        //                }
        //                else
        //                {
        //                    obj.Status = "Successfully Downloaded";
        //                    obj.ClearanceDate = dateformat;

        //                    if (obj != null)
        //                    {
        //                        list.Remove(obj);
        //                        list.Add(obj);
        //                        listTracking.Remove(obj);
        //                        listTracking.Add(obj);
        //                    }
        //                    saveMainExcel_AccountNumbers(list);
        //                    saveExcel(listTracking);
        //                }
        //            }
        //        }
        //    }
        //}

        private void Start_Button(object sender, EventArgs e)
        {
            try
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
                    }
                }
                else
                {
                    numberofDocumentsSelected = int.Parse(Math.Round(Convert.ToDecimal(numberOfDocuments / numberOFTurns)).ToString());

                    switch (numberOFTurns)
                    {
                        case 2:
                            for (int i = 0; i < numberOFTurns; i++)
                            {
                                if (i == 0 && i < 1)
                                {
                                    List<Document> listSelected = listDocuments.Take(numberofDocumentsSelected).ToList();
                                    Thread obj = new Thread(() => startingThread(listSelected));
                                    listDocuments.RemoveAll(x => listSelected.Contains(x));
                                    obj.Start();
                                }
                                else
                                {
                                    Thread obj = new Thread(() => startingThread3(listDocuments));
                                    obj.Start();
                                }
                            }
                            break;
                        case 4:
                            for (int i = 0; i < numberOFTurns; i++)
                            {
                                if (i >= 0 && i < 3)
                                {
                                    List<Document> listSelected = listDocuments.Take(numberofDocumentsSelected).ToList();
                                    Thread obj = new Thread(() => startingThread(listSelected));
                                    listDocuments.RemoveAll(x => listSelected.Contains(x));
                                    obj.Start();
                                }
                                else
                                {
                                    Thread obj = new Thread(() => startingThread3(listDocuments));
                                    obj.Start();
                                }
                            }
                            break;
                        default:
                            for (int i = 0; i < numberOFTurns; i++)
                            {
                                if (i >= 0 && i < 3)
                                {
                                    List<Document> listSelected = listDocuments.Take(numberofDocumentsSelected).ToList();
                                    Thread obj = new Thread(() => startingFinalThread(listSelected));
                                    listDocuments.RemoveAll(x => listSelected.Contains(x));
                                    obj.Start();
                                }
                                else if (i >= 3 && i < 5)
                                {
                                    List<Document> listSelected = listDocuments.Take(numberofDocumentsSelected).ToList();
                                    Thread obj = new Thread(() => startingFinalThread(listSelected));
                                    listDocuments.RemoveAll(x => listSelected.Contains(x));
                                    obj.Start();
                                }
                                else
                                {
                                    Thread obj = new Thread(() => startingThread3(listDocuments));
                                    obj.Start();
                                }
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Stop_Click(object sender, EventArgs e)
        {
            try
            {
                //Resetting box
                ExcelTextbox.Text = "";
                DownloadPath.Text = "";
                RecordsCount.Text = "";
                comboBox1.ResetText();


                foreach (var process in Process.GetProcessesByName("EXCEL"))
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

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void saveExcel(List<Document> listSavedPdf)
        {
            try
            {
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                //listSavedPdf.Add(obj);
                //listSavedPdf = listSavedPdf.GroupBy(x => x.Legalname).Select(y => y.First()).ToList();

                xlWorkSheet.Cells[1, 1] = "Client Account Number[Numeric]";
                xlWorkSheet.Cells[1, 2] = "Client Name[Text]";
                xlWorkSheet.Cells[1, 3] = "Client  Country[Text]";
                xlWorkSheet.Cells[1, 4] = "Client  Address[Text]";
                xlWorkSheet.Cells[1, 5] = "Client  City[Text]";
                xlWorkSheet.Cells[1, 6] = "Client  Province[Text]";
                xlWorkSheet.Cells[1, 7] = "Client  Postal code Eg: M5H 3R4";
                xlWorkSheet.Cells[1, 8] = "Client  Phone number###-###-####";
                xlWorkSheet.Cells[1, 9] = "Client Ext[Numeric]";
                xlWorkSheet.Cells[1, 10] = "Client  Email address[Mail Format]";
                xlWorkSheet.Cells[1, 11] = "Contractor Account Number[Numeric]";
                xlWorkSheet.Cells[1, 12] = "Contractor Name[Text]";
                xlWorkSheet.Cells[1, 13] = "Client Code";
                xlWorkSheet.Cells[1, 14] = "Status";
                xlWorkSheet.Cells[1, 15] = "Clearance Date";

                for (int j = 0; j < listSavedPdf.Count; j++)
                {
                    xlWorkSheet.Cells[j + 2, 1] = listSavedPdf[j].AccountNumber;
                    xlWorkSheet.Cells[j + 2, 2] = listSavedPdf[j].Legalname;
                    xlWorkSheet.Cells[j + 2, 3] = listSavedPdf[j].Country;
                    xlWorkSheet.Cells[j + 2, 4] = listSavedPdf[j].Address;
                    xlWorkSheet.Cells[j + 2, 5] = listSavedPdf[j].City;
                    xlWorkSheet.Cells[j + 2, 6] = listSavedPdf[j].Province;
                    xlWorkSheet.Cells[j + 2, 7] = listSavedPdf[j].Postalcode;
                    xlWorkSheet.Cells[j + 2, 8] = listSavedPdf[j].Phonenumber;
                    xlWorkSheet.Cells[j + 2, 9] = listSavedPdf[j].Ext;
                    xlWorkSheet.Cells[j + 2, 10] = listSavedPdf[j].Emailaddress;
                    xlWorkSheet.Cells[j + 2, 11] = listSavedPdf[j].ClientAccoutnumber;
                    xlWorkSheet.Cells[j + 2, 12] = listSavedPdf[j].Legalname_Tradename;
                    xlWorkSheet.Cells[j + 2, 13] = listSavedPdf[j].ClientCode;
                    xlWorkSheet.Cells[j + 2, 14] = listSavedPdf[j].Status;
                    xlWorkSheet.Cells[j + 2, 15] = listSavedPdf[j].ClearanceDate;
                }
                xlApp.DisplayAlerts = false;
                xlWorkBook.SaveAs(DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy") + "\\Tracking.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void saveMainExcel_AccountNumbers(List<Document> listDocumentDetails)
        {
            try
            {
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "Client Account Number[Numeric]";
                xlWorkSheet.Cells[1, 2] = "Client Name[Text]";
                xlWorkSheet.Cells[1, 3] = "Client  Country[Text]";
                xlWorkSheet.Cells[1, 4] = "Client  Address[Text]";
                xlWorkSheet.Cells[1, 5] = "Client  City[Text]";
                xlWorkSheet.Cells[1, 6] = "Client  Province[Text]";
                xlWorkSheet.Cells[1, 7] = "Client  Postal code Eg: M5H 3R4";
                xlWorkSheet.Cells[1, 8] = "Client  Phone number###-###-####";
                xlWorkSheet.Cells[1, 9] = "Client Ext[Numeric]";
                xlWorkSheet.Cells[1, 10] = "Client  Email address[Mail Format]";
                xlWorkSheet.Cells[1, 11] = "Contractor Account Number[Numeric]";
                xlWorkSheet.Cells[1, 12] = "Contractor Name[Text]";
                xlWorkSheet.Cells[1, 13] = "Client Code";

                for (int f = 1; f <= 13; f++)
                {
                    xlWorkSheet.Cells[1, f].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SandyBrown);
                }
                xlWorkSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                for (int j = 0; j < listDocumentDetails.Count; j++)
                {
                    xlWorkSheet.Cells[j + 2, 1] = listDocumentDetails[j].AccountNumber;
                    xlWorkSheet.Cells[j + 2, 2] = listDocumentDetails[j].Legalname;
                    xlWorkSheet.Cells[j + 2, 3] = listDocumentDetails[j].Country;
                    xlWorkSheet.Cells[j + 2, 4] = listDocumentDetails[j].Address;
                    xlWorkSheet.Cells[j + 2, 5] = listDocumentDetails[j].City;
                    xlWorkSheet.Cells[j + 2, 6] = listDocumentDetails[j].Province;
                    xlWorkSheet.Cells[j + 2, 7] = listDocumentDetails[j].Postalcode;
                    xlWorkSheet.Cells[j + 2, 8] = listDocumentDetails[j].Phonenumber;
                    xlWorkSheet.Cells[j + 2, 9] = listDocumentDetails[j].Ext;
                    xlWorkSheet.Cells[j + 2, 10] = listDocumentDetails[j].Emailaddress;
                    xlWorkSheet.Cells[j + 2, 11] = listDocumentDetails[j].ClientAccoutnumber;
                    xlWorkSheet.Cells[j + 2, 12] = listDocumentDetails[j].Legalname_Tradename;
                    xlWorkSheet.Cells[j + 2, 13] = listDocumentDetails[j].ClientCode;
                }
                xlApp.DisplayAlerts = false;
                //ofd.FileName
                var OldPath = ofd.FileName;
                var newPath = ofd.FileName.Replace(".xlsx", ".xls");
                System.IO.File.Delete(OldPath);
                xlWorkBook.SaveAs(newPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void KillSpecificExcelFileProcess(string excelFileName)
        {
            try
            {
                foreach (var process in Process.GetProcessesByName("EXCEL"))
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void savingPDFAllOnce()
        {
            ////Saving and moving file code  starts
            string sourcePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            string targetPath = DownloadPath.Text + "/" + DateTime.Now.ToString("MM-dd-yyyy");
            //  Thread.Sleep(1500);
            DirectoryInfo d = new DirectoryInfo(sourcePath);//Assuming Test is your Folder

            //lock (thisLock)
            //{
            //Getting Text files
            string sourceFile = "";
            string dateformat = "";
            int clearance_Index = 0;
            string destFile = "";

            FileInfo[] Files = d.GetFiles("*.pdf"); //Getting Text files
            int filesCount = Files.Count();
            int currentFileCount = 0;
            foreach (FileInfo file in Files)
            {
                currentFileCount = currentFileCount + 1;
                string code_clientName = file.Name.Replace(".pdf", "").ToLower();
                Document objDocument = list.Where(x => x.ClientCode_ClientName == code_clientName).FirstOrDefault();

                if (objDocument != null)
                {
                    sourceFile = System.IO.Path.Combine(sourcePath, code_clientName.ToUpper()).Replace("\\", "/") + ".pdf";
                    destFile = System.IO.Path.Combine(targetPath, code_clientName.ToUpper()).Replace("\\", "/") + ".pdf";
                    // To copy a file to another location and 
                    // overwrite the destination file if it already exists.

                    if (File.Exists(sourceFile))
                    {
                        if (!File.Exists(destFile))
                        {
                            System.IO.File.Move(sourceFile, destFile);
                        }

                        ITextExtractionStrategy pdfSharp = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();

                        if (sourceFile != "" && sourceFile != null)
                        {
                            ////Reading Clearence Date from PDF 
                            using (PdfReader reader = new PdfReader(destFile))
                            {
                                StringBuilder text = new StringBuilder();

                                for (int j = 1; j <= reader.NumberOfPages; j++)
                                {
                                    string thePage = PdfTextExtractor.GetTextFromPage(reader, j, pdfSharp);
                                    string[] theLines = thePage.Split('\n');
                                    foreach (var theLine in theLines)
                                    {
                                        text.AppendLine(theLine);
                                    }
                                }

                                if (text.ToString().Contains("clearance status is due on"))
                                {
                                    clearance_Index = text.ToString().LastIndexOf("clearance status is due on");
                                    dateformat = text.ToString().Substring(clearance_Index + 27, 18).Split('.')[0];
                                }
                                else if (text.ToString().Contains("above-referenced firm to"))
                                {
                                    clearance_Index = text.ToString().LastIndexOf("above-referenced firm to");
                                    dateformat = text.ToString().Substring(clearance_Index + 24, 18).Split('.')[0];
                                }
                                else if (text.ToString().Contains("assessment remittance requirements to"))
                                {
                                    clearance_Index = text.ToString().LastIndexOf("assessment remittance requirements to");
                                    dateformat = text.ToString().Substring(clearance_Index + 37, 18).Split('.')[0];
                                }
                                if (dateformat == null && dateformat == "")
                                {
                                    dateformat = DateTime.Now.ToString("MM-dd-yyyy");
                                }
                            }
                        }


                        //       Thread.Sleep(1500);

                        if (DateTime.Now > Convert.ToDateTime(dateformat))
                        {
                            //ClearanceAdded = fileName.Replace(".pdf", "  Clearance Expired " + date.ToString("MM-dd-yyyy") + ".pdf");
                            string expireddestFile = System.IO.Path.Combine(targetPath + "/Expired", objDocument.ClientCode_ClientName.ToUpper() + ".pdf");
                            objDocument.Status = "Clearance Expired";
                            objDocument.ClearanceDate = dateformat;

                            if (objDocument != null)
                            {
                                list.Remove(objDocument);
                                list.Add(objDocument);
                                listTracking.Remove(objDocument);
                                listTracking.Add(objDocument);
                            }

                            //     saveExcel(list);
                            if (!File.Exists(expireddestFile))
                            {
                                System.IO.File.Move(destFile, expireddestFile);
                            }
                        }
                        else
                        {
                            objDocument.Status = "Successfully Downloaded";
                            objDocument.ClearanceDate = dateformat;

                            if (objDocument != null)
                            {
                                list.Remove(objDocument);
                                list.Add(objDocument);
                                listTracking.Remove(objDocument);
                                listTracking.Add(objDocument);
                            }
                        }
                    }


                }
            }
            if (filesCount == currentFileCount)
            {
                saveMainExcel_AccountNumbers(list);
                saveExcel(listTracking);
            }
        }
    }
}
