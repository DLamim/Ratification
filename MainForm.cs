using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Avoidable_Cost_Tool;
using System.Configuration;
using System.IO;
using Avoidable_Cost_Tool.Class;
using System.Data.SqlClient;
using System.Globalization;
using ExportToExcel;
using OfficeOpenXml;

namespace Avoidable_Cost_Tool
{
    public partial class MainForm : Form
    {
        string userEmail;
        int userLevel;
        ProcessResult processResult = new ProcessResult();


        //Default
        public MainForm()
        {
            InitializeComponent();
        }
        //Overloaded
        public MainForm(string email, int lvl)
        {
            InitializeComponent();
            userEmail = email;
            userLevel = lvl;
        }
        //Form load
        private void MainForm_Load(object sender, EventArgs e)
        {
            //Populate customer combobox
            LoadClientCB();

            //Check level
            if (userLevel > 1)
            {
                editToolStripMenuItem.Enabled = true;
                if (userLevel > 2)
                {
                    //processToolStripMenuItem.Enabled = true;
                }
            }

        }
        //Load clients CB
        private void LoadClientCB()
        {
            //Load list of clients
            Database1DataSetTableAdapters.clientsTableAdapter clientTA = new Database1DataSetTableAdapters.clientsTableAdapter();
            Database1DataSet.clientsDataTable clientDT = new Database1DataSet.clientsDataTable();
            clientTA.FillClientsCb(clientDT);

            clientsCb.DataSource = clientDT;
            clientsCb.DisplayMember = "clientName";
            clientsCb.ValueMember = "ID";

            clientsCb.SelectedIndex = -1;
        }
        //function read DB to load process control DGV
        private void UserInfoDB(string mode)
        {
            string gap = "";
            DateTime date = System.DateTime.Now;
            DateTime dateFirst = System.DateTime.Now;
            double daysGap = 0;
            Int32 rowNumber = 0;

            try
            {
                Database1DataSetTableAdapters.clientsTableAdapter clientsTA = new Database1DataSetTableAdapters.clientsTableAdapter();
                Database1DataSet.clientsDataTable clientsDT = new Database1DataSet.clientsDataTable();

                if (mode == "myClients")
                {
                    clientsTA.FillByMyClients(clientsDT, userEmail);
                }
                else
                {
                    clientsTA.FillByAllClientsData(clientsDT);
                }

                clientsDT.Columns.Add(new DataColumn("Gap", typeof(double)));
                clientsDT.Columns.Add(new DataColumn("Ready", typeof(string)));

                customerProcessDGV.DataSource = clientsDT;

                do
                {
                    if (!(customerProcessDGV.Rows[rowNumber].Cells[5].Value == null) && !string.IsNullOrWhiteSpace(customerProcessDGV.Rows[rowNumber].Cells[5].Value.ToString()))
                    {
                        gap = customerProcessDGV.Rows[rowNumber].Cells[5].Value.ToString();
                        gap = gap.Substring(13);

                        date = DateTime.Parse(gap);
                        TimeSpan t = dateFirst - date;
                        daysGap = Math.Round(t.TotalDays) - 1;

                        customerProcessDGV.Rows[rowNumber].Cells["Gap"].Value = daysGap;

                        if (daysGap > Convert.ToDouble(customerProcessDGV.Rows[rowNumber].Cells[2].Value))
                        {
                            customerProcessDGV.Rows[rowNumber].Cells[10].Value = "Yes";
                        }
                        else
                        {
                            customerProcessDGV.Rows[rowNumber].Cells[10].Value = "No";
                        }
                    }
                    rowNumber++;
                }
                while (rowNumber < customerProcessDGV.RowCount);

                customerProcessDGV.Columns[0].Visible = false;
                customerProcessDGV.Columns[3].Visible = false;
                customerProcessDGV.Columns[4].Visible = false;
                customerProcessDGV.Columns[6].Visible = false;
                customerProcessDGV.Columns[7].Visible = false;
                customerProcessDGV.Columns[8].Visible = false;

                customerProcessDGV.Columns[1].HeaderText = "Client name";
                customerProcessDGV.Columns[2].HeaderText = "Interval";
                customerProcessDGV.Columns[4].HeaderText = "Last process date";
                customerProcessDGV.Columns[5].HeaderText = "Period used in last process";
                customerProcessDGV.Columns[9].HeaderText = "Number of days not logged";
                customerProcessDGV.Columns[10].HeaderText = "Is ready to process?";

                customerProcessDGV.Columns[1].Width = 240;
                customerProcessDGV.Columns[2].Width = 70;
                customerProcessDGV.Columns[4].Width = 80;
                customerProcessDGV.Columns[5].Width = 140;

            }
            catch (Exception exx)
            {
                MessageBox.Show("Could not access database. Error: " + exx, "Alert");
            }

        }
        //Button check avoidable cost exception click
        private void checkAvoidExpBT_Click(object sender, EventArgs e)
        {
            AvoidableCostException excptForm = new AvoidableCostException();
            excptForm.ShowDialog();
        }
        //Button customer configuration click
        private void configCustomerBt_Click(object sender, EventArgs e)
        {
            //TBD
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
        //Button search customer history click
        private void searchCustBt_Click(object sender, EventArgs e)
        {

        }
        //Button log avoidable cost click
        private void logAvoidBt_Click(object sender, EventArgs e)
        {

        }
        //Button extract avoidable cost click
        private async void extractAvoidBt_Click(object sender, EventArgs e)
        {
            if (clientsCb.SelectedIndex >= 0)
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Title = "Please select spreadsheet with list of avoidable costs already logged";
                DialogResult result = openFile.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(openFile.FileName))
                {
                    string avoidableLogged = openFile.FileName;
                    openFile.Title = "Please select spreadsheet of invoices to seach for avoidable costs. ";
                    result = openFile.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(openFile.FileName))
                    {
                        string avoidableToSearch = openFile.FileName;

                        mainStatusLb.Text = "Please wait, application is extracting information from infoLCM...";
                        processResult.Started = DateTime.Now;

                        //Progress report
                        var progress = new Progress<string>(str =>
                            {
                                mainStatusLb.Text = str;
                                toolStrip1.Refresh();
                            });
                        var alreadyLogged = await ReturnLoggedFromInfo(avoidableLogged, progress);
                        var idsToSearch = await ReturnIDsToSearchTest(avoidableToSearch, progress);
                        var acFound = await FindAcInFiles(idsToSearch, progress);
                        //var ordanizedLines = await OrganizeRawLines(FindAcInFiles(avoidableToSearch));
                        //CompareList(ReturnLoggedFromInfo(avoidableLogged), OrganizeRawLines(FindAcInFiles(avoidableToSearch)));
                        CompareList(alreadyLogged, OrganizeRawLines(acFound));
                        processResult.Ended = DateTime.Now;
                        mainStatusLb.Text = "Ready";
                        toolStrip1.Refresh();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a client first.", "Alert");
            }
        }
        //button load my clients info
        private void loadCustInfoBT_Click(object sender, EventArgs e)
        {
            UserInfoDB("myClients");
        }
        //User control menu 
        private void userControlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Users users = new Users(userEmail, userLevel);
            users.ShowDialog();

            this.Show();
        }
        //Clients control menu
        private void customerControlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Clients clients = new Clients(userEmail);
            clients.ShowDialog();

            this.Show();
        }
        //Flag control menu
        private void avoidableCostToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Flags flags = new Flags(userEmail);
            flags.ShowDialog();

            this.Show();
        }
        //Adm control menu
        private void processToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            Process pro = new Process();
            pro.ShowDialog();

            this.Show();
        }
        //button load all clients info
        private void allCliLoadInfoBt_Click(object sender, EventArgs e)
        {
            UserInfoDB("allClients");
        }
        //Function to search line items file (.LIA .LIB...)
        private async Task<List<InvoiceInfo>> FindAcInFiles(List<InvoiceToSearch> listIdsFromInfo, IProgress<string> progress)
        {
            //Call function to bring line codes from DB
            //lineCodes = GetLineCode();
            var lineCodes = GetLineCodeTest();
            //Call function to get client's folder path
            //string path = GetClientPath().FolderPath;
            string path = "D:\\Dionisio\\Visual Studio\\Avoidable cost files\\ACH";
            //Avoidable costs will be store in this array
            List<InvoiceInfo> mainACList = new List<InvoiceInfo>();

            //If path is not null and list has items
            if (path != "" && listIdsFromInfo.Count > 0)
            {
                string idPath = path;
                string targetDir = System.IO.Path.GetTempPath();
                string targetDirFull = string.Format("{0}\\", targetDir);

                //Loop througth folders
                for (char a = 'a'; a <= 'z'; a++)
                {
                    string entry = a.ToString().ToUpper();
                    string folderDir = string.Format("{0}\\{1}\\", idPath, entry);

                    //Loop throught list of IDs
                    for (int i = 0; i < listIdsFromInfo.Count; i++)
                    {
                        var oneTask = Task.Factory.StartNew(() =>
                        {
                            //Add file extension according to the folder
                            string invoice = listIdsFromInfo[i].InvoiceID;
                            string invoiceLi = invoice.Substring(0, invoice.Length - 1) + ".LI" + invoice.Substring(invoice.Length - 1).ToUpper();
                            string invoicePath = folderDir + invoiceLi;

                            //Show status to user
                            progress.Report(string.Format("Remaining invoices to read: {0} Reading lines from invoice {1}", listIdsFromInfo.Count, invoice));

                            if (File.Exists(invoicePath))
                            {
                                string line = "";
                                //Copy file to a temp folder
                                File.Copy(folderDir + invoiceLi, Path.Combine(targetDir, Path.GetFileName(folderDir + invoiceLi)), true);

                                StreamReader reader = new StreamReader(targetDirFull + invoiceLi);
                                while ((line = reader.ReadLine()) != null)
                                {
                                    double lineCost = 0;

                                    if (double.TryParse(line.Substring(98, 11), out lineCost))
                                    {
                                        string lineCode = line.Substring(17, 4);
                                        //If line cost not zero//loop to check if lineCode matches element from codesAc
                                        if (lineCost > 0)
                                        {
                                            for (int c = 0; c < lineCodes.Count; c++)
                                            {
                                                //if code matches//get info in array and add to list
                                                string compareCode = lineCodes[c].Code;

                                                if (lineCode == compareCode)
                                                {
                                                    //Execute function to check avoidable costs in this invoice
                                                    List<InvoiceInfo> newAvoidables = new List<InvoiceInfo>();
                                                    newAvoidables = CheckLines(targetDirFull + invoiceLi, lineCodes, invoice, listIdsFromInfo[i].ProjectCode);

                                                    for (int k = 0; k < newAvoidables.Count; k++)
                                                    {
                                                        mainACList.Add(newAvoidables[k]);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                reader.Close();
                                listIdsFromInfo.RemoveAt(i);
                                i--;
                            }
                        });
                        Task.WaitAll(oneTask);
                    }

                    

                }
            }
            return mainACList;
        }
        //Function to get project code list from client selected
        private List<ClientInfo> GetClientInfo()
        {
            List<ClientInfo> listProjCode = new List<ClientInfo>();

            try
            {
                Int32 clientID = Convert.ToInt32(clientsCb.SelectedValue);

                Database1DataSetTableAdapters.projectCodesTableAdapter getProjTA = new Database1DataSetTableAdapters.projectCodesTableAdapter();
                Database1DataSet.projectCodesDataTable getProjDT = new Database1DataSet.projectCodesDataTable();
                getProjTA.FillByClient(getProjDT, clientID.ToString());

                for (int i = 0; i < getProjDT.Rows.Count; i++)
                {
                    ClientInfo clientInfo = new ClientInfo();

                    clientInfo.ProjectCodes = getProjDT[i].projectCode.ToString();
                    clientInfo.Country = getProjDT[i].country.ToString();

                    listProjCode.Add(clientInfo);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not access database. Error: " + ex, "Alert");
            }

            return listProjCode;
        }
        //Function to get folder path from client selected
        private ClientInformation GetClientPath()
        {
            ClientInformation client = new ClientInformation();

            try
            {
                Int32 clientID = Convert.ToInt32(clientsCb.SelectedValue);

                Database1DataSetTableAdapters.clientsTableAdapter getClientTA = new Database1DataSetTableAdapters.clientsTableAdapter();
                Database1DataSet.clientsDataTable getClientDT = new Database1DataSet.clientsDataTable();
                getClientTA.FillByClientInfo(getClientDT, clientID);

                client.Interval = getClientDT[0].period;
                client.FolderPath = getClientDT[0].folderPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not access database. Error: " + ex, "Alert");
            }

            return client;
        }
        //Function to get line codes related to avoidable cost
        private List<LineCode> GetLineCode()
        {
            List<LineCode> lineCodes = new List<LineCode>();

            try
            {
                Database1DataSetTableAdapters.flagsTableAdapter lineCodeTA = new Database1DataSetTableAdapters.flagsTableAdapter();
                Database1DataSet.flagsDataTable lineCodeDT = new Database1DataSet.flagsDataTable();

                Database1DataSetTableAdapters.avoidableCostCategoryTableAdapter categoriesTA = new Database1DataSetTableAdapters.avoidableCostCategoryTableAdapter();
                Database1DataSet.avoidableCostCategoryDataTable categoriesDT = new Database1DataSet.avoidableCostCategoryDataTable();


                lineCodeTA.FillByFlags(lineCodeDT);

                for (int i = 0; i < lineCodeDT.Rows.Count; i++)
                {
                    categoriesTA.FillByCategorySub(categoriesDT, Convert.ToInt32(lineCodeDT[i].subCategory));

                    LineCode lineCode = new LineCode();

                    lineCode.Code = lineCodeDT[i].flag.ToString();
                    lineCode.Description = lineCodeDT[i].description.ToString();
                    lineCode.Subcategory = categoriesDT[0].subcategory;
                    lineCode.Category = categoriesDT[0].category;

                    lineCodes.Add(lineCode);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not access database. Error: " + ex, "Alert");
            }
            return lineCodes;
        }
        //Funtion to export avoidable cost already logged from infoLCM
        private async Task<List<InvoiceToSearch>> ReturnLoggedFromInfo(string fileName, IProgress<string> progress)
        {
            //List of project codes
            //listProjCode = GetClientInfo();
            var listProjCode = GetPJCode();
            //List of avoidable cost logged from infoLCM//Class to store invoice data
            List<InvoiceToSearch> avoidableCostLogged = new List<InvoiceToSearch>();
            //Connect to info
            SqlConnection con;
            string sheet1 = "";

            try
            {
                //Code to read excell spreadsheet with all logged ACs from InfoLCM
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended properties=""Excel 8.0;HDR=No;IMEX=1;MAXSCANROWS=0;ImportMixedTypes=Text;TypeGuessRows=0""", fileName);

                using (OleDbConnection connTwo = new OleDbConnection(conn.ConnectionString))
                {
                    connTwo.Open();
                    DataTable dtSchema = connTwo.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                }
                try
                {
                    string connect = string.Format("SELECT F1, F2, F12, F13 FROM [{0}]", sheet1);

                    OleDbCommand command = new OleDbCommand();
                    command.CommandText = connect;
                    command.Connection = conn;

                    System.Data.DataTable dtCustomers = new System.Data.DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                    adapter.Fill(dtCustomers);

                    var tasks = Task.Factory.StartNew(() =>
                    {
                        for (int i = 2; i < dtCustomers.Rows.Count; i++)
                        {
                            InvoiceToSearch invoiceInfoLCM = new InvoiceToSearch();

                            invoiceInfoLCM.InvoiceID = dtCustomers.Rows[i][0].ToString();
                            invoiceInfoLCM.ProjectCode = dtCustomers.Rows[i][1].ToString();
                            invoiceInfoLCM.Subcategory = dtCustomers.Rows[i][3].ToString();

                            avoidableCostLogged.Add(invoiceInfoLCM);

                            //Show status to user
                            progress.Report(string.Format("Reading data from spreadsheet. Current row {0} ", i));
                        }
                    });
                    await Task.WhenAll(tasks);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not read the spreadsheet. Error: " + ex);
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show("Could not open connection with database! " + exp);
            }
            return avoidableCostLogged;
        }
        //Function to process and compare lists
        private bool CompareList(List<InvoiceToSearch> LoggedFromInfo, List<InvoiceInfo> avoidableToLog)
        {
            bool ok = true;

            List<InvoiceToSearch> CopyLoggedInfo = new List<InvoiceToSearch>();
            CopyLoggedInfo = LoggedFromInfo;
            //Lists to organize avoidable costs 
            List<AvoidableCost> ListACMain = new List<AvoidableCost>();
            List<InvoiceInfo> ListACOther = new List<InvoiceInfo>();
            List<AvoidableException> ListACExceptions = new List<AvoidableException>();
            //Show status to user
            mainStatusLb.Text = "Comparing avoidable costs found...";
            toolStrip1.Refresh();
            //Check if IDs from avoidableToLog is already logged
            for (int i = 0; i < avoidableToLog.Count; i++)
            {
                int index = LoggedFromInfo.FindIndex(f => f.InvoiceID == avoidableToLog[i].InvoiceID);
                //If ID was found
                if (index >= 0)
                {
                    //If exists but not for the same subcategory
                    if (!(avoidableToLog[i].Subcategory == LoggedFromInfo[index].Subcategory))
                    {
                        ListACOther.Add(avoidableToLog[i]);
                    }
                    avoidableToLog.RemoveAt(i);
                    i--;
                }
                else
                {
                    int indexTwo = avoidableToLog.FindIndex(a => a.InvoiceID == avoidableToLog[i].InvoiceID && a.Subcategory != avoidableToLog[i].Subcategory);
                    //If not logged on InfoLCM but there is more than one AC for the same invoice to be logged
                    if (indexTwo >= 0)
                    {
                        //Check if avoidableToLog[i] is for demand related issues//And if it is confirmed
                        if (avoidableToLog[i].Confirmed && avoidableToLog[i].Category == "Demand related issues")
                        {
                            //Transfer second AC from the same ID to the exception list
                            ListACOther.Add(avoidableToLog[indexTwo]);
                            avoidableToLog.RemoveAt(indexTwo);
                        }
                        if (avoidableToLog[i].Category == "Demand related issues")
                        {
                            //Transfer this AC to exception list
                            ListACOther.Add(avoidableToLog[i]);
                            avoidableToLog.RemoveAt(i);
                            i--;
                        }
                        else
                        {
                            //Transfer second AC from the same ID to the exception list
                            ListACOther.Add(avoidableToLog[indexTwo]);
                            avoidableToLog.RemoveAt(indexTwo);
                        }
                    }
                    else if (!avoidableToLog[i].Confirmed)
                    {
                        //Transfer this AC to exception list
                        ListACOther.Add(avoidableToLog[i]);
                        avoidableToLog.RemoveAt(i);
                        i--;
                    }
                }
            }
            //If avoidable costs were found
            if (avoidableToLog.Count > 0 || ListACOther.Count > 0)
            {
                //Show status to user
                mainStatusLb.Text = "Incerting information to each avoidable cost...";
                toolStrip1.Refresh();
                //Organize ACs from ListACMain
                if (avoidableToLog.Count > 0)
                {
                    for (int t = 0; t < avoidableToLog.Count; t++)
                    {
                        AvoidableCost avoidableCost = new AvoidableCost();

                        avoidableCost.InvoiceNo = avoidableToLog[t].InvoiceID;
                        avoidableCost.ProjectCode = avoidableToLog[t].ProjectCode;
                        avoidableCost.Utility = "Electricity";
                        avoidableCost.Cost = Math.Round(avoidableToLog[t].Cost, 2);
                        avoidableCost.Category = avoidableToLog[t].Category;
                        avoidableCost.Subcategory = avoidableToLog[t].Subcategory;
                        avoidableCost.DuplicateOf = "";
                        avoidableCost.Months = "1";
                        avoidableCost.Status = "Confirmed - Avoidable Cost";
                        avoidableCost.DefaultComment = string.Format("{0} - Avoidable cost of {1}{2} related to {3} (excluding taxes).", avoidableToLog[t].InvoiceID, avoidableToLog[t].CurrencySymbol, avoidableToLog[t].Cost, avoidableToLog[t].Description);

                        ListACMain.Add(avoidableCost);
                    }
                    //Add number of avoidable costs found to result class
                    processResult.AvoidablesFound = ListACMain.Count;
                }
                //Organize ACs from exception list
                if (ListACOther.Count > 0)
                {
                    for (int t = 0; t < ListACOther.Count; t++)
                    {
                        AvoidableException exceptions = new AvoidableException();

                        exceptions.InvoiceNo = ListACOther[t].InvoiceID;
                        exceptions.ProjectCode = ListACOther[t].ProjectCode;
                        exceptions.Utility = "Electricity";
                        exceptions.Cost = Math.Round(ListACOther[t].Cost, 2);
                        exceptions.Category = ListACOther[t].Category;
                        exceptions.Subcategory = ListACOther[t].Subcategory;
                        exceptions.DefaultComment = string.Format("{0} - Avoidable cost of {1}{2} related to {3} (excluding taxes).", ListACOther[t].InvoiceID, ListACOther[t].CurrencySymbol, ListACOther[t].Cost, ListACOther[t].Description);
                        exceptions.Confirmed = ListACOther[t].Confirmed;

                        ListACExceptions.Add(exceptions);
                    }
                    //Add number of avoidable costs found to result class
                    processResult.AvoidablesFound += ListACExceptions.Count;
                }
                //Log process to DB
                LogProcessIDs(ListACMain, ListACExceptions);
                //Show listMain in DGV
                ShowMainAC(ListACMain);
                //Convert main list to dataset
                DataSet dsACMain = new DataSet();
                dsACMain = ToDataSetMain<AvoidableCost>(ListACMain);
                //Convert exception list to dataset
                DataSet dsExcep = new DataSet();
                dsExcep = ToDataSetException<AvoidableException>(ListACExceptions);
                //Set process end time
                processResult.Ended = DateTime.Now;
                //Show result of process
                MessageBox.Show("Please click Ok and select a folder to save the avoidable cost spreadsheets. " + System.Environment.NewLine + System.Environment.NewLine + "Number of invoices searched: " + processResult.InvoicesSearched + System.Environment.NewLine + "Total number of avoidable cost found: " + processResult.AvoidablesFound + System.Environment.NewLine + "Process started at: " + processResult.Started + "    Process ended at: " + processResult.Ended, "Process result");
                //Select saving location for avoidable cost spreadsheet
                SaveFileDialog folder = new SaveFileDialog();
                folder.Title = "Please select a folder to save the spreadsheet. ";
                folder.FileName = "Avoidable Cost";
                DialogResult result = new DialogResult();

                do
                {
                    result = folder.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        string filePath = folder.FileName;

                        if (!string.IsNullOrWhiteSpace(filePath))
                        {
                            WriteExcel(dsACMain, filePath);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select a location to save the spreadsheet. ", "Alert");
                    }
                }
                while (result != DialogResult.OK);

                //Select saving location for exception avoidable cost spreadsheet
                folder.Title = "Please select a folder to save the exception spreadsheet. ";
                folder.FileName = "Avoidable Cost - Exception";

                do
                {
                    result = folder.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        string filePath = folder.FileName;

                        if (!string.IsNullOrWhiteSpace(filePath))
                        {
                            CreateExcelFile.CreateExcelDocument(dsExcep, filePath + ".xlsx");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select a location to save the spreadsheet. ", "Alert");
                    }
                }
                while (result != DialogResult.OK);
            }
            else
            {
                MessageBox.Show(string.Format("The application searched for avoidable costs in {0} invoices, but no avoidable costs were found. ", processResult.InvoicesSearched), "Process result");
            }
            //Log process date and interval used to clients table
            LogProcessClient();

            return ok;
        }
        //Function test to organize list before comparing
        private List<InvoiceInfo> OrganizeRawLines(List<InvoiceInfo> linesWithAc)
        {
            List<InvoiceInfo> listOfAvoidables = new List<InvoiceInfo>(linesWithAc);
            listOfAvoidables.Sort(new InvoiceComparer(true, "ID"));
            listOfAvoidables.Sort(new InvoiceComparer(true, "Code"));
            //Loop through ids to check for more than one avoidable cost in the same id            
            for (int i = 0; i < listOfAvoidables.Count; i++)
            {
                int indexTwo = listOfAvoidables.FindLastIndex(a => a.InvoiceID == listOfAvoidables[i].InvoiceID && a.Subcategory == listOfAvoidables[i].Subcategory);
                if (indexTwo >= 0)
                {
                    if (indexTwo != listOfAvoidables.IndexOf(listOfAvoidables[i]))
                    {
                        double thisIdCost = listOfAvoidables[i].Cost;
                        double nextIdCost = listOfAvoidables[indexTwo].Cost;
                        listOfAvoidables[i].Cost = thisIdCost + nextIdCost;

                        listOfAvoidables.RemoveAt(indexTwo);
                        i = i - 1;
                    }
                }
            }
            //Loop to add country and currency symbol to each AC
            Database1DataSetTableAdapters.projectCodesTableAdapter projInfoTA = new Database1DataSetTableAdapters.projectCodesTableAdapter();
            Database1DataSet.projectCodesDataTable projInfoDT = new Database1DataSet.projectCodesDataTable();

            for (int i = 0; i < listOfAvoidables.Count; i++)
            {
                try
                {
                    //projInfoTA.FillInfoByProjectCode(projInfoDT, listOfAvoidables[i].ProjectCode);

                    //var regions = CultureInfo.GetCultures(CultureTypes.SpecificCultures).Select(x => new RegionInfo(x.LCID));
                    //var englishRegion = regions.FirstOrDefault(region => region.EnglishName.Contains(projInfoDT[0].country));

                    //string currencyIsoSymbol = englishRegion.CurrencySymbol;

                    listOfAvoidables[i].Country = "United States";
                    listOfAvoidables[i].CurrencySymbol = "$";
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Could not access datase. Error: " + exx);
                }
            }
            return listOfAvoidables;
        }
        //Function to convert list avoidable cost main to dataset
        public static DataSet ToDataSetMain<T>(IList<T> list)
        {
            Type elementType = typeof(AvoidableCost);
            DataSet ds = new DataSet();
            DataTable t = new DataTable();
            ds.Tables.Add(t);

            if (list.Count >= 0)
            {
                //add a column to table for each public property on T
                foreach (var propInfo in elementType.GetProperties())
                {
                    Type ColType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

                    t.Columns.Add(propInfo.Name, ColType);
                }

                //go through each property on T and add each value to the table
                foreach (T item in list)
                {
                    DataRow row = t.NewRow();

                    foreach (var propInfo in elementType.GetProperties())
                    {
                        row[propInfo.Name] = propInfo.GetValue(item, null) ?? DBNull.Value;
                    }

                    t.Rows.Add(row);
                }
                t.Columns[0].ColumnName = "Invoice No";
                t.Columns[1].ColumnName = "Customer Code";
                t.Columns[2].ColumnName = "Utility Type";
                t.Columns[3].ColumnName = "Savings Value";
                t.Columns[4].ColumnName = "Saving Category";
                t.Columns[5].ColumnName = "Saving Subcategory";
                t.Columns[6].ColumnName = "Duplicate Invoice Of";
                t.Columns[7].ColumnName = "# Months Saving Multiplier";
                t.Columns[8].ColumnName = "Status";
                t.Columns[9].ColumnName = "Default Customer Display Comments1";
                t.Columns.Add(new DataColumn("Comment Date 1", typeof(string)));
                t.Columns.Add(new DataColumn("Default Customer Display Comments2", typeof(string)));
                t.Columns.Add(new DataColumn("Comment Date 2", typeof(string)));
                t.Columns.Add(new DataColumn("Default Customer Display Comments3", typeof(string)));
                t.Columns.Add(new DataColumn("Comment Date 3", typeof(string)));
                t.Columns.Add(new DataColumn("Comments Added to (Account/Bill)", typeof(string)));
                t.Columns.Add(new DataColumn("Reminder Date", typeof(string)));
                t.Columns.Add(new DataColumn("Reminder Email", typeof(string)));

                t.TableName = "Plan1";
                t.AcceptChanges();
            }
            return ds;
        }
        //Function to convert list avoidable cost exception to dataset
        public static DataSet ToDataSetException<T>(IList<T> list)
        {
            Type elementType = typeof(AvoidableException);
            DataSet ds = new DataSet();
            DataTable t = new DataTable();
            ds.Tables.Add(t);

            if (list.Count >= 0)
            {
                //add a column to table for each public property on T
                foreach (var propInfo in elementType.GetProperties())
                {
                    Type ColType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

                    t.Columns.Add(propInfo.Name, ColType);
                }

                //go through each property on T and add each value to the table
                foreach (T item in list)
                {
                    DataRow row = t.NewRow();

                    foreach (var propInfo in elementType.GetProperties())
                    {
                        row[propInfo.Name] = propInfo.GetValue(item, null) ?? DBNull.Value;
                    }

                    t.Rows.Add(row);
                }
                t.Columns[0].ColumnName = "Invoice No";
                t.Columns[1].ColumnName = "Customer Code";
                t.Columns[2].ColumnName = "Utility Type";
                t.Columns[3].ColumnName = "Savings Value";
                t.Columns[4].ColumnName = "Saving Category";
                t.Columns[5].ColumnName = "Saving Subcategory";
                t.Columns[6].ColumnName = "Customer Display Comments";
                t.Columns[7].ColumnName = "Confirmed";

                t.AcceptChanges();
            }
            return ds;
        }
        //Function to show main list of ACs to log in DGV
        private void ShowMainAC(List<AvoidableCost> ac)
        {
            //Show listmain in DGV
            var bindingList = new BindingList<AvoidableCost>(ac);
            var source = new BindingSource(bindingList, null);
            avoidablesDGV.DataSource = source;
        }
        //Test to show invoices not found
        private void ShowNotFound(List<InvoiceToSearch> ac)
        {
            //Show listmain in DGV
            var bindingList = new BindingList<InvoiceToSearch>(ac);
            var source = new BindingSource(bindingList, null);
            avoidablesDGV.DataSource = source;
        }
        //Function to log process//will log all avoidable costs identified
        private void LogProcessIDs(List<AvoidableCost> listMain, List<AvoidableException> listExcep)
        {
            //try
            //{
            //    Database1DataSetTableAdapters.invoicesTableAdapter logIdsTA = new Database1DataSetTableAdapters.invoicesTableAdapter();
            //    Database1DataSet.invoicesDataTable logIdsDT = new Database1DataSet.invoicesDataTable();

            //    Database1DataSetTableAdapters.SpreadsheetTableAdapter logSpreadsheetTA = new Database1DataSetTableAdapters.SpreadsheetTableAdapter();
            //    Database1DataSet.SpreadsheetDataTable logSpreadsheetDT = new Database1DataSet.SpreadsheetDataTable();
            //    //If main has invoices//Log process from list main//Log info in spreadsheet table in DB
            //    if (listMain.Count >= 0)
            //    {
            //        string sheetName = string.Format("AvoidableCost_{0}_{1}", clientsCb.Text.Replace(" ", string.Empty), DateTime.Now.ToLongDateString());

            //        logSpreadsheetTA.InsertLog(sheetName, "normal", userEmail, DateTime.Now);
            //        logSpreadsheetTA.FillByName(logSpreadsheetDT, sheetName);

            //        for (int i = 0; i < listMain.Count; i++)
            //        {
            //            logIdsTA.InsertQuery(listMain[i].InvoiceNo, listMain[i].Subcategory, clientsCb.SelectedValue.ToString(), logSpreadsheetDT[0].ID, DateTime.Now);
            //        }
            //    }
            //    //If exception has invoices//Log process from list exception//Log info in spreadsheet table in DB
            //    if (listExcep.Count >= 0)
            //    {
            //        string sheetName = string.Format("AvoidableCostException_{0}_{1}", clientsCb.Text.Replace(" ", string.Empty), DateTime.Now.ToLongDateString());

            //        logSpreadsheetTA.InsertLog(sheetName, "exception", userEmail, DateTime.Now);
            //        logSpreadsheetTA.FillByName(logSpreadsheetDT, sheetName);

            //        for (int i = 0; i <= listExcep.Count - 1; i++)
            //        {
            //            logIdsTA.InsertQuery(listExcep[i].InvoiceNo, listExcep[i].Subcategory, clientsCb.SelectedValue.ToString(), logSpreadsheetDT[0].ID, DateTime.Now);
            //        }
            //    }
            //}
            //catch (Exception e)
            //{
            //    MessageBox.Show("Could not open connection with database! " + e);
            //}

        }
        //Function to record process in  clients table in DB
        private void LogProcessClient()
        {
            //Database1DataSetTableAdapters.clientsTableAdapter logTA = new Database1DataSetTableAdapters.clientsTableAdapter();
            //Database1DataSet.clientsDataTable logDT = new Database1DataSet.clientsDataTable();

            //string period = fromTimePicker.Value.Date.ToShortDateString() + " to " + toTimePicker.Value.Date.ToShortDateString();

            //logTA.UpdateLastProcess(DateTime.Today, period, Convert.ToInt32(clientsCb.SelectedValue));
        }
        //Test function to create excel from template
        private void WriteExcel(DataSet ds, string file)
        {
            DataTable dt = new DataTable();
            dt = ds.Tables[0];
            //Save main spreadsheet 
            try
            {
                //Show status to user
                mainStatusLb.Text = "Loading data to spreadsheet...";
                toolStrip1.Refresh();

                string targetDir = file;

                FileInfo newFile = new FileInfo(@targetDir + ".xlsx");
                FileInfo template = new FileInfo(@"D:\Dionisio\Visual Studio\Avoidable cost files\\SavingTemplate.xlsx");

                using (ExcelPackage excelPack = new ExcelPackage(newFile, template))
                {
                    foreach (ExcelWorksheet oneWorkSheet in excelPack.Workbook.Worksheets)
                    {
                        oneWorkSheet.Cells[1, 1].Value = oneWorkSheet.Cells[1, 1].Value;
                    }
                    excelPack.Save();
                }
                //Open file
                Microsoft.Office.Interop.Excel.Application oXL = null;
                Microsoft.Office.Interop.Excel._Workbook oWB = null;
                Microsoft.Office.Interop.Excel._Worksheet oSheet = null;

                try
                {
                    oXL = new Microsoft.Office.Interop.Excel.Application();
                    oWB = oXL.Workbooks.Open(@targetDir + ".xlsx");
                    oSheet = String.IsNullOrEmpty("Plan1") ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["Plan1"];

                    //Loop through rows in ds
                    //Int32 sheetRow = 0;
                    for (int row = 2; row <= dt.Rows.Count + 1; row++)
                    {

                        //Loop through columns in ds
                        for (int col = 1; col <= 10; col++)
                        {
                            oSheet.Cells[row, col] = dt.Rows[row - 2][col - 1].ToString();
                        }
                    }

                    oWB.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    if (oWB != null)
                        oWB.Close();
                }
            }
            catch (Exception et)
            {
                MessageBox.Show("Could not open template. Erro: " + et);
            }
        }
        //ClientCB index changed event
        private void clientsCb_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (clientsCb.SelectedIndex > 0)
            {
                //double interval = GetClientPath().Interval;
                //DateTime dateNow = DateTime.Today;

                //fromTimePicker.Value = dateNow.AddDays(-interval);
                //toTimePicker.Value = dateNow;
                //avoidablesDGV.DataSource = null;
            }
        }
        //Function to export invoices by WIB date and project code from infoLCM
        private async Task<List<InvoiceToSearch>> ReturnIDsToSearchTest(string fileName, IProgress<string> progress)
        {
            //List of project codes
            //listProjCode = GetClientInfo();
            var listProjCode = GetPJCode();
            //List of avoidable cost logged from infoLCM//Class to store invoice data
            List<InvoiceToSearch> invoicesToSearch = new List<InvoiceToSearch>();
            string sheet1 = "";
            DateTime wibDate = DateTime.Today;

            try
            {
                //Code to read excell spreadsheet with all logged ACs from InfoLCM
                OleDbConnection conn = new OleDbConnection();
                conn.ConnectionString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended properties=""Excel 8.0;HDR=No;IMEX=1;MAXSCANROWS=0;ImportMixedTypes=Text;TypeGuessRows=0""", fileName);

                using (OleDbConnection connTwo = new OleDbConnection(conn.ConnectionString))
                {
                    connTwo.Open();
                    DataTable dtSchema = connTwo.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                }
                //Connection string
                //string connect = string.Format("SELECT * FROM [{0}] WHERE F6 = '{1}'", sheet1, listProjCode[t].ProjectCodes);
                string connect = string.Format("SELECT * FROM [{0}] ", sheet1);

                OleDbCommand command = new OleDbCommand();
                command.CommandText = connect;
                command.Connection = conn;

                System.Data.DataTable dtCustomers = new System.Data.DataTable();
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                var taskAdapter = Task.Factory.StartNew(() => adapter.Fill(dtCustomers));
                await Task.WhenAll(taskAdapter);

                dtCustomers.Columns[1].ColumnName = "ID";
                dtCustomers.Columns[5].ColumnName = "PROJECTCODE";
                dtCustomers.Columns[12].ColumnName = "WIBDATE";

                int rowNumber = 0;

                var tasks = Task.Factory.StartNew(() =>
                {
                    

                    for (int i = 0; i < listProjCode.Count; i++)
                    {
                        string selectCommand = "PROJECTCODE = '" + listProjCode[i].ProjectCodes + "'";

                        

                        DataRow[] listIDs = dtCustomers.Select(selectCommand);
                        foreach (DataRow row in listIDs)
                        {
                            rowNumber++;
                            ////Show status to user
                            progress.Report(string.Format("Searching for invoices by WIB date in project code {0} - Current row: {1}", listProjCode[i].ProjectCodes, rowNumber));

                            if (row[12].ToString() != "NULL")
                            {
                                DateTime.TryParse(row[12].ToString(), out wibDate);

                                if (wibDate >= fromTimePicker.Value && wibDate <= toTimePicker.Value)
                                {
                                    InvoiceToSearch invoiceInfoLCM = new InvoiceToSearch();

                                    invoiceInfoLCM.InvoiceID = row[1].ToString();
                                    invoiceInfoLCM.ProjectCode = row[5].ToString();
                                    invoiceInfoLCM.WibDate = row[12].ToString();

                                    invoicesToSearch.Add(invoiceInfoLCM);
                                }
                            }
                        }
                    }
                });
                await Task.WhenAll(tasks);

                MessageBox.Show("Number of invoices to search for avoidable cost: " + invoicesToSearch.Count.ToString() + " and " + rowNumber, "Important note");
                //Save number of invoices to search on result class
                processResult.InvoicesSearched = invoicesToSearch.Count;
            }
            catch (Exception exp)
            {
                MessageBox.Show("Could not open connection with database! " + exp);
            }
            return invoicesToSearch;
        }
        //Funtion to analize all lines of invoice
        private List<InvoiceInfo> CheckLines(string invoicePath, List<LineCode> lineCodes, string id, string projectCode)
        {
            //List of possible avoidable costs found and normal lines
            List<LineItem> avoidables = new List<LineItem>();
            List<LineItem> normalLines = new List<LineItem>();
            //List of ACs to return
            List<InvoiceInfo> avoidablesReturn = new List<InvoiceInfo>();

            string[] demandCodes = { "1202", "1502", "1552", "1553", "1203", "1503", "1292", "1293", "1294" };
            string line = "";

            StreamReader reader = new StreamReader(invoicePath);
            while ((line = reader.ReadLine()) != null)
            {
                //Check ok if avoidable costs line
                bool check = false;
                //Class lineItem
                LineItem lineItem = new LineItem();
                //Substring to get line code and cost                                
                lineItem.Code = line.Substring(17, 4);
                double lineCost = 0;
                double.TryParse(line.Substring(98, 11), out lineCost);
                lineItem.Cost = lineCost;
                lineItem.Flag = line.Substring(21, 1);
                double lineUnits = 0;
                double.TryParse(line.Substring(75, 10), out lineUnits);
                lineItem.Units = lineUnits;
                double lineRate = 0;
                double.TryParse(line.Substring(75, 10), out lineRate);
                lineItem.Rate = lineRate;

                //If line cost bigger than zero
                if (lineCost > 0)
                {
                    //If lineItem does not have a rate
                    if (lineItem.Rate == 0 && lineUnits > 0)
                    {
                        lineItem.Rate = lineItem.Cost / lineItem.Units;
                    }
                    //Check if charge is a possible avoidable cost
                    for (int c = 0; c < lineCodes.Count; c++)
                    {
                        if (lineCodes[c].Code == lineItem.Code)
                        {
                            lineItem.Category = lineCodes[c].Category;
                            lineItem.Subcategory = lineCodes[c].Subcategory;
                            lineItem.Description = lineCodes[c].Description;

                            avoidables.Add(lineItem);
                            check = true;
                        }
                    }
                }
                //If this line does not have an avoidable cost//add it to normal line list
                if (check == false)
                {
                    normalLines.Add(lineItem);
                }
            }
            reader.Close();

            //Check avoidable costs if found
            if (avoidables.Count > 0)
            {
                for (int t = 0; t < avoidables.Count; t++)
                {
                    bool confirmed = false;
                    //Class invoice info
                    InvoiceInfo avoidableCost = new InvoiceInfo();

                    if (avoidables[t].Category == "Demand related issues")
                    {
                        int check = 0;
                        //Check if there is another AC with the same category
                        for (int z = 0; z < avoidables.Count; z++)
                        {
                            if (z != t)
                            {
                                if (avoidables[z].Category == avoidables[t].Category)
                                {
                                    check += 1;
                                }
                            }
                        }
                        //If there is only one AC for demand
                        if (check == 0)
                        {
                            int index = 0;
                            int count = 0;
                            //Check how much charges related to demand in normal lines
                            for (int n = 0; n < normalLines.Count; n++)
                            {
                                if (normalLines[n].Cost > 0)
                                {
                                    for (int h = 0; h < demandCodes.Length; h++)
                                    {
                                        if (normalLines[n].Code == demandCodes[h])
                                        {
                                            count += 1;
                                            index = n;
                                        }
                                    }
                                }
                            }
                            //If only one charge for demand in normal lines and rate of avoidable cost bigger than in normal line
                            if (count == 1 && !string.IsNullOrWhiteSpace(avoidables[t].Rate.ToString()) && !string.IsNullOrWhiteSpace(normalLines[index].Rate.ToString()) && avoidables[t].Rate > normalLines[index].Rate)
                            {
                                confirmed = true;
                            }
                        }
                    }
                    else
                    {
                        confirmed = true;
                    }
                    //Add information to AC
                    avoidableCost.InvoiceID = id;
                    avoidableCost.ProjectCode = projectCode;
                    avoidableCost.Code = avoidables[t].Code;
                    avoidableCost.Cost = avoidables[t].Cost;
                    avoidableCost.Category = avoidables[t].Category;
                    avoidableCost.Subcategory = avoidables[t].Subcategory;
                    avoidableCost.Description = avoidables[t].Description;
                    avoidableCost.Confirmed = confirmed;

                    avoidablesReturn.Add(avoidableCost);
                }
            }
            return avoidablesReturn;
        }
        //Test1 remover depois
        private List<ClientInfo> GetPJCode()
        {
            List<ClientInfo> getWastePJ = new List<ClientInfo>();

            string[] projj = { "ENWA3", "ENWA7", "ENWA4", "ENWA8", "ENWAM", "ENWA1", "ENWA5", "ENWA2" };

            foreach (var item in projj)
            {
                ClientInfo pj = new ClientInfo();
                pj.ProjectCodes = item;
                pj.Country = "United States";

                getWastePJ.Add(pj);
            }

            return getWastePJ;
        }
        //Test2 remover depois
        private List<LineCode> GetLineCodeTest()
        {
            List<LineCode> lineCode = new List<LineCode>();

            string[] codes = { "1204", "1205", "1207" };

            foreach (var item in codes)
            {
                LineCode code = new LineCode();
                code.Code = item;
                code.Category = "Power factor";
                code.Subcategory = "Penalty for low power factor";
                code.Description = "test";

                lineCode.Add(code);
            }

            return lineCode;
        }
    }
}
