using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Xml.Linq;
using System.Net.Mail;
using static System.Collections.Specialized.BitVector32;
using SuperConvert;
using SuperConvert.Extensions;
using System.Data.SqlClient;
using System.IO;
using Microsoft.VisualBasic;
using System.ServiceModel.Syndication;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Windows.Forms.VisualStyles;
using System.Text.RegularExpressions;

namespace usersMailboxAccess
{
    public partial class frmHome : Form
    {
        // Create the main DataTable
        DataTable dataTable = new DataTable();
        DataTable dataTableTmp = new DataTable();
        string userType = "";
        string identityType = "";
        string username = Environment.UserName;

        public frmHome()
        {
            InitializeComponent();
        }

        private void btnGetAccess_Click(object sender, EventArgs e)
        {
            dataTable.Rows.Clear();

            // Disable the get access button so it wont mess up the process
            btnGetAccess.Enabled = false;

            //insert users into database after retreive
            insertUsers();

            //insert groups into database after retreive
            insertGroups();

            // receive all data function
            getAllAccessListData();

            // insert relationships in database after retrieve
            updateRelationships();

            MessageBox.Show("All data fetched from Exchange Online and Insertred into the Database!", "Notification", MessageBoxButtons.OK);

            // Re Enable get access button
            btnGetAccess.Enabled = true;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            // Exit
            Environment.Exit(0);
        }

        private void frmHome_Load(object sender, EventArgs e)
        {
            dbConn.conn.ConnectionString = dbConn.SQL_CONNECT;

            // Turn drop down search to receive no input
            cmbSearchCat.DropDownStyle = ComboBoxStyle.DropDownList;

            // create datatable columns
            dataTable.Columns.Add("UserAccount");
            dataTable.Columns.Add("Mailbox");
            dataTable.Columns.Add("AccessRights");
            dataTable.Columns.Add("InheritanceType");

            // create datatable columns
            dataTableTmp.Columns.Add("UserAccount");
            dataTableTmp.Columns.Add("Mailbox");
            dataTableTmp.Columns.Add("AccessRights");
            dataTableTmp.Columns.Add("InheritanceType");

            // create combobox items
            cmbSearchCat.Items.Add("Mailbox");
            cmbSearchCat.Items.Add("UserAccount");
            cmbSearchCat.Items.Add("AccessRights");
            //cmbSearchCat.Items.Add("InheritanceType");

            cmbSearchCat.SelectedIndex = 0;

            // allow multiple row selection for export purposes
            //dataGridView1.MultiSelect = true;
            //dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            if (username != "theodoros.h" && username != "aristos.a" && username != "vrionis.n")
            {
                btnGetAccess.Visible = false;
            }

            loadData();

            gotGet();

            dataGridView1.DataSource = dataTable;

            dataGridView1.Columns["InheritanceType"].Visible = false;

            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 430;
            dataGridView1.Columns[2].Width = 175;

            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[3].ReadOnly = true;

            string[] unbulleted = richTextBox1.Lines;
            string[] bulleted = new string[unbulleted.Length];

            for (int i = 0; i < bulleted.Length; i++)
            {
                if (i > 0)
                {
                    bulleted[i] = "\u2022" + unbulleted[i];
                }
                else
                {
                    bulleted[i] = unbulleted[i];
                }
            }

            richTextBox1.Lines = bulleted;
        }

        private void getAllAccessListData()
        {
            dataTable.Clear();

            // get all mailboxes and which users have full access
            getFullAccessList();

            // get all mailboxes and which users have send as access
            getSendAsList();

            // get all mailboxes and which users have send on behalf access
            getSendOnBehalfList();

            // get all mailboxes and which users have forwarding to access
            getForwardingToList();

            // get all security/distribution groups members
            getSGMembers();

            // give finalized datatable to datagrid and sort by user
            dataGridView1.DataSource = dataTable;
            dataGridView1.Sort(dataGridView1.Columns["UserAccount"], ListSortDirection.Ascending);
        }

        private void getFullAccessList()
        {
            // Powershell script that gets connects to exchange online with a certificate and gets all mailboxes with users who have full access on
            string scriptFullAccess = @"Set-ExecutionPolicy Unrestricted;
                            Import-Module ExchangeOnlineManagement
                            Connect-ExchangeOnline -CertificateFilePath ""C:\Temp\MailboxUsage.pfx"" -CertificatePassword (ConvertTo-SecureString -String ""certificate_pass"" -AsPlainText -Force) -AppID ""app_id_code"" -Organization "".onmicrosoft.com""
                            Get-Mailbox -ResultSize ""Unlimited"" | Get-MailboxPermission | where { ($_.AccessRights -eq ""FullAccess"") }";

            // start the runspace
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();
                //create pipeline
                Pipeline pipe = runspace.CreatePipeline();
                // give the script
                pipe.Commands.AddScript(scriptFullAccess);

                // invoke the pipeline to get the data
                var mailboxAccessData = pipe.Invoke();

                // run through the data and insert them into main datatable
                foreach (var mailboxAccessRow in mailboxAccessData)
                {
                    var row = dataTable.NewRow();
                    row["UserAccount"] = mailboxAccessRow.Properties["User"].Value.ToString();
                    row["Mailbox"] = mailboxAccessRow.Properties["Identity"].Value.ToString();
                    row["AccessRights"] = mailboxAccessRow.Properties["AccessRights"].Value.ToString();
                    row["InheritanceType"] = mailboxAccessRow.Properties["InheritanceType"].Value.ToString();
                    dataTable.Rows.Add(row);
                }
            }
        }

        private void getSendAsList()
        {
            // Powershell script that gets connects to exchange online with a certificate and gets all mailboxes with users who have send as access
            string scriptSendAs = @"Set-ExecutionPolicy Unrestricted;
                            Import-Module ExchangeOnlineManagement
                            Connect-ExchangeOnline -CertificateFilePath ""C:\Temp\MailboxUsage.pfx"" -CertificatePassword (ConvertTo-SecureString -String ""certificate_pass"" -AsPlainText -Force) -AppID ""app_id"" -Organization ""orgorg.onmicrosoft.com""
                            Get-Mailbox -resultsize unlimited | Get-RecipientPermission| where {($_.trustee -ne ""NT AUTHORITY\SELF"")}";

            // start the runspace
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();
                //create pipeline
                Pipeline pipe = runspace.CreatePipeline();
                pipe.Commands.AddScript(scriptSendAs);

                var mailboxAccessData = pipe.Invoke();

                // run through the data and insert them into main datatable
                foreach (var mailboxAccessRow in mailboxAccessData)
                {
                    var row = dataTable.NewRow();
                    row["UserAccount"] = mailboxAccessRow.Properties["Trustee"].Value.ToString();
                    row["Mailbox"] = mailboxAccessRow.Properties["Identity"].Value.ToString();
                    row["AccessRights"] = mailboxAccessRow.Properties["AccessRights"].Value.ToString();
                    row["InheritanceType"] = "";
                    dataTable.Rows.Add(row);
                }
            }
        }

        private void getSendOnBehalfList()
        {
            // Powershell script that gets connects to exchange online with a certificate and gets all mailboxes with users who have send on behalf access
            string scriptSendOnBehalf = @"Set-ExecutionPolicy Unrestricted;
                            Import-Module ExchangeOnlineManagement
                            Connect-ExchangeOnline -CertificateFilePath ""C:\Temp\MailboxUsage.pfx"" -CertificatePassword (ConvertTo-SecureString -String ""certificate_pass"" -AsPlainText -Force) -AppID ""app_id"" -Organization ""org.onmicrosoft.com""
                            Get-Mailbox | where {$_.GrantSendOnBehalfTo -ne $null} | select DisplayName,Name,Alias,UserPrincipalName,PrimarySmtpAddress,@{l='SendOnBehalfOf';e={$_.GrantSendOnBehalfTo -join "";""}}";

            // start the runspace
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();
                //create pipeline
                Pipeline pipe = runspace.CreatePipeline();
                pipe.Commands.AddScript(scriptSendOnBehalf);

                var mailboxAccessData = pipe.Invoke();

                // run through the data and insert them into main datatable
                foreach (var mailboxAccessRow in mailboxAccessData)
                {
                    string allUsers = mailboxAccessRow.Properties["SendOnBehalfOf"].Value.ToString();
                    // split on semicolon to get every user retreived
                    string[] allUsersArray = allUsers.Split(';');

                    // run through the data and insert them into main datatable
                    foreach (var idk in allUsersArray)
                    {
                        if (idk != ";")
                        {
                            var row = dataTable.NewRow();
                            row["UserAccount"] = idk;
                            row["Mailbox"] = mailboxAccessRow.Properties["Name"].Value.ToString();
                            row["AccessRights"] = "SendOnBehalf";
                            row["InheritanceType"] = "";
                            dataTable.Rows.Add(row);
                        }
                    }
                }
            }
        }

        private void getForwardingToList()
        {
            // Powershell script that gets connects to exchange online with a certificate and gets all mailboxes that have forwarding access and to which groups/users
            string scriptForwarding = @"Set-ExecutionPolicy Unrestricted;
                            Import-Module ExchangeOnlineManagement
                            Connect-ExchangeOnline -CertificateFilePath ""C:\Temp\MailboxUsage.pfx"" -CertificatePassword (ConvertTo-SecureString -String ""certificate_pass"" -AsPlainText -Force) -AppID ""app_id"" -Organization ""org.onmicrosoft.com""
                            Get-Mailbox -ResultSize Unlimited | Where {($_.ForwardingAddress -ne $Null) -or ($_.ForwardingsmtpAddress -ne $Null)} | Select DisplayName, Name, ForwardingAddress,ForwardingsmtpAddress, DeliverToMailboxAndForward";

            //start runspace
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();
                //create pipeline
                Pipeline pipe = runspace.CreatePipeline();
                pipe.Commands.AddScript(scriptForwarding);

                var mailboxAccessData = pipe.Invoke();

                // run through the data and insert them into main datatable
                foreach (var mailboxAccessRow in mailboxAccessData)
                {
                    // if both are null then do nothing
                    // else if both not null then insert both names in the User Column
                    // else if one is not null then enter the one that is not null in User Column
                    if (mailboxAccessRow.Properties["ForwardingAddress"].Value is null && mailboxAccessRow.Properties["ForwardingSmtpAddress"].Value is null)
                    {

                    }
                    else if (mailboxAccessRow.Properties["ForwardingAddress"].Value is not null && mailboxAccessRow.Properties["ForwardingSmtpAddress"].Value is not null)
                    {
                        var row = dataTable.NewRow();
                        row["UserAccount"] = mailboxAccessRow.Properties["ForwardingAddress"].Value.ToString() + ", " + mailboxAccessRow.Properties["ForwardingSmtpAddress"].Value.ToString();
                        row["Mailbox"] = mailboxAccessRow.Properties["Name"].Value.ToString();
                        row["AccessRights"] = "Forwarding";

                        bool test = (bool)mailboxAccessRow.Properties["DeliverToMailboxAndForward"].Value;

                        if (test != true)
                        {
                            row["InheritanceType"] = "Deliver To Mailbox And Forward = No";
                        }
                        else
                        {
                            row["InheritanceType"] = "Deliver To Mailbox And Forward = Yes";
                        }

                        dataTable.Rows.Add(row);
                    }
                    else if (mailboxAccessRow.Properties["ForwardingAddress"].Value is not null || mailboxAccessRow.Properties["ForwardingSmtpAddress"].Value is not null)
                    {
                        var row = dataTable.NewRow();
                        if (mailboxAccessRow.Properties["ForwardingAddress"].Value is not null)
                        {
                            row["UserAccount"] = mailboxAccessRow.Properties["ForwardingAddress"].Value.ToString();
                        }
                        else if (mailboxAccessRow.Properties["ForwardingSmtpAddress"].Value is not null)
                        {
                            row["UserAccount"] = mailboxAccessRow.Properties["ForwardingSmtpAddress"].Value.ToString();
                        }
                        row["Mailbox"] = mailboxAccessRow.Properties["Name"].Value.ToString();
                        row["AccessRights"] = "Forwarding";

                        bool test = (bool)mailboxAccessRow.Properties["DeliverToMailboxAndForward"].Value;

                        if (test != true)
                        {
                            row["InheritanceType"] = "Deliver To Mailbox And Forward = No";
                        }
                        else
                        {
                            row["InheritanceType"] = "Deliver To Mailbox And Forward = Yes";
                        }

                        dataTable.Rows.Add(row);
                    }
                }
            }
        }

        private void getSGMembers()
        {
            // Powershell script that connects to exchange online with a certificate and gets all users in every security group/distribution list
            // and if there are other sgs/dls in the specific group nest loop and get the users from those groups also and print a list of the sg/dl and its users
            string scriptSGMembers = @"Set-ExecutionPolicy Unrestricted;
                            Import-Module ExchangeOnlineManagement
                            Connect-ExchangeOnline -CertificateFilePath ""C:\Temp\MailboxUsage.pfx"" -CertificatePassword (ConvertTo-SecureString -String ""certificate_pass"" -AsPlainText -Force) -AppID ""app_id"" -Organization ""org.onmicrosoft.com""
                            $Groups = Get-Group -ResultSize Unlimited | Select DisplayName, Identity, ManagedBy, GroupType, RecipientTypeDetails, RecipientType,PrimarySmtpAddress, @{l='Members';e={$_.Members -join "";""}} | where { ($_.DisplayName -ne """") }
                            $global = """"
                            # Define a function to recursively get group members
                            function Get-NestedGroupMembers {
                                param (
                                    [Parameter(Mandatory=$true)]
                                    [string]$GroupName
                                )
                                $GroupMembers = Get-DistributionGroupMember -Identity $GroupName -ResultSize Unlimited                              
                                foreach ($Member in $GroupMembers) {
                                    if ($Member.RecipientType -eq ""UserMailbox"") {                                        
                                        Write-Output $Member
                                    } elseif ($Member.RecipientType -eq ""MailUniversalDistributionGroup"" -or $Member.RecipientType -eq ""MailUniversalSecurityGroup"") {
                                        Get-NestedGroupMembers -GroupName $Member.Identity
                                    }
                                }
                            }

                            # Iterate through each group and get its nested members
                            foreach ($Group in $Groups) {
                                Write-Host ""Group:"" $Group.Identity
                                $global = $Group.Identity
                                Get-DistributionGroup -Identity $global | SELECT Name, PrimarySmtpAddress, GroupType, RecipientType
                                Get-NestedGroupMembers -GroupName $Group.Identity                               
                            }";

            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();
                Pipeline pipe = runspace.CreatePipeline();
                pipe.Commands.AddScript(scriptSGMembers);

                var mailboxAccessData = pipe.Invoke();

                string groupName = "";

                // run through the data retreived and get the group first then run through its users and prin the group in the identity
                // also clean duplicate values for every security group
                foreach (var mailboxAccessRow in mailboxAccessData)
                {
                    string groupTypeOrUserMailbox = mailboxAccessRow.Properties["RecipientType"].Value.ToString();

                    if (groupTypeOrUserMailbox == "UserMailbox")
                    {
                        bool found = false;

                        foreach (DataRow ifExistsRow in dataTable.Rows)
                        {
                            if (ifExistsRow["UserAccount"].ToString().Contains(mailboxAccessRow.Properties["PrimarySmtpAddress"].Value.ToString()) && ifExistsRow["Mailbox"].ToString() == groupName)
                            {
                                found = true;
                                goto goHere;
                            }
                        }

                    goHere:

                        if (found != false)
                        {

                        }
                        else
                        {
                            var row = dataTable.NewRow();
                            row["UserAccount"] = mailboxAccessRow.Properties["Name"].Value.ToString();
                            row["Mailbox"] = groupName;
                            row["AccessRights"] = "GroupMember";
                            row["InheritanceType"] = "";
                            dataTable.Rows.Add(row);
                        }
                    }
                    else
                    {
                        groupName = mailboxAccessRow.Properties["Name"].Value.ToString();
                    }
                }
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            //if datagrid doesnt have rows then show a message
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Table is empty you cannot search for data.", "Important!", MessageBoxButtons.OK);
            }
            else
            {
                string searchValue = txtSearch.Text;

                //dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                // create a new filtered rows array
                DataRow[] filteredRows;

                // create a new temp datatable with filtered rows
                DataTable tempSearch = new DataTable();

                tempSearch.Columns.Add("UserAccount");
                tempSearch.Columns.Add("Mailbox");
                tempSearch.Columns.Add("AccessRights");
                tempSearch.Columns.Add("InheritanceType");

                // check if match case is checked 
                if (chkMatchCase.Checked == true && chkRecipients.Checked == true)
                {
                    // if it is depends on combo box search exactly what the user inserted in textbox
                    if (cmbSearchCat.SelectedIndex == 0)
                    {
                        filteredRows = dataTable.Select("Mailbox = '" + searchValue + "' AND (AccessRights = 'GroupMembers' OR AccessRights = 'MailboxForwardingToGroup')");
                    }
                    else if (cmbSearchCat.SelectedIndex == 1)
                    {
                        filteredRows = dataTable.Select("UserAccount = '" + searchValue + "' AND (AccessRights = 'GroupMembers' OR AccessRights = 'MailboxForwardingToGroup')");
                    }
                    else if (cmbSearchCat.SelectedIndex == 2)
                    {
                        filteredRows = dataTable.Select("AccessRights = '" + searchValue + "'");
                    }
                    else
                    {
                        filteredRows = dataTable.Select("InheritanceType = '" + searchValue + "'");
                    }
                }
                else if (chkMatchCase.Checked == true && chkRecipients.Checked == false)
                {
                    // if it is depends on combo box search exactly what the user inserted in textbox
                    if (cmbSearchCat.SelectedIndex == 0)
                    {
                        filteredRows = dataTable.Select("Mailbox = '" + searchValue + "'");
                    }
                    else if (cmbSearchCat.SelectedIndex == 1)
                    {
                        filteredRows = dataTable.Select("UserAccount = '" + searchValue + "'");
                    }
                    else if (cmbSearchCat.SelectedIndex == 2)
                    {
                        filteredRows = dataTable.Select("AccessRights = '" + searchValue + "'");
                    }
                    else
                    {
                        filteredRows = dataTable.Select("InheritanceType = '" + searchValue + "'");
                    }
                }
                else if (chkMatchCase.Checked == false && chkRecipients.Checked == true)
                {
                    // if it is depends on combo box search anything that contains what the user inserted in textbox
                    if (cmbSearchCat.SelectedIndex == 0)
                    {
                        filteredRows = dataTable.Select("Mailbox LIKE '%" + searchValue + "%' AND (AccessRights = 'GroupMembers' OR AccessRights = 'MailboxForwardingToGroup')");
                    }
                    else if (cmbSearchCat.SelectedIndex == 1)
                    {
                        filteredRows = dataTable.Select("UserAccount LIKE '%" + searchValue + "%' AND (AccessRights = 'GroupMembers' OR AccessRights = 'MailboxForwardingToGroup')");
                    }
                    else if (cmbSearchCat.SelectedIndex == 2)
                    {
                        filteredRows = dataTable.Select("AccessRights LIKE '%" + searchValue + "%'");
                    }
                    else
                    {
                        filteredRows = dataTable.Select("InheritanceType LIKE '%" + searchValue + "%'");
                    }
                }
                else
                {
                    // if it is depends on combo box search anything that contains what the user inserted in textbox
                    if (cmbSearchCat.SelectedIndex == 0)
                    {
                        filteredRows = dataTable.Select("Mailbox LIKE '%" + searchValue + "%'");
                    }
                    else if (cmbSearchCat.SelectedIndex == 1)
                    {
                        filteredRows = dataTable.Select("UserAccount LIKE '%" + searchValue + "%'");
                    }
                    else if (cmbSearchCat.SelectedIndex == 2)
                    {
                        filteredRows = dataTable.Select("AccessRights LIKE '%" + searchValue + "%'");
                    }
                    else
                    {
                        filteredRows = dataTable.Select("InheritanceType LIKE '%" + searchValue + "%'");
                    }
                }

                // insert filtered rows to temp datatable
                foreach (var dr in filteredRows)
                {
                    var row = tempSearch.NewRow();
                    row["UserAccount"] = dr.ItemArray[0];
                    row["Mailbox"] = dr.ItemArray[1];
                    row["AccessRights"] = dr.ItemArray[2];
                    row["InheritanceType"] = dr.ItemArray[3];
                    tempSearch.Rows.Add(row);
                }

                // assign datasource
                dataGridView1.DataSource = tempSearch;
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            // reset view to main datatable
            dataGridView1.DataSource = dataTable;

            cmbSearchCat.SelectedIndex = 0;

            txtSearch.Text = "";
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //Create Data Row and Table
            int i = 0;
            DataRow datarow = null;
            DataTable expDT = new DataTable();

            expDT.Columns.Add("UserAccount");
            expDT.Columns.Add("Mailbox");
            expDT.Columns.Add("AccessRights");
            expDT.Columns.Add("InheritanceType");

            // Loop through datagrid and get selected rows in a temp dt
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Selected)
                {
                    try
                    {
                        datarow = ((DataRowView)row.DataBoundItem).Row;
                        expDT.ImportRow(datarow);
                        i += 1;
                    }
                    catch
                    {

                    }
                }
            }

            if (i == 0)
            {
                MessageBox.Show("Please select full rows to export into CSV.", "Important Message!", MessageBoxButtons.OK);
                return;
            }

            // open file save dialog box
            SaveFileDialog oSaveFileDialog = new SaveFileDialog();
            oSaveFileDialog.Filter = "CSV|*.csv";

            // if okay is pressed then save file to user's specified directory
            if (oSaveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string pathCSVDirectory = Path.GetDirectoryName(oSaveFileDialog.FileName);
                string pathCSVFileName = Path.GetFileNameWithoutExtension(oSaveFileDialog.FileName);

                expDT.ToCsv(pathCSVDirectory, pathCSVFileName);

                MessageBox.Show("CSV Exported Successfully!", "Notification", MessageBoxButtons.OK);
            }
        }

        private void insertUsers()
        {
            // Powershell script that gets connects to exchange online with a certificate and gets all mailboxes with users who have full access on
            string scriptFullAccess = @"Set-ExecutionPolicy Unrestricted;
                            Import-Module ExchangeOnlineManagement
                            Connect-ExchangeOnline -CertificateFilePath ""C:\Temp\MailboxUsage.pfx"" -CertificatePassword (ConvertTo-SecureString -String ""certificate_pass"" -AsPlainText -Force) -AppID ""app_id"" -Organization ""org.onmicrosoft.com""
                            Get-Mailbox -ResultSize ""Unlimited"" | Select DisplayName, Name, PrimarySmtpAddress, RecipientType | Where {($_.RecipientType -eq ""UserMailbox"")}";

            // start the runspace
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();
                //create pipeline
                Pipeline pipe = runspace.CreatePipeline();
                // give the script
                pipe.Commands.AddScript(scriptFullAccess);

                // invoke the pipeline to get the data
                var mailboxAccessData = pipe.Invoke();

                // run through the data and insert them into main datatable
                foreach (var mailboxAccessRow in mailboxAccessData)
                {
                    string userExists = "";
                    string emailExists = "";
                    string userName = mailboxAccessRow.Properties["Name"].Value.ToString();
                    //bool containsNumeric = Regex.IsMatch(userName, @"\d");

                    //if (containsNumeric)
                    //{
                    //    userName = mailboxAccessRow.Properties["DisplayName"].Value.ToString();
                    //}

                    // check if uer exists in table users
                    try
                    {
                        string SQL = "SELECT * FROM mailboxes.dbo.users WHERE userName ='" + userName + "'";
                        dbConn.sqlCmd.Connection = dbConn.conn;
                        dbConn.conn.Close();
                        dbConn.conn.Open();
                        dbConn.sqlCmd.CommandText = SQL;
                        dbConn.sqlRdr = dbConn.sqlCmd.ExecuteReader();

                        if (dbConn.sqlRdr.HasRows)
                        {
                            while (dbConn.sqlRdr.Read())
                            {
                                userExists = dbConn.sqlRdr.GetValue(dbConn.sqlRdr.GetOrdinal("userName")).ToString();
                                emailExists = dbConn.sqlRdr.GetValue(dbConn.sqlRdr.GetOrdinal("userEmail")).ToString();
                            }
                        }
                    }
                    catch
                    {

                    }

                    // if user doesnt exist insert it with their email
                    if ((emailExists == "" || emailExists is null) && (userExists == "" || userExists is null))
                    {
                        string SQL = "INSERT INTO mailboxes.dbo.users(userName, userEmail) VALUES (@userName, @userEmail)";

                        dbConn.sqlCmd.Parameters.Clear();

                        try
                        {

                            dbConn.sqlCmd.Parameters.AddWithValue("@userName", userName);
                        }
                        catch
                        {
                        }

                        try
                        {
                            dbConn.sqlCmd.Parameters.AddWithValue("@userEmail", mailboxAccessRow.Properties["PrimarySmtpAddress"].Value.ToString());
                        }
                        catch
                        {

                        }

                        try
                        {
                            dbConn.sqlCmd.Connection = dbConn.conn;
                            dbConn.conn.Close();
                            dbConn.conn.Open();
                            dbConn.sqlCmd.CommandText = SQL;
                            dbConn.sqlCmd.ExecuteNonQuery();
                            dbConn.sqlCmd.Parameters.Clear();
                        }
                        catch
                        {

                        }
                    }
                    // if they exist then update their email
                    else if ((emailExists != "" || emailExists is not null) && (userExists != "" || userExists is not null))
                    {
                        if (emailExists != mailboxAccessRow.Properties["PrimarySmtpAddress"].Value.ToString())
                        {
                            string SQL = "UPDATE mailboxes.dbo.users SET userEmail=@userEmail WHERE userName = '" + userName + "'";

                            dbConn.sqlCmd.Parameters.Clear();

                            try
                            {
                                dbConn.sqlCmd.Parameters.AddWithValue("@userEmail", mailboxAccessRow.Properties["PrimarySmtpAddress"].Value.ToString());
                            }
                            catch
                            {

                            }

                            try
                            {
                                dbConn.sqlCmd.Connection = dbConn.conn;
                                dbConn.conn.Close();
                                dbConn.conn.Open();
                                dbConn.sqlCmd.CommandText = SQL;
                                dbConn.sqlCmd.ExecuteNonQuery();
                                dbConn.sqlCmd.Parameters.Clear();
                            }
                            catch
                            {

                            }
                        }
                    }
                }
            }
        }

        private void insertGroups()
        {
            // Powershell script that gets connects to exchange online with a certificate and gets all mailboxes with users who have full access on
            string scriptFullAccess = @"Set-ExecutionPolicy Unrestricted;
                            Import-Module ExchangeOnlineManagement
                            Connect-ExchangeOnline -CertificateFilePath ""C:\Temp\MailboxUsage.pfx"" -CertificatePassword (ConvertTo-SecureString -String ""certificate_pass"" -AsPlainText -Force) -AppID ""app_id"" -Organization ""org.onmicrosoft.com""
                            Get-Group -ResultSize Unlimited | Select DisplayName, Name, WindowsEmailAddress";

            // start the runspace
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();
                //create pipeline
                Pipeline pipe = runspace.CreatePipeline();
                // give the script
                pipe.Commands.AddScript(scriptFullAccess);

                // invoke the pipeline to get the data
                var mailboxAccessData = pipe.Invoke();

                // run through the data and insert them into main datatable
                foreach (var mailboxAccessRow in mailboxAccessData)
                {
                    string groupExists = "";
                    string emailExists = "";

                    // check if group exists in groups
                    try
                    {
                        string SQL = "SELECT * FROM mailboxes.dbo.groups WHERE groupName ='" + mailboxAccessRow.Properties["Name"].Value.ToString() + "'";
                        dbConn.sqlCmd.Connection = dbConn.conn;
                        dbConn.conn.Close();
                        dbConn.conn.Open();
                        dbConn.sqlCmd.CommandText = SQL;
                        dbConn.sqlRdr = dbConn.sqlCmd.ExecuteReader();

                        if (dbConn.sqlRdr.HasRows)
                        {
                            while (dbConn.sqlRdr.Read())
                            {
                                groupExists = dbConn.sqlRdr.GetValue(dbConn.sqlRdr.GetOrdinal("groupName")).ToString();
                                emailExists = dbConn.sqlRdr.GetValue(dbConn.sqlRdr.GetOrdinal("groupEmail")).ToString();
                            }
                        }
                    }
                    catch
                    {

                    }

                    //if group doesnt exist then insert
                    if ((emailExists == "" || emailExists is null) && (groupExists == "" || groupExists is null))
                    {
                        string SQL = "INSERT INTO mailboxes.dbo.groups(groupName, groupEmail) VALUES (@groupName, @groupEmail)";

                        dbConn.sqlCmd.Parameters.Clear();

                        try
                        {
                            dbConn.sqlCmd.Parameters.AddWithValue("@groupName", mailboxAccessRow.Properties["Name"].Value.ToString());
                        }
                        catch
                        {

                        }

                        try
                        {
                            dbConn.sqlCmd.Parameters.AddWithValue("@groupEmail", mailboxAccessRow.Properties["WindowsEmailAddress"].Value.ToString());
                        }
                        catch
                        {

                        }

                        try
                        {
                            dbConn.sqlCmd.Connection = dbConn.conn;
                            dbConn.conn.Close();
                            dbConn.conn.Open();
                            dbConn.sqlCmd.CommandText = SQL;
                            dbConn.sqlCmd.ExecuteNonQuery();
                            dbConn.sqlCmd.Parameters.Clear();
                        }
                        catch
                        {

                        }
                    }
                    // if it does then update email
                    else if ((emailExists != "" || emailExists is not null) && (groupExists != "" || groupExists is not null))
                    {
                        if (emailExists != mailboxAccessRow.Properties["WindowsEmailAddress"].Value.ToString())
                        {
                            string SQL = "UPDATE mailboxes.dbo.groups SET groupEmail=@groupEmail WHERE groupName = '" + mailboxAccessRow.Properties["Name"].Value.ToString() + "'";

                            dbConn.sqlCmd.Parameters.Clear();

                            try
                            {
                                dbConn.sqlCmd.Parameters.AddWithValue("@groupEmail", mailboxAccessRow.Properties["WindowsEmailAddress"].Value.ToString());
                            }
                            catch
                            {

                            }

                            try
                            {
                                dbConn.sqlCmd.Connection = dbConn.conn;
                                dbConn.conn.Close();
                                dbConn.conn.Open();
                                dbConn.sqlCmd.CommandText = SQL;
                                dbConn.sqlCmd.ExecuteNonQuery();
                                dbConn.sqlCmd.Parameters.Clear();
                            }
                            catch
                            {

                            }
                        }
                    }
                }
            }
        }

        private void updateRelationships()
        {
            int accessRightID = 0;

            // clear all access tables
            deleteForwarding();
            deleteFullAccess();
            deleteGroupMembers();
            deleteSendAs();
            deleteSendOnBehalf();

            // loop on datatable data
            foreach (DataRow row in dataTable.Rows)
            {

                string SQL = "";
                string userID = "";
                string groupID = "";

                if (row["AccessRights"].ToString() != "Forwarding")
                {
                    userID = getUserID(row["UserAccount"].ToString(), 0);
                    groupID = getGroupID(row["Mailbox"].ToString(), 0);
                }
                else
                {
                    userID = getUserID(row["Mailbox"].ToString(), 0);
                    groupID = getGroupID(row["UserAccount"].ToString(), 0);
                }

                // depends on the access of each row insert into the appropriate table
                if (row["AccessRights"].ToString() == "FullAccess")
                {
                    SQL = "INSERT INTO mailboxes.dbo.fullaccess(user_ID, identity_ID, accessRights_ID, inheritanceType, identityType, userType) VALUES (@user_ID, @identity_ID, @accessRights_ID, @inheritanceType, @identityType, @userType)";
                    accessRightID = 1;
                }
                else if (row["AccessRights"].ToString() == "SendAs")
                {
                    SQL = "INSERT INTO mailboxes.dbo.sendas(user_ID, identity_ID, accessRights_ID, inheritanceType, identityType, userType) VALUES (@user_ID, @identity_ID, @accessRights_ID, @inheritanceType, @identityType, @userType)";
                    accessRightID = 2;
                }
                else if (row["AccessRights"].ToString() == "SendOnBehalf")
                {
                    SQL = "INSERT INTO mailboxes.dbo.sendonbehalf(user_ID, identity_ID, accessRights_ID, inheritanceType, identityType, userType) VALUES (@user_ID, @identity_ID, @accessRights_ID, @inheritanceType, @identityType, @userType)";
                    accessRightID = 3;
                }
                else if (row["AccessRights"].ToString() == "Forwarding")
                {
                    SQL = "INSERT INTO mailboxes.dbo.forwarding(user_ID, identity_ID, accessRights_ID, inheritanceType, identityType, userType) VALUES (@user_ID, @identity_ID, @accessRights_ID, @inheritanceType, @identityType, @userType)";
                    accessRightID = 4;
                }
                else if (row["AccessRights"].ToString() == "GroupMember")
                {
                    SQL = "INSERT INTO mailboxes.dbo.groupmembers(user_ID, identity_ID, accessRights_ID, inheritanceType, identityType, userType) VALUES (@user_ID, @identity_ID, @accessRights_ID, @inheritanceType, @identityType, @userType)";
                    accessRightID = 5;
                }

                dbConn.sqlCmd.Parameters.Clear();

                try
                {
                    dbConn.sqlCmd.Parameters.AddWithValue("@user_ID", userID);
                }
                catch
                {

                }

                try
                {
                    dbConn.sqlCmd.Parameters.AddWithValue("@identity_ID", groupID);
                }
                catch
                {

                }

                try
                {
                    dbConn.sqlCmd.Parameters.AddWithValue("@accessRights_ID", accessRightID.ToString());
                }
                catch
                {

                }

                try
                {
                    dbConn.sqlCmd.Parameters.AddWithValue("@inheritanceType", row["InheritanceType"].ToString());
                }
                catch
                {

                }

                try
                {
                    dbConn.sqlCmd.Parameters.AddWithValue("@identityType", identityType);
                }
                catch
                {

                }

                try
                {
                    dbConn.sqlCmd.Parameters.AddWithValue("@userType", userType);
                }
                catch
                {

                }

                try
                {
                    dbConn.sqlCmd.Connection = dbConn.conn;
                    dbConn.conn.Close();
                    dbConn.conn.Open();
                    dbConn.sqlCmd.CommandText = SQL;
                    dbConn.sqlCmd.ExecuteNonQuery();
                    dbConn.sqlCmd.Parameters.Clear();
                }
                catch
                {

                }

            }
        }

        // delete from table forwarding
        private void deleteForwarding()
        {
            string SQL = "DELETE FROM mailboxes.dbo.forwarding";

            dbConn.sqlCmd.Parameters.Clear();

            try
            {
                dbConn.sqlCmd.Connection = dbConn.conn;
                dbConn.conn.Close();
                dbConn.conn.Open();
                dbConn.sqlCmd.CommandText = SQL;
                dbConn.sqlCmd.ExecuteNonQuery();
                dbConn.sqlCmd.Parameters.Clear();
            }
            catch
            {

            }
        }

        // delete from table fullaccess
        private void deleteFullAccess()
        {
            string SQL = "DELETE FROM mailboxes.dbo.fullaccess";

            dbConn.sqlCmd.Parameters.Clear();

            try
            {
                dbConn.sqlCmd.Connection = dbConn.conn;
                dbConn.conn.Close();
                dbConn.conn.Open();
                dbConn.sqlCmd.CommandText = SQL;
                dbConn.sqlCmd.ExecuteNonQuery();
                dbConn.sqlCmd.Parameters.Clear();
            }
            catch
            {

            }
        }

        // delete from table groupmembers
        private void deleteGroupMembers()
        {
            string SQL = "DELETE FROM mailboxes.dbo.groupmembers";

            dbConn.sqlCmd.Parameters.Clear();

            try
            {
                dbConn.sqlCmd.Connection = dbConn.conn;
                dbConn.conn.Close();
                dbConn.conn.Open();
                dbConn.sqlCmd.CommandText = SQL;
                dbConn.sqlCmd.ExecuteNonQuery();
                dbConn.sqlCmd.Parameters.Clear();
            }
            catch
            {

            }
        }

        // delete from table sendas
        private void deleteSendAs()
        {
            string SQL = "DELETE FROM mailboxes.dbo.sendas";

            dbConn.sqlCmd.Parameters.Clear();

            try
            {
                dbConn.sqlCmd.Connection = dbConn.conn;
                dbConn.conn.Close();
                dbConn.conn.Open();
                dbConn.sqlCmd.CommandText = SQL;
                dbConn.sqlCmd.ExecuteNonQuery();
                dbConn.sqlCmd.Parameters.Clear();
            }
            catch
            {

            }
        }

        // delete from table sendonbehalf
        private void deleteSendOnBehalf()
        {
            string SQL = "DELETE FROM mailboxes.dbo.sendonbehalf";

            dbConn.sqlCmd.Parameters.Clear();

            try
            {
                dbConn.sqlCmd.Connection = dbConn.conn;
                dbConn.conn.Close();
                dbConn.conn.Open();
                dbConn.sqlCmd.CommandText = SQL;
                dbConn.sqlCmd.ExecuteNonQuery();
                dbConn.sqlCmd.Parameters.Clear();
            }
            catch
            {

            }
        }

        // delete from table sendonbehalf
        private string getUserID(string userDetailParam, int temp)
        {
            string dbUserID = "";
            string SQL = "";
            // check if uer exists in table users
            try
            {
                if (userDetailParam.Contains("@"))
                {
                    SQL = "SELECT * FROM mailboxes.dbo.users WHERE userEmail ='" + userDetailParam + "'";
                }
                else
                {
                    SQL = "SELECT * FROM mailboxes.dbo.users WHERE userName ='" + userDetailParam + "'";
                }
                dbConn.sqlCmd.Connection = dbConn.conn;
                dbConn.conn.Close();
                dbConn.conn.Open();
                dbConn.sqlCmd.CommandText = SQL;
                dbConn.sqlRdr = dbConn.sqlCmd.ExecuteReader();

                if (dbConn.sqlRdr.HasRows)
                {
                    while (dbConn.sqlRdr.Read())
                    {
                        try
                        {
                            dbUserID = dbConn.sqlRdr.GetValue(dbConn.sqlRdr.GetOrdinal("userID")).ToString();
                        }
                        catch
                        {
                            dbUserID = "";
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }

            // if id is null then check if its a group and assign identity type of user or group
            if (dbUserID == "" && temp == 0)
            {
                dbUserID = getGroupID(userDetailParam, 1);
                userType = "Group";
            }
            else
            {
                userType = "User";
            }

            return dbUserID;
        }

        // delete from table sendonbehalf
        private string getGroupID(string groupDetailParam, int temp)
        {
            string dbGroupID = "";
            string SQL = "";

            // check if uer exists in table users
            try
            {
                // check if username or email
                if (groupDetailParam.Contains("@"))
                {
                    SQL = "SELECT * FROM mailboxes.dbo.groups WHERE groupEmail ='" + groupDetailParam + "'";
                }
                else
                {
                    SQL = "SELECT * FROM mailboxes.dbo.groups WHERE groupName ='" + groupDetailParam + "'";
                }
                dbConn.sqlCmd.Connection = dbConn.conn;
                dbConn.conn.Close();
                dbConn.conn.Open();
                dbConn.sqlCmd.CommandText = SQL;
                dbConn.sqlRdr = dbConn.sqlCmd.ExecuteReader();

                if (dbConn.sqlRdr.HasRows)
                {
                    while (dbConn.sqlRdr.Read())
                    {
                        try
                        {
                            dbGroupID = dbConn.sqlRdr.GetValue(dbConn.sqlRdr.GetOrdinal("groupID")).ToString();
                        }
                        catch
                        {
                            dbGroupID = "";
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }

            // if id is null then check if its a User and assign identity type of user or group
            if (dbGroupID == "" && temp == 0)
            {
                dbGroupID = getUserID(groupDetailParam, 1);
                identityType = "User";
            }
            else
            {
                identityType = "Group";
            }

            return dbGroupID;
        }

        private void loadData()
        {
            loadFullAccess();
            loadSendAs();
            loadSendOnBehalf();
            loadForwarding();
            loadGroupMembers();
        }

        private void loadFullAccess()
        {
            string SQL = "";

            if (username != "theodoros.h" && username != "arisadmin" && username != "vrionis.n")
            {
                SQL = "SELECT CASE WHEN (fullaccess.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE fullaccess.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE fullaccess.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (fullaccess.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE fullaccess.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE fullaccess.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE fullaccess.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.fullaccess WHERE user_ID NOT IN (0, 242, 109, 52, 32, 256, 142, 28, 106) and identity_ID <> 0";
            }
            else
            {
                SQL = "SELECT CASE WHEN (fullaccess.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE fullaccess.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE fullaccess.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (fullaccess.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE fullaccess.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE fullaccess.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE fullaccess.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.fullaccess WHERE user_ID <> 0 and identity_ID <> 0";
            }

            DataTable dtFullAccess = dbConn.retrieveDB(SQL);

            foreach (DataRow row in dtFullAccess.Rows)
            {
                var rowAuth = dataTable.NewRow();
                rowAuth["UserAccount"] = row["UserAccount"];
                rowAuth["Mailbox"] = row["Mailbox"];
                rowAuth["AccessRights"] = row["AccessRights"];
                rowAuth["InheritanceType"] = row["InheritanceType"];
                dataTable.ImportRow(row);
            }
        }

        private void loadSendAs()
        {
            string SQL = "";

            if (username != "theodoros.h" && username != "arisadmin" && username != "vrionis.n")
            {
                SQL = "SELECT CASE WHEN (sendas.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE sendas.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE sendas.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (sendas.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE sendas.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE sendas.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE sendas.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.sendas WHERE user_ID NOT IN (0, 242, 109, 52, 32, 256, 142, 28, 106) and identity_ID <> 0";
            }
            else
            {
                SQL = "SELECT CASE WHEN (sendas.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE sendas.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE sendas.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (sendas.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE sendas.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE sendas.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE sendas.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.sendas WHERE user_ID <> 0 and identity_ID <> 0";
            }

            DataTable dtFullAccess = dbConn.retrieveDB(SQL);

            foreach (DataRow row in dtFullAccess.Rows)
            {
                var rowAuth = dataTable.NewRow();
                rowAuth["UserAccount"] = row["UserAccount"];
                rowAuth["Mailbox"] = row["Mailbox"];
                rowAuth["AccessRights"] = row["AccessRights"];
                rowAuth["InheritanceType"] = row["InheritanceType"];
                dataTable.ImportRow(row);
            }
        }

        private void loadSendOnBehalf()
        {
            string SQL = "";

            if (username != "theodoros.h" && username != "arisadmin" && username != "vrionis.n")
            {
                SQL = "SELECT CASE WHEN (sendonbehalf.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE sendonbehalf.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE sendonbehalf.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (sendonbehalf.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE sendonbehalf.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE sendonbehalf.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE sendonbehalf.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.sendonbehalf WHERE user_ID NOT IN (0, 242, 109, 52, 32, 256, 142, 28, 106) and identity_ID <> 0";
            }
            else
            {
                SQL = "SELECT CASE WHEN (sendonbehalf.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE sendonbehalf.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE sendonbehalf.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (sendonbehalf.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE sendonbehalf.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE sendonbehalf.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE sendonbehalf.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.sendonbehalf WHERE user_ID <> 0 and identity_ID <> 0";
            }

            DataTable dtFullAccess = dbConn.retrieveDB(SQL);

            foreach (DataRow row in dtFullAccess.Rows)
            {
                var rowAuth = dataTable.NewRow();
                rowAuth["UserAccount"] = row["UserAccount"];
                rowAuth["Mailbox"] = row["Mailbox"];
                rowAuth["AccessRights"] = row["AccessRights"];
                rowAuth["InheritanceType"] = row["InheritanceType"];
                dataTable.ImportRow(row);
            }
        }

        private void loadForwarding()
        {
            string SQL = "";

            if (username != "theodoros.h" && username != "arisadmin"  && username != "vrionis.n")
            {
                SQL = "SELECT CASE WHEN (forwarding.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE forwarding.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE forwarding.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (forwarding.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE forwarding.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE forwarding.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE forwarding.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.forwarding WHERE  user_ID NOT IN (0, 242, 109, 52, 32, 256, 142, 28, 106) and identity_ID <> 0";
            }
            else
            {
                SQL = "SELECT CASE WHEN (forwarding.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE forwarding.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE forwarding.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (forwarding.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE forwarding.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE forwarding.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE forwarding.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.forwarding WHERE user_ID <> 0 and identity_ID <> 0";
            }

            DataTable dtFullAccess = dbConn.retrieveDB(SQL);

            foreach (DataRow row in dtFullAccess.Rows)
            {
                var rowAuth = dataTable.NewRow();
                rowAuth["UserAccount"] = row["UserAccount"];
                rowAuth["Mailbox"] = row["Mailbox"];
                rowAuth["AccessRights"] = row["AccessRights"];
                rowAuth["InheritanceType"] = row["InheritanceType"];
                dataTable.ImportRow(row);
            }
        }

        private void loadGroupMembers()
        {
            string SQL = "";

            if (username != "theodoros.h" && username != "arisadmin" && username != "vrionis.n")
            {
                SQL = "SELECT CASE WHEN (groupmembers.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE groupmembers.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE groupmembers.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (groupmembers.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE groupmembers.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE groupmembers.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE groupmembers.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.groupmembers WHERE user_ID NOT IN (0, 242, 109, 52, 32, 256, 142, 28, 106) and identity_ID <> 0";
            }
            else
            {
                SQL = "SELECT CASE WHEN (groupmembers.userType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE groupmembers.user_ID = users.userID and userName is not null) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE groupmembers.user_ID = groups.groupID and groupName is not null) END as UserAccount, CASE WHEN (groupmembers.identityType = 'User') THEN (SELECT userEmail FROM mailboxes.dbo.users WHERE groupmembers.identity_ID = users.userID) ELSE (SELECT groupEmail FROM mailboxes.dbo.groups WHERE groupmembers.identity_ID = groups.groupID) END as Mailbox, (SELECT accessRightsName FROM mailboxes.dbo.accessrights WHERE groupmembers.accessRights_ID = accessrights.accessRightsID) as AccessRights, inheritanceType FROM mailboxes.dbo.groupmembers WHERE user_ID <> 0 and identity_ID <> 0";
            }

            DataTable dtFullAccess = dbConn.retrieveDB(SQL);

            foreach (DataRow row in dtFullAccess.Rows)
            {
                var rowAuth = dataTable.NewRow();
                rowAuth["UserAccount"] = row["UserAccount"];
                rowAuth["Mailbox"] = row["Mailbox"];
                rowAuth["AccessRights"] = row["AccessRights"];
                rowAuth["InheritanceType"] = row["InheritanceType"];
                dataTable.ImportRow(row);
            }
        }

        private void gotGet()
        {
            foreach (DataRow rowForwarding in dataTable.Rows)
            {
                if (rowForwarding["AccessRights"].ToString() == "Forwarding")
                {
                    string dlGroup = rowForwarding["Mailbox"].ToString();
                    string userAcc = rowForwarding["UserAccount"].ToString();

                    foreach (DataRow rowGroupMembers in dataTable.Rows)
                    {
                        if (rowGroupMembers["AccessRights"].ToString() == "GroupMember" && rowGroupMembers["Mailbox"].ToString() == dlGroup)
                        {
                            var rowAuth = dataTableTmp.NewRow();
                            rowAuth["UserAccount"] = rowGroupMembers["UserAccount"];
                            rowAuth["Mailbox"] = userAcc + " To " + dlGroup;
                            rowAuth["AccessRights"] = "MailboxForwardingToGroup";
                            rowAuth["InheritanceType"] = "";
                            dataTableTmp.Rows.Add(rowAuth);
                        }
                    }
                }
            }

            dataTable.Merge(dataTableTmp);
        }

        // ***********************
        //          END
        // ***********************
    }
}
