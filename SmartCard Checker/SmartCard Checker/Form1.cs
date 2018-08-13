using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading;

namespace WindowsFormsApp5
{
    public partial class Form1 : Form
    {

        static Thread t = new Thread(new ThreadStart(SplashStart));
        private DataTable dt = new DataTable();
        private DataTable antigo = new DataTable();
        private DataTable atual = new DataTable();

        private static PrincipalContext context = new PrincipalContext(ContextType.Domain, "gbl.ad.hedani.net", "OU=BR,OU=CS,DC=gbl,DC=ad,DC=hedani,DC=net");
        private static PrincipalSearcher searcher = new PrincipalSearcher(new UserPrincipal(context));

        private static PrincipalContext context2 = new PrincipalContext(ContextType.Domain, "gbl.ad.hedani.net", "OU=FunctionalGroups,DC=gbl,DC=ad,DC=hedani,DC=net");

        private static PrincipalContext context3 = new PrincipalContext(ContextType.Domain, "gbl.ad.hedani.net", "OU=PersonalGroups,DC=gbl,DC=ad,DC=hedani,DC=net");
        private static String hoje = DateTime.Now.ToString("dd.MM.yyyy");
        private static String hora = DateTime.Now.ToString();
        private static String path = @"\\csao11p20011d\rwapps\CSHG\CSHGDsl\SUPORTE\SmartCardChecker\" + hoje+" SmartCardLog.csv";
        


        GroupPrincipal group = GroupPrincipal.FindByIdentity(context2, "DSO-ALL-00-pwd_allowed_edge_app");
        GroupPrincipal group1 = GroupPrincipal.FindByIdentity(context2, "DSO-ALL-00-pwd_allowed_businesslaptop");
        GroupPrincipal group2 = GroupPrincipal.FindByIdentity(context2, "DSO-ALL-00-pwd_allowed_multi_machine_TF");
        GroupPrincipal group3 = GroupPrincipal.FindByIdentity(context3, "DSO-ALL-00-pwd_allowed_mypartner_user");
        GroupPrincipal group4 = GroupPrincipal.FindByIdentity(context2, "DSO-ALL-00-pwd_allowed_non_std_acc_N_PID");
        GroupPrincipal group5 = GroupPrincipal.FindByIdentity(context2, "DSO-ALL-00-pwd_allowed_multi_machine_nonTF");
        GroupPrincipal group6 = GroupPrincipal.FindByIdentity(context2, "DSO-ALL-00-pwd_allowed_oth_reasons");



        private void geragrid(ref DataTable dt, ref PrincipalContext context, ref PrincipalSearcher searcher)
        {


            dt.Columns.Add("Nome", typeof(string));
            dt.Columns.Add("Login", typeof(string));
            dt.Columns.Add("PID", typeof(string));
            dt.Columns.Add("Office", typeof(string));
            dt.Columns.Add("SmartCard", typeof(string));
            dt.Columns.Add("SmartCard2", typeof(bool));
            dt.Columns.Add("Account", typeof(string));
            dt.Columns.Add("Exception", typeof(string));
            dt.Columns.Add("Groups", typeof(string));
            dt.Columns.Add("Altered", typeof(string));

            int i = 0;
            foreach (var result in searcher.FindAll())
            {
             
                    DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                    UserPrincipal user = (UserPrincipal)result;
                    DataRow dr = dt.NewRow();
                    dr["Nome"] = user.DisplayName;
                    dr["Login"] = user.SamAccountName;
                    dr["PID"] = user.Name;
                    dr["Office"] = de.Properties["physicalDeliveryOfficeName"].Value;


                if (user.SmartcardLogonRequired == true)
                {
                    dr["SmartCard"] = "Enabled";
                }

                if (user.SmartcardLogonRequired == false)
                {
                    dr["SmartCard"] = "Disabled";
                }


                dr["SmartCard2"] = user.SmartcardLogonRequired.ToString();

                    if(user.Enabled == true)
                    {
                        dr["Account"] = "Enabled";
                    }

                    if (user.Enabled == false)
                    {
                        dr["Account"] = "Disabled";
                    }
                    

                    
                    bool Exception = false;
                    string groupcell = "";

                    if (user.IsMemberOf(group))
                    { 
                        Exception = true;
                        groupcell += group.Name + "|";
                    }
                    if (user.IsMemberOf(group1))
                    {
                         Exception = true;
                         groupcell += group1.Name + "|";
                    }
                    if (user.IsMemberOf(group2))
                    {
                         Exception = true;
                         groupcell += group2.Name + "|";
                     }
                    if (user.IsMemberOf(group3))
                    {
                         Exception = true;
                         groupcell += group3.Name + "|";
                    }
                    if (user.IsMemberOf(group4))
                    {
                         Exception = true;
                         groupcell += group4.Name + "|";
                     }
                    if (user.IsMemberOf(group5))
                    {
                         Exception = true;
                         groupcell += group5.Name + "|";
                    }
                    if (user.IsMemberOf(group6))
                    {
                         Exception = true;
                         groupcell += group6.Name + "|";
                    }
                groupcell.TrimEnd('|');


                dr["Exception"] = Exception;
                dr["groups"] = groupcell;

                    dt.Rows.Add(dr);
                    i++;
                
                    
            }
            
            dataGridView1.DataSource = dt;

            dataGridView1.Columns["SmartCard2"].Visible = false;
            dataGridView1.Columns[0].Width = 150;
            dataGridView1.Columns["Login"].Width = 60;
            dataGridView1.Columns["SmartCard"].Width = 60;
            dataGridView1.Columns["SmartCard2"].Width = 60;
            dataGridView1.Columns["Account"].Width = 60;
            dataGridView1.Columns["Exception"].Width = 60;
            dataGridView1.Columns["Altered"].Width = 50;

            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView1.Sort(dataGridView1.Columns["Login"], System.ComponentModel.ListSortDirection.Ascending);




        }

        private void exportExcel()
        {
            if (dt.Columns.Count != 0)
            {
                copyAlltoClipboard();
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Microsoft.Office.Interop.Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            if (dt.Columns.Count == 0)
            {
                MessageBox.Show("Gere a planilha antes!");
            }
        }

        private void copyAlltoClipboard()
        {
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView1.MultiSelect = true;
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null) Clipboard.SetDataObject(dataObj);
        }
        private void trocaCard(DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["SmartCard2"].Index)
            {
                var ci = e.ColumnIndex;
                var ri = e.RowIndex;
                bool cell = (bool)dataGridView1.Rows[ri].Cells[ci].Value;
                bool valor = true;
                string palavra = "Disabled";
                if (cell == true)
                {
                    UserPrincipal p = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, dataGridView1.Rows[ri].Cells[1].Value.ToString());
                    p.SmartcardLogonRequired = false;
                    p.Save();
                    valor = false;
                    palavra = "Disabled";

                }
                if (cell == false)
                {
                    UserPrincipal p = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, dataGridView1.Rows[ri].Cells[1].Value.ToString());
                    p.SmartcardLogonRequired = true;
                    p.Save();
                    valor = true;
                    palavra = "Enabled";
                }
                dataGridView1.Rows[ri].Cells[ci].Value = valor;
                dataGridView1.Rows[ri].Cells[ci - 1].Value = palavra;
                dataGridView1.Rows[ri].Cells[ci+4].Value = "X";
            }
        }
        private void mudaColuna()
        {
            if (checkBox2.Checked == true)
            {
                MessageBox.Show("ATENÇÃO ALTERAÇÕES HABILITADAS");
                dataGridView1.Columns["SmartCard2"].Visible = true;
                dataGridView1.Columns["SmartCard"].Visible = false;
                dataGridView1.Columns["SmartCard2"].HeaderText = "SmartCard";
            }
            if (checkBox2.Checked == false)
            {
                dataGridView1.Columns["SmartCard2"].Visible = false;
                dataGridView1.Columns["SmartCard"].Visible = true;
                dataGridView1.Columns["SmartCard2"].HeaderText = "SmartCard";
            }
        }
        public Form1()
        {
            
            t.Start();
            InitializeComponent();
        }

       static public void SplashStart()
        {
            Application.Run(new WaitForm());
        }
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Em construção");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            
            geragrid(ref dt, ref context, ref searcher);
            dataGridView1.Sort(dataGridView1.Columns["Login"], System.ComponentModel.ListSortDirection.Ascending);
        }

        private void geraLog(String s)
        {
            if (!File.Exists(path))
            {
                using (var tw = new StreamWriter(path, true))
                {
                    tw.WriteLine("Executer;Affected User;Before;After;Time");
                    tw.Close();
                }
            }
            using (var tw = new StreamWriter(path, true))
            {
                tw.WriteLine(Environment.UserName + ";" +s+";"+hora+";");
                tw.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                exportExcel();
            }
            catch (Exception t)
            {
                MessageBox.Show("Gere a planilha antes!");
            }
        }


        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            
                mudaColuna();
                dataGridView1.Sort(dataGridView1.Columns["Login"], System.ComponentModel.ListSortDirection.Ascending);

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                
                if (e.ColumnIndex == dataGridView1.Columns["SmartCard2"].Index)
                {
                    int ci = e.ColumnIndex;
                    int ri = e.RowIndex;
                    bool valor = (bool)dataGridView1.Rows[ri].Cells[ci].Value;
                    String nome = (String)dataGridView1.Rows[ri].Cells[ci - 4].Value;

                    trocaCard(e);
                    if(valor == true)
                    {
                        geraLog(nome+ ";TRUE;FALSE");
                    }
                    if (valor == true)
                    {
                        geraLog(nome + ";FALSE;TRUE");
                    }
                }
                    
            }
            catch (Exception t)
            {
                
                int ci = e.ColumnIndex;
                int ri = e.RowIndex;
                String nome = (String)dataGridView1.Rows[ri].Cells[ci - 3].Value;
                MessageBox.Show(t.Message);
            }
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            geragrid(ref dt, ref context, ref searcher);
            label1.Text = Environment.UserName.ToUpper();
            dataGridView1.Columns["Account"].HeaderText = "NT Account Status";
            richTextBox1.KeyPress += new KeyPressEventHandler (tb_KeyDown);
            t.Abort();
            
        }

        private void tb_KeyDown(object sender, KeyPressEventArgs e)
        {
            if (Char.IsControl(e.KeyChar))
            {
                if (richTextBox1.Text == "")
                {
                    dataGridView1.DataSource = dt;
                    dataGridView1.Sort(dataGridView1.Columns["Login"], System.ComponentModel.ListSortDirection.Ascending);
                    checkBox3.Visible = true;
                }
                if (richTextBox1.Text != "")
                {
                    checkBox3.Visible = false;
                    try
                    {
                        DataTable query = dt.AsEnumerable().Where(m => richTextBox1.Lines.Contains(m.ItemArray[1])).CopyToDataTable();
                        dataGridView1.DataSource = query;
                        dataGridView1.Sort(dataGridView1.Columns["Login"], System.ComponentModel.ListSortDirection.Ascending);

                    }
                    catch
                    {

                    }
                }
                
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            
            if (checkBox3.Checked == true)
            {
                aplicaFiltro();
            }
           
            if (checkBox3.Checked == false)
            {
                dataGridView1.DataSource = dt;
                dataGridView1.Sort(dataGridView1.Columns["Login"], System.ComponentModel.ListSortDirection.Ascending);

            }
        }

        private void aplicaFiltro()
        {
            
            var myRegex = new Regex(@"N[0-9]*");
            var myRegex1 = new Regex(@"M[0-9]*");
            var myRegex2 = new Regex(@"Z[0-9]*");
            var myRegex3 = new Regex(@"K[0-9]*");
            var myRegex4 = new Regex(@"Y[0-9]*");
            var myRegex5 = new Regex(@"G[0-9]*");

             atual = dt.AsEnumerable()
                      .Where(x => x.Field<string>("SmartCard") == "Disabled" && x.Field<string>("Exception") == "False" && x.Field<string>("Account") == "Enabled" && x.Field<string>("Office") != "")
                      .Where(x => myRegex.IsMatch(x["Login"].ToString()) == false)
                      .Where(x => myRegex1.IsMatch(x["Login"].ToString()) == false)
                      .Where(x => myRegex2.IsMatch(x["Login"].ToString()) == false)
                      .Where(x => myRegex3.IsMatch(x["Login"].ToString()) == false)
                      .Where(x => myRegex4.IsMatch(x["Login"].ToString()) == false)
                      .Where(x => myRegex5.IsMatch(x["Login"].ToString()) == false)
                      .CopyToDataTable();
                      
            DataView results = atual.DefaultView;
            results.RowFilter = "Login NOT LIKE 'sv_*'";

            dataGridView1.DataSource = results;
            dataGridView1.Sort(dataGridView1.Columns["Login"], System.ComponentModel.ListSortDirection.Ascending);

        }


    }
}
