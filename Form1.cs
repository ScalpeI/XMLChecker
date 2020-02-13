using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace XMLChecker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            PopulateTreeView();
            treeView1.ExpandAll();
            lblCheck.Text = "";
            lblSort.Text = "";
            if (listView1.Items.Count == 0)
            {
                button1.Enabled = false;
                button2.Enabled = false;
            }
            else
            {
                button1.Enabled = true;
                button2.Enabled = true;
            }
        }

        private void PopulateTreeView()
        {
            TreeNode rootNode;

            DirectoryInfo info = new DirectoryInfo(Directory.GetCurrentDirectory());
            if (info.Exists)
            {
                rootNode = new TreeNode(info.Name);
                rootNode.Tag = info;
                GetDirectories(info.GetDirectories(), rootNode);
                treeView1.Nodes.Add(rootNode);
            }
        }

        private void GetDirectories(DirectoryInfo[] subDirs,   TreeNode nodeToAddTo)
        {
            TreeNode aNode;
            DirectoryInfo[] subSubDirs;
            foreach (DirectoryInfo subDir in subDirs)
            {
                aNode = new TreeNode(subDir.Name, 0, 1);
                aNode.Tag = subDir;
                aNode.ImageKey = "folder";
                subSubDirs = subDir.GetDirectories();
                if (subSubDirs.Length != 0)
                {
                    GetDirectories(subSubDirs, aNode);
                }
                nodeToAddTo.Nodes.Add(aNode);
            }
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeNode newSelected = e.Node;
            listView1.Items.Clear();
            DirectoryInfo nodeDirInfo = (DirectoryInfo)newSelected.Tag;
            ListViewItem.ListViewSubItem[] subItems;
            ListViewItem item = null;

            foreach (FileInfo file in nodeDirInfo.GetFiles("*.xml"))
            {
                item = new ListViewItem(file.Name, 2);
                subItems = new ListViewItem.ListViewSubItem[]
                          { new ListViewItem.ListViewSubItem(item, file.Length.ToString() + " байт"),
                   new ListViewItem.ListViewSubItem(item, file.LastAccessTime.ToShortDateString()),
                          new ListViewItem.ListViewSubItem(item, file.FullName)};
                
                item.SubItems.AddRange(subItems);
                listView1.Items.Add(item);
            }
            if (listView1.Items.Count == 0)
            {
                button1.Enabled = false;
                button2.Enabled = false;
            }
            else
            {
                button1.Enabled = true;
                button2.Enabled = true;
            }
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = false;
            listView2.Items.Clear();
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            for (int i = 0; i< listView1.Items.Count; i++)
            {
                xmlCheck(listView1.Items[i].SubItems[3].Text);
                listView2.Items.Add(listView1.Items[i].Text,2);
                lblCheck.Text = string.Format("Всего обработанно файлов: {0}; Последний обработанный файл: {1};",i+1, listView1.Items[i].Text);
            }
            button1.Enabled = true;
            button2.Enabled = true;
        }

        private void xmlCheck(string filepath)
        {
            try
            {
                string s1, s2;
                string connString = @"Server=192.168.1.101;Database=Oms_Buryatiya;User ID=ss;Password=qwerty;";
                SqlConnection con = new SqlConnection(connString);
                XDocument xml = XDocument.Load(filepath);
                IEnumerable<XElement> xElms;
                xml.Root.Attribute("NRECORDS").Value = xml.Root.Elements().Count().ToString();
                //Исправление
                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П040");
                foreach (XElement xElm in xElms)
                {
                    xElm.Element("VIZIT").Element("FPOLIS").Value = "0";
                    xElm.Element("INSURANCE").Element("POLIS").Element("DBEG").Value = xElm.Element("VIZIT").Element("DVIZIT").Value;
                }//Конец
                 //Исправление окончание действия полиса и удаление OLD_PERSON и OLDDOC_LIST на П060
                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П060");
                foreach (XElement xElm in xElms)
                {
                    //xElm.Element("INSURANCE").Element("POLIS").Element("DBEG").Value = xElm.Element("VIZIT").Element("DVIZIT").Value;
                    if (xElm.Element("OLDDOC_LIST") != null)
                        xElm.Element("OLDDOC_LIST").Remove();
                    if (xElm.Element("OLD_PERSON") != null)
                        xElm.Element("OLD_PERSON").Remove();
                    if (xElm.Element("PERSON").Element("C_OKSM").Value != "RUS")
                    {
                        if (xElm.Element("INSURANCE").Element("POLIS").Element("DEND") == null)
                            xElm.Element("INSURANCE").Element("POLIS").Add(new XElement("DEND", xElm.Element("DOC_LIST").Elements("DOC").Last().Element("DOCEXP").Value));
                        if (int.Parse(xElm.Element("INSURANCE").Element("POLIS").Element("DEND").Value.Substring(0, 4)) >
                                int.Parse(xElm.Element("INSURANCE").Element("POLIS").Element("DBEG").Value.Substring(0, 4)))
                            if (DateTime.Parse(xElm.Element("INSURANCE").Element("POLIS").Element("DBEG").Value) >
                                DateTime.Parse(xElm.Element("INSURANCE").Element("POLIS").Element("DBEG").Value.Substring(0, 4) + "-11-20"))
                            {
                                if (int.Parse(xElm.Element("INSURANCE").Element("POLIS").Element("DEND").Value.Substring(0, 4)) >
                                      (int.Parse(xElm.Element("INSURANCE").Element("POLIS").Element("DBEG").Value.Substring(0, 4)) + 1))
                                    xElm.Element("INSURANCE").Element("POLIS").Element("DEND").Value =
                                        (int.Parse(xElm.Element("INSURANCE").Element("POLIS").Element("DBEG").Value.Substring(0, 4)) + 1).ToString() + "-12-31";
                            }
                            else
                                xElm.Element("INSURANCE").Element("POLIS").Element("DEND").Value =
                                            xElm.Element("INSURANCE").Element("POLIS").Element("DBEG").Value.Substring(0, 4) + "-12-31";
                    }
                }//конец
                 //Очистка OLD_PERSON по П031
                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П031");
                foreach (XElement xElm in xElms)
                {
                    if (xElm.Element("OLD_PERSON") != null)
                        xElm.Element("OLD_PERSON").Remove();
                }//конец
                 //Корректировка OLDDOC
                xElms = xml.Descendants("OP").
                    Where(x => x.Element("OLDDOC_LIST") != null);
                foreach (XElement xElm in xElms)
                {

                    if (xElm.Element("OLDDOC_LIST") != null)
                        if (xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCDATE") != null)
                            xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCDATE").Remove();
                    if (xElm.Element("OLDDOC_LIST") != null)
                        if (xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("NAME_VP") != null)
                            xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("NAME_VP").Remove();
                    if (xElm.Element("OLDDOC_LIST") != null)
                        if (xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCSER") == null && xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCNUM") != null)
                            if (xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCNUM").Value == "")
                                xElm.Element("OLDDOC_LIST").Remove();
                    if (xElm.Element("OLDDOC_LIST") != null)
                        if (xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCNUM") == null && xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCSER") != null)
                            if (xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCSER").Value == "")
                                xElm.Element("OLDDOC_LIST").Remove();
                    if (xElm.Element("OLDDOC_LIST") != null)
                        if (xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCNUM") != null && xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCSER") != null)
                            if (xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCSER").Value == "" & xElm.Element("OLDDOC_LIST").Element("OLD_DOC").Element("DOCNUM").Value == "")
                                xElm.Element("OLDDOC_LIST").Remove();

                }//конец
                 //Очистка OLD_PERSON по П032
                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П032");
                foreach (XElement xElm in xElms)
                {
                    if (xElm.Element("OLD_PERSON") != null)
                        xElm.Element("OLD_PERSON").Remove(); 
                }//конец
                 //Проверка пола
                xElms = xml.Descendants("OP").
                    Where(x => x.Element("PERSON") != null);
                foreach (XElement xElm in xElms)
                {
                    if (xElm.Element("PERSON").Element("FAM") != null)
                        if (xElm.Element("PERSON").Element("IM") != null)
                            if (xElm.Element("PERSON").Element("OT") != null)
                                if (xElm.Element("PERSON").Element("W") != null)
                                {
                                    s1 = xElm.Element("PERSON").Element("OT").Value.Substring(xElm.Element("PERSON").Element("OT").Value.Length - 1, 1);
                                    s2 = xElm.Element("PERSON").Element("W").Value;
                                    if ((s1 == "а") || (s1 == "А")) if (s2 == "1") listBox1.Items.Add(xml.Root.Attribute("FILENAME").Value + " " +
                                        xElm.Element("PERSON").Element("FAM").Value + " " + xElm.Element("PERSON").Element("IM").Value + " " + 
                                        xElm.Element("PERSON").Element("OT").Value + " Мужской");
                                    if ((s1 == "ч") || (s1 == "Ч")) if (s2 == "2") listBox1.Items.Add(xml.Root.Attribute("FILENAME").Value + " " +
                                        xElm.Element("PERSON").Element("FAM").Value + " " + xElm.Element("PERSON").Element("IM").Value + " " +
                                        xElm.Element("PERSON").Element("OT").Value + " Женский");
                                }
                }//конец
                 //Сверка с БД имен и отчеств
                xElms = xml.Descendants("OP").
                    Where(x => x.Element("PERSON") != null);
                foreach (XElement xElm in xElms)
                {
                    if (xElm.Element("PERSON").Element("IM") != null)
                        if (xElm.Element("PERSON").Element("W") != null)
                        {
                            string fam, im, ot, w;
                            string query = @"SELECT count([Name]) FROM[Oms_Buryatiya].[dbo].[_sIMA] WHERE Name='" + xElm.Element("PERSON").Element("IM").Value + "'";
                            if (xElm.Element("PERSON").Element("FAM") != null) fam = xElm.Element("PERSON").Element("FAM").Value; else fam = "";
                            if (xElm.Element("PERSON").Element("IM") != null) im = xElm.Element("PERSON").Element("IM").Value; else im = "";
                            if (xElm.Element("PERSON").Element("OT") != null) ot = xElm.Element("PERSON").Element("OT").Value; else ot = "";
                            SqlCommand comm = new SqlCommand(query, con);
                            con.Open();
                            if (int.Parse(Convert.ToString(comm.ExecuteScalar())) < 1)
                                listBox2.Items.Add(xml.Root.Attribute("FILENAME").Value + " " +
                                        fam + " " +
                                        im + " " +
                                        ot + " " +
                                        " Имени нет в БД");
                            con.Close();
                        }
                    if (xElm.Element("PERSON").Element("OT") != null)
                        if (xElm.Element("PERSON").Element("W") != null)
                        {
                            string fam, im, ot, w;
                            string query = @"SELECT count([Name]) FROM[Oms_Buryatiya].[dbo].[_sOtch] WHERE Name='" + xElm.Element("PERSON").Element("OT").Value + "'";
                            if (xElm.Element("PERSON").Element("FAM") != null) fam = xElm.Element("PERSON").Element("FAM").Value; else fam = "";
                            if (xElm.Element("PERSON").Element("IM") != null) im = xElm.Element("PERSON").Element("IM").Value; else im = "";
                            if (xElm.Element("PERSON").Element("OT") != null) ot = xElm.Element("PERSON").Element("OT").Value; else ot = "";
                            SqlCommand comm = new SqlCommand(query, con);
                            con.Open();
                            if (int.Parse(Convert.ToString(comm.ExecuteScalar())) < 1)
                                listBox2.Items.Add(xml.Root.Attribute("FILENAME").Value + " " +
                                        fam + " " +
                                        im + " " +
                                        ot + " " +
                                        " Отчества нет в БД");
                            con.Close();
                        }
                }//конец
                xml.Save(filepath);
            }
            catch (XmlException xmlex)
            {
                MessageBox.Show(Path.GetFileName(xmlex.SourceUri) + Environment.NewLine + xmlex.Message,"Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listView2.Items.Clear();
            button1.Enabled = false;
            button2.Enabled = false;
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                xmlSort(listView1.Items[i].SubItems[3].Text);
                listView2.Items.Add(listView1.Items[i].Text, 2);
                lblSort.Text = string.Format("Всего обработанно файлов: {0}; Последний обработанный файл: {1};", i + 1, listView2.Items[i].Text);
            }
            button1.Enabled = true;
            button2.Enabled = true;
        }

        private void xmlSort(string filepath)
        {
            try
            {
                XDocument xml = XDocument.Load(filepath);
                listBox1.Items.Clear();
                xml.Root.Attribute("NRECORDS").Value = xml.Root.Elements().Count().ToString();
                textBox1.Text = xml.Declaration.ToString() + Environment.NewLine;
                textBox1.Text += "<OPLIST VERS = \"" + xml.Root.Attribute("VERS").Value + "\" FILENAME = \"" + xml.Root.Attribute("FILENAME").Value + "\" SMOCOD = \"" +
                    xml.Root.Attribute("SMOCOD").Value + "\" PRZCOD = \"" + xml.Root.Attribute("PRZCOD").Value + "\" NRECORDS = \"" + xml.Root.Attribute("NRECORDS").Value + "\">";
                IEnumerable<XElement> xElms;
                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П021");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П022");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П023");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П040");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П061");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П034");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П035");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П010");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П031");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П032");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П033");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П062");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П063");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value != "П021" && x.Element("TIP_OP").Value != "П022" &&
                    x.Element("TIP_OP").Value != "П023" && x.Element("TIP_OP").Value != "П040" &&
                    x.Element("TIP_OP").Value != "П061" && x.Element("TIP_OP").Value != "П034" &&
                    x.Element("TIP_OP").Value != "П035" && x.Element("TIP_OP").Value != "П010" &&
                    x.Element("TIP_OP").Value != "П031" && x.Element("TIP_OP").Value != "П032" &&
                    x.Element("TIP_OP").Value != "П033" && x.Element("TIP_OP").Value != "П062" &&
                    x.Element("TIP_OP").Value != "П063" && x.Element("TIP_OP").Value != "П060");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }

                xElms = xml.Descendants("OP").
                    Where(x => x.Element("TIP_OP").Value == "П060");
                foreach (XElement xElm in xElms)
                {
                    textBox1.Text += xElm.ToString().ToUpper() + Environment.NewLine;
                }
                textBox1.Text += "</OPLIST>";
                XDocument xml2 = XDocument.Parse(textBox1.Text);
                xml2.Save(filepath);
            }
            catch (XmlException xmlex)
            {
                MessageBox.Show(Path.GetFileName(xmlex.SourceUri) + Environment.NewLine + xmlex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
           AboutBox1 AboutBox = new AboutBox1();
            AboutBox.ShowDialog();
        }
    }
}
