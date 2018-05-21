using OPCAutomation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        Form1 fr1 = new Form1();
        public Form2(Form1 f1)
        {
            InitializeComponent();
            fr1 = f1;
        }
        OPCBrowser browser = ServerConnect.KepServer.CreateBrowser();
        TreeNode CountNode;
        int itmHandlerClient = 0;
        int itmHandlerServer = 0;
        bool a=true;
        OPCServer KepServer;
        OPCGroups kepGroups;
        OPCGroup KepGroup;
        OPCItem KepItem;
        OPCItems KepItems;
       

        #region 方法
        /// <summary>
        /// 列出opc服务器中所有节点
        /// </summary>
        /// <param name="TransactionID"></param>
        /// <param name="NumItems"></param>
        /// <param name="ClientHandles"></param>
        /// <param name="Errors"></param>
        private void RecurBrowse(OPCBrowser oPCBrowser)
        {
            //写入接到treeview
           
            LoadDataToTree(oPCBrowser, treeView1.Nodes[0].Nodes);
            oPCBrowser.ShowBranches();
            oPCBrowser.ShowLeafs(true);
            foreach (object turn in oPCBrowser)
            {
               listBox1.Items.Add(turn.ToString());
            }
        }

        /// <summary>
        /// Treeview 遍历显示节点
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        //private void treeViewsShow(bool blsRoot,string parent)
        //{
        //    browser.ShowBranches();
        //    int total = browser.Count;
        //    String[] trName = { };
        //    TreeNode nodex;
        //    if (total > 0)
        //    {
        //        for (int i = 1; i <=total; i++)
        //        {
        //            trName[i] = browser.Item(i);
                    
        //            nodex = treeView1.Nodes.Add(trName[i], trName[i]);
               
        //            browser.MoveDown(trName[i]);
        //            treeViewsShow(false, trName[i]);
        //            browser.MoveUp();
        //        }
        //    }
        //    else
        //    {
        //        return;
        //    }
        //    browser.MoveToRoot();
        //    treeViewsShow(true, "");
        //}

        /// <summary>
        /// 将treeview选中的item中值，显示到listbox中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        //private void listBoxShow()
        //{
        //    string[] strBranch = { };
        //    browser.ShowBranches();
        //    browser.ShowLeafs(true);
        //    for (int j = 1; j <= browser.Count; j++)
        //    {
        //        listView1.Items.Add(browser.Item(j));
        //    }

        //}
        /// <summary>
        /// 将数据获取到treeview
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoadDataToTree(OPCBrowser oPCBrowser, TreeNodeCollection treeNodeCollection)
        {
            oPCBrowser.Organization.ToString();
            oPCBrowser.ShowBranches();
            foreach (object turn in oPCBrowser)
            {
                TreeNode node = treeNodeCollection.Add(turn.ToString());
                treeView1.SelectedNode = node;
                oPCBrowser.MoveDown(turn.ToString());
                LoadDataToTree(oPCBrowser, node.Nodes);
                oPCBrowser.MoveUp();
            }
            oPCBrowser.ShowLeafs(false);
            foreach (object turn in oPCBrowser)
            {
                treeView1.SelectedNode.Nodes.Add(turn.ToString());

            }
        }
    /// <summary>
    /// Form1 ListView显示相应的值
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    /// 
        private void getListView()
        {
            try
            {
                itmHandlerClient = 1234;
                KepItem =KepItems.AddItem(textBox2.Text.ToString(), itmHandlerClient);
                itmHandlerServer = KepItem.ServerHandle;
                ListViewItem lv = new ListViewItem(textBox2.Text.ToString());
                lv.SubItems.Add(KepItem.Value.toString());
                lv.SubItems.Add(KepItem.TimeStamp.ToShortTimeString());
                fr1.listView1.Items.Add(lv);
            }
            catch (Exception err)
            {
                //没有任何权限的项，都是opc服务器保留的系统项，此处不做任何处理。
                itmHandlerClient = 0;
                MessageBox.Show("此项为保留项：" + err.Message, "提示信息");
            }
        }
        




        /// <summary>
        /// 每当数据有变化时，执行的事件
        /// </summary>
        /// <param name="TransactionID">处理ID</param>
        /// <param name="NumItems">项个数</param>
        /// <param name="ClientHandles">项客户端句柄</param>
        /// <param name="ItemValues">Tag值</param>
        /// <param name="Qualities">品质</param>
        /// <param name="TimeStamps">时间戳</param>
        void KepGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            for (int i = 1; i <= NumItems; i++)
            {
             //  ListViewItem lv = new ListViewItem(ItemValues.GetValue(i).ToString()); 
             //  lv.SubItems.Add(Qualities.GetValue(i).ToString());
              // lv.SubItems.Add(TimeStamps.GetValue(i).ToString());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        ///

        /// <summary>
        /// 创建组
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CreateGroup()
        {
            try
            {
            
                //KepServer  = new OPCServer();
                kepGroups = KepServer.OPCGroups;
                KepGroup = kepGroups.Add("opcdontnetgroup1");
                SetGroupProperty();
                KepGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(KepGroup_DataChange);
                // ServerConnect.KepGroup.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(KepGroup_AsyncWriteComplete);
                KepItems = KepGroup.OPCItems;
                
            }
            catch (Exception err)
            {
                MessageBox.Show("创建组出错：" + err.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);


            }
          
        }

        private void SetGroupProperty()
        {
            //ServerConnect.KepServer.OPCGroups.DefaultGroupIsActive =true;
            //ServerConnect.KepServer.OPCGroups.DefaultGroupDeadband =0;
            //ServerConnect.KepGroup.IsActive = true;
            //ServerConnect.KepGroup.IsSubscribed = true;
            //ServerConnect.KepGroup.UpdateRate = 250;
        }
     #endregion
        #region 事件触发
        private void Form2_Load(object sender, EventArgs e)
        {
           
            CountNode = new TreeNode("Root");
            treeView1.Nodes.Add(CountNode);
            RecurBrowse(browser);
           
        }

        //private void TreeViewShow(TreeNode CountNode)
        //{
        //    browser.ShowLeafs(true);
            
            
        //    if (CountNode.Nodes.Count == 0)
        //    {
        //        if (CountNode.Parent == null)
        //        {
        //            listView1.Items.Add("root");
        //        }
        //        else
        //        {
        //            foreach (object turn in browser)
        //            {
        //                CountNode.Nodes.Add(turn.ToString());
        //            }
        //        }
        //    }
        //    else
        //    {

        //    }
        //}

        //private void ListViewShow(TreeNode CountNode)
        //{
        //    listView1.Clear();
        //    if (CountNode.Parent != null)
        //    {

        //    }
        //}


        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            
        }

        private void treeView1_Click(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 将listbox1中选中的项，反馈到textbox中去。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = listBox1.SelectedItem.ToString();
        }
        /// <summary>
        /// 将textbox中添加的item值反馈到form1中去。测试可读性能。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //ListViewItem lv = new ListViewItem("x");
            //lv.SubItems.Add("x");
            //lv.SubItems.Add("x");
            //fr1.listView1.Items.Add(lv);
            if (a) 
            {
                CreateGroup();
                a = false;
            }
            getListView();
          //  ServerConnect.KepGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(KepGroup_DataChange);

           
        }
    }

}
