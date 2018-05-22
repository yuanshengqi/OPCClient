using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OPCAutomation;

namespace WindowsFormsApplication1
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        #region 定义变量
        ListViewItem lv;
        ServerConnect sc1 = new ServerConnect();
        OPCServer server = ServerConnect.KepServer;

        OPCGroups kepGroups;
        OPCGroup KepGroup;
        OPCItem KepItem;
        OPCItems KepItems;
        List<String> list;
        bool connect = true;
        int itmHandlerClient=1234;
        int itmHandlerServer = 0;
        object value;
        object quality;
        object timestamp;
        private List<string> itemNames = new List<string>();
        
        private Dictionary<string, string> itemValues = new Dictionary<string, string>();
        static  List<string> list1=null;
        #endregion

        /// <summary>
        /// 创建组
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private bool CreateGroup()
        {

            try
            {
                kepGroups = ServerConnect.KepServer.OPCGroups;
                KepGroup = kepGroups.Add("opcdontnetgroup" + DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                SetGroupProperty();
                KepGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(KepGroup_DataChange);

                KepGroup.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(KepGroup_AsyncWriteComplete);
                KepItems = KepGroup.OPCItems;


            }

            catch (Exception err)
            {
                MessageBox.Show("创建组出错：" + err.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;

            }

            return true;

        }

        /// <summary>
        /// 设置组属性
        /// </summary>
        /// <param name="TransactionID"></param>
        /// <param name="NumItems"></param>
        /// <param name="ClientHandles"></param>
        /// <param name="Errors"></param>
        private void SetGroupProperty()
        {
            kepGroups.DefaultGroupIsActive = true;
            kepGroups.DefaultGroupDeadband = 0;
            KepGroup.IsActive = true;
            KepGroup.IsSubscribed = true;
            KepGroup.UpdateRate = 100;
        }

        /// <summary>
        /// 获取连接后改变toolstrip内的状态。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void changeToolStrip()
        {

            toolStripButton3.Enabled = false;
            toolStripButton4.Enabled = true;
            toolStripButton5.Enabled = true;
            toolStripButton6.Enabled = true;
        }

        private void changeToolStrip1()
        {

            toolStripButton3.Enabled = true;
            toolStripButton4.Enabled = false;
            toolStripButton5.Enabled = false;
            toolStripButton6.Enabled = false;
        }

        /// <summary>
        /// 触发连接方法
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void GetToConnect()
        {
            //ServerConnect sc1 = new ServerConnect();
            if (sc1.ShowDialog() == DialogResult.OK)
            {
                //将服务器名设为root
                string root = sc1.listBox1.SelectedItem.ToString();
                TreeNode CountNode = new TreeNode(root);
                treeView1.Nodes.Add(CountNode);


                try
                {
                    // ShowListView();
                    //枚举服务器下所有tag点，并以树形显示
                    OPCBrowser browser = ServerConnect.KepServer.CreateBrowser();
                    LoadDataToTree(browser, treeView1.Nodes[0].Nodes);
                    list = RecurBrowse(browser);
                    changeToolStrip();



                }
                catch (Exception err)
                {

                    MessageBox.Show("初始化出错：" + err.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        /// <summary>
        /// 获取节点的所有值，显示在listview中
        /// </summary>
        private void ShowListView()
        {
            if (!CreateGroup())
            {

                return;
            }
            object value;
            object quality;
            object timestamp;
            //itmHandlerClient = 1234;
            OPCBrowser opcbrowser = ServerConnect.KepServer.CreateBrowser();
            list = RecurBrowse(opcbrowser);
            for (int i = 0; i < list.Count; i++)
            {

                KepItem = KepItems.AddItem(list[i], itmHandlerClient);
                itmHandlerServer = KepItem.ServerHandle;
                KepItem.Read((short)OPCDataSource.OPCDevice, out value, out quality, out timestamp);
                lv = new ListViewItem(list[i]);
                lv.SubItems.Add(Convert.ToString(value));
                lv.SubItems.Add(quality.ToString());
                lv.SubItems.Add(timestamp.ToString());
                listView1.Items.Add(lv);
            }
        }

        /// <summary>
        /// 获取listview中相应的值
        /// </summary>
        /// <param name="opcbrowser"></param>
        public List<string> RecurBrowse(OPCBrowser opcbrowser)
        {
            //  OPCBrowser opcbrowser = ServerConnect.KepServer.CreateBrowser();
            opcbrowser.ShowBranches();
            opcbrowser.ShowLeafs(true);
            list = new List<string>();
            foreach (object turn in opcbrowser)
            {
                list.Add(turn.ToString());
                //lv = new ListViewItem(turn.ToString());
                //listView1.Items.Add(lv);
            }
            return list;
        }

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

        #region 事件触发
        /// <summary>
        /// 写入tag值时执行的事件
        /// </summary>
        /// <param name="TransactionID"></param>
        /// <param name="NumItems"></param>
        /// <param name="ClientHandles"></param>
        /// <param name="Errors"></param>
        void KepGroup_AsyncWriteComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array Errors)
        {
            //lblState.Text = "";
            //for (int i = 1; i <= NumItems; i++)
            //{
            //    lblState.Text = "Tran:" + TransactionID.ToString() + " CH:" + ClientHandles.GetValue(i).ToString() + " Error:" + Errors.GetValue(i).ToString();
            //}
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
            Console.WriteLine("********" + TransactionID.ToString() + NumItems.ToString() + "*********");
            //for (int i = 1; i <= NumItems; i++)
            //{
            //    string number = itemNames[Convert.ToInt32(ClientHandles.GetValue(i)) - 1];
            //    itemValues[number] = ItemValues.GetValue(i).ToString();
            //    Console.WriteLine(itemValues[number] + "....." + number);

            //}
            //for (int i = 0; i <= NumItems; i++)
            //{
            //    //string name = itemNames[Convert.ToInt32(ClientHandles.GetValue(i))];
            //   // string name = itemNames[i-1];
                
            //        //int index = listView1.Columns[itemNames[i - 1]].Index;
            //        ////this.toolStripTextBox1.Text = ItemValues.GetValue(i).ToString();
            //        ////this.toolStripTextBox2.Text = Qualities.GetValue(i).ToString();
            //        ////this.toolStripTextBox3.Text = TimeStamps.GetValue(i).ToString();
            //        //listView1.Items[index].SubItems["value"].Text = ItemValues.GetValue(i).ToString();
            //        //listView1.Items[index].SubItems["status"].Text = Qualities.GetValue(i).ToString();
            //        //listView1.Items[index].SubItems["Time"].Text = TimeStamps.GetValue(i).ToString();
           // int index = listView1.Columns[itemNames[0]].Index;
          if(NumItems>0)
          {
              for (int i = 0; i < NumItems;i++ )
              {
                 // Console.WriteLine(itemNames[i]);
                  Console.WriteLine(listView1.Items[i].SubItems[tag.Index].Text);
                  listView1.Items[i].SubItems[Value.Index].Text = ItemValues.GetValue(i+1).ToString();
                  listView1.Items[i].SubItems[status.Index].Text = Qualities.GetValue(i+1).ToString();
                  listView1.Items[i].SubItems[Time.Index].Text = TimeStamps.GetValue(i+1).ToString();
                  Console.WriteLine(listView1.Items[i].SubItems[Value.Index].Text);
                  Console.WriteLine(listView1.Items[i].SubItems[status.Index].Text);
                  Console.WriteLine(listView1.Items[i].SubItems[Time.Index].Text);
                  //int index = listView1.Columns[itemNames[i]].Index;
                  //listView1.Items[index].SubItems["value"].Text = ItemValues.GetValue(i).ToString();
                  //listView1.Items[index].SubItems["status"].Text = Qualities.GetValue(i).ToString();
                  //listView1.Items[index].SubItems["Time"].Text = TimeStamps.GetValue(i).ToString();
              }
          }
               
            //}  
        }
    


        private void 连接ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetToConnect();

        }
        private void toolStripButton3_Click_1(object sender, EventArgs e)
        {
            GetToConnect();
            if (!CreateGroup())
            {

                return;
            }
        }

        /// <summary>
        /// 关于公司的介绍
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            AboutBox1 ab1 = new AboutBox1();
            ab1.Show();

        }

        /// <summary>
        /// Add Item 事件触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            Form2 fr2 = new Form2(this);
            fr2.ShowDialog();

        }



        private void Form1_Load(object sender, EventArgs e)
        {

            list1 = new List<string>();
        }

        #endregion

        #region 暂时性demo测试功能
        /// <summary>
        /// 暂时性测试 listview的属性和change事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_Click(object sender, EventArgs e)
        {


        }


        /// <summary>
        /// treeview选择后事件demo
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
              
            string itemName1 = getFullName(treeView1);
            //if (listView1.Items.IndexOfKey(itemName1) < 0)
            if (listView1.Items.IndexOfKey(itemName1) < 0)
            {
                ShowData(itemName1);
                list1.Add(itemName1);
                
            }

            else
            {
                MessageBox.Show("你已经选择过该项目，请直接在listview中查看即可。","提示");
            }
            //foreach (var item in list1)
            //{
            //    Console.WriteLine(item);
            //}
          

            Timer timer1 = new Timer();
            timer1.Enabled = true;
            timer1.Tick += timer1_Tick;
            timer1.Interval = 50000;



        }
        /// <summary>
        /// 获取所选treeview的完整itemname
        /// </summary>
        /// <param name="treeview1"></param>
        /// <returns></returns>
        public string getFullName(TreeView treeview1)
        {
            string fullname1 = treeView1.SelectedNode.FullPath;
            string fullname2 = treeView1.Nodes[0].FullPath + ".";
            string fullname = fullname1.ToString().Replace(fullname2, "");
            return fullname;
        }
        /// <summary>
        /// 通过itemname来获取相应的value值和时间并展示在listview中
        /// </summary>
        /// <param name="node"></param>
        public void ShowData(string node)
        {
            itemNames.Add(node);
            //itmHandlerClient = ++itmHandlerClient;
            KepItem = KepItems.AddItem(node, itmHandlerClient);
            itmHandlerServer = KepItem.ServerHandle;
            KepItem.Read(1, out value, out quality, out timestamp);
            listView1.BeginUpdate();
            lv = new ListViewItem(node);
            lv.Name = node;
            lv.SubItems.Add(Convert.ToString(value));
            lv.SubItems.Add(quality.ToString());
            lv.SubItems.Add(timestamp.ToString());
            listView1.Items.Add(lv);
            listView1.EndUpdate();
        }
        //public void ShowData(string[] node)
        //{
        //    listView1.BeginUpdate();
        //    for (int i = 0; i < node.Length; i++) { 
        //        KepItem = KepItems.AddItem(node[i], itmHandlerClient);
        //    itmHandlerServer = KepItem.ServerHandle;
        //    KepItem.Read(1, out value, out quality, out timestamp);
           
        //    lv = new ListViewItem(node[i]);
        //    lv.Name = node[i];
        //    lv.SubItems.Add(Convert.ToString(value));
        //    lv.SubItems.Add(quality.ToString());
        //    lv.SubItems.Add(timestamp.ToString());
        //    listView1.Items.Add(lv);
        //    }
        //    listView1.EndUpdate();
        //}

        #endregion
        /// <summary>
        /// 断开连接，执行事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 断开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (KepGroup != null)
                {
                    KepGroup.DataChange -= new DIOPCGroupEvent_DataChangeEventHandler(KepGroup_DataChange);
                }
                if (server != null)
                {
                    server.Disconnect();
                    server = null;
                }
                changeToolStrip1();

                treeView1.Nodes.Clear();
                listView1.Items.Clear();
                timer1.Enabled = false;
            }
            catch (Exception err)
            {

                MessageBox.Show(err.Message);
            }
        }
        /// <summary>
        /// 右键写入value值。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 写ItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OPCItem bItem = KepItems.GetOPCItem(itmHandlerServer);
            int[] temp = new int[2] { 0, bItem.ServerHandle };
            Array serverHandles = (Array)temp;
            object[] valueTemp = new object[2] { "", toolStripTextBox1.Text };
            Array values = (Array)valueTemp;
            Array Errors;
            int CancelID;
            KepGroup.AsyncWrite(1, ref serverHandles, ref values, out Errors, 2017, out CancelID);
            GC.Collect();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (KepGroup != null)
            {

                //KepItem.Read((short)OPCDataSource.OPCDevice, out value, out quality, out timestamp);


                listView1.Refresh();
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //string itemName1 = getFullName(treeView1);
           

            //if (listView1.Items.Count > 1 && listView1.Items.ContainsKey(itemName1))
            //{
            //    listView1.Items.Clear();
            //    foreach (string s in list1)
            //    {
            //        ShowData(s);
            //    }
            //}
            //else
            //{
            //    listView1.Items.Clear();
            //     itemName1 = getFullName(treeView1);
            //    ShowData(itemName1);

            //}



        }







    }
}
