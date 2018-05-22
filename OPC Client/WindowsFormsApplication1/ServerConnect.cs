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
using System.Net;

namespace WindowsFormsApplication1
{
    public partial class ServerConnect : Form
    {
        public ServerConnect()
        {
            InitializeComponent();
        }
        #region 私有变量


        public static OPCServer KepServer;
        bool opc_connected = false;
        public static string strHostIP = "";
        public static string strHostName = "";


        #endregion

        #region 方法
        private void GetLocalServer()
        {
            //获取本地计算机ip，计算机名称
            IPHostEntry IPHost = Dns.Resolve(Environment.MachineName);
            //IPHostEntry IPHost = Dns.GetHostEntry(Environment.MachineName);
            if (IPHost.AddressList.Length > 0)
            {
                strHostIP = IPHost.AddressList[0].ToString();

            }
            else
            {
                return;
            }
            //通过ip来获取计算机名称，可用于局域网内
            IPHostEntry ipHostEntry = Dns.GetHostByAddress(strHostIP);
            strHostName = ipHostEntry.HostName.ToString();
            //获取本地计算机上的opcservername
            try
            {
                KepServer = new OPCServer();
                object serverList = KepServer.GetOPCServers(strHostName);
                foreach (string turn in (Array)serverList)
                {
                    listBox1.Items.Add(turn);

                }
            }
            catch (Exception err)
            {

                MessageBox.Show("枚举本地服务器出错：" + err.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        /// <summary>
        /// 连接opc服务器
        /// </summary>
        /// <param name="TransactionID"></param>
        /// <param name="NumItems"></param>
        /// <param name="ClientHandles"></param>
        /// <param name="Errors"></param>
        /// 
        private static bool ConnectRemoteServer(string remoteServerIP, string remoteServerName)
        {
            try
            {
                KepServer.Connect(remoteServerName, remoteServerIP);
                if (KepServer.ServerState == (int)OPCServerState.OPCRunning)
                {

                    //  tsslServerState.Text = "已连接到：" + KepServer.ServerName + " ";

                }
                else
                {
                    //  tsslServerState.Text = "状态：" + KepServer.ServerState.ToString() + " ";
                }
            }
            catch (Exception err)
            {

                MessageBox.Show("连接到服务器出现错误：" + err.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        #endregion

        #region 事件触发
        private void ServerConnect_Load(object sender, EventArgs e)
        {
            GetLocalServer();
        }

        public void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = listBox1.SelectedItem.ToString();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ConnectRemoteServer(strHostIP, textBox1.Text))
                {
                    return;
                }
                opc_connected = true;
                this.Hide();


            }
            catch (Exception err)
            {

                MessageBox.Show("初始化出错：" + err.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        #endregion


    }
}
