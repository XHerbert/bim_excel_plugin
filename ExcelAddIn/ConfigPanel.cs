using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Configuration;
using CCWin.SkinControl;
using System.Reflection;

namespace ExcelAddIn
{
    public partial class ConfigPanel : Form
    {
        public ConfigPanel()
        {
            InitializeComponent();
        }

        private void skinButton2_Click(object sender, EventArgs e)
        {
            string appsetting = ConfigurationManager.ConnectionStrings["mysql"].ToString();
            MySqlConnection conn = new MySqlConnection(appsetting);
            conn.Open();
            string state = conn.State.ToString();
            string message = state == "Open" ? "测试连接成功！" : "测试连接失败！";
            CCWin.MessageBoxEx.Show(message);
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {
            string assemblyConfigFile = Assembly.GetExecutingAssembly().Location;
            string appDomainConfigFile = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);


            //获取appSettings节点
            AppSettingsSection appSettings = (AppSettingsSection)config.GetSection("connectionStrings");

            //获取连接串节点
            ConnectionStringsSection connectionSettings = (ConnectionStringsSection)config.GetSection("connectionStrings");


            //删除name，然后添加新值
            //appSettings.Settings.Remove("name");
            appSettings.Settings.Add("name", "new");

            //connectionSettings.ConnectionStrings.


            //保存配置文件
            config.Save();
        }
    }
}
