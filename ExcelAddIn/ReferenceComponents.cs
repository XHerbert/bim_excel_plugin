using ExcelAddIn.Events;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn
{
    public partial class ReferenceComponents : Form
    {

        public void SetReferComponents(string text,string objectId)
        {
            this.skinTextBox1.Text = text;
            this.skinGroupBox1.Tag = objectId;
        }

        public ReferenceComponents()
        {
            InitializeComponent();
        }

        public void AfterParentFrmTextChange(object sender, EventArgs e)
        {
            CustonEventArgs arg = e as CustonEventArgs;
            this.SetReferComponents(arg.Text,arg.ObjectId);
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this.skinGroupBox1.Tag.ToString());
            //TODO: 更新数据库关联关系
        }
    }
}
