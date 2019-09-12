using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;//using Excel
using System.Windows.Forms;
using System.Configuration;

namespace ExcelAddIn
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //TODO:加载初始化数据，如数据库连接
            string appsetting = ConfigurationManager.AppSettings["excel"].ToString();
            
        }

        private void System_Group_Click(object sender, RibbonControlEventArgs e)
        {
            //声明一个Excel的单元格区域变量。
            Excel.Range rang;

            //获得Excel所选区域
            rang = Globals.ThisAddIn.Application.Selection;

            //所选区域底纹颜色设置为黄色。
            //rang.Interior.Color = System.Drawing.Color.Yellow;            
            int columns = rang.Columns.Count+1;
            int rows = rang.Rows.Count+1;
            
            // 遍历时先列后行，否则会出现矩阵反转
            for (int i = 1; i < columns; i++)
            {
                for (int k = 1; k < rows; k++)
                {
                    Excel.Range cell  = rang.Item[i][k];
                    if(cell.Value2 == null)
                    {
                        continue;
                    }
                    string cellValue = ((string)cell.Value2).ToString();
                    if (!cellValue.Contains(','))
                    {
                        cell.Value2 = "'" + cellValue.Trim() + "'";
                    }else
                    {
                        cellValue = cellValue.Replace(",", "','");
                        cell.Value2 = "'" + cellValue.Trim() + "'";
                    }
                    
                }
            }
            
        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            ConfigPanel fm1 = new ConfigPanel();//创建新实例。
            fm1.ShowDialog();//以对话框的形式，在Excel中显示Form1。
        }
    }
}
