using ADOX;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfDemo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Thread thread_pBar;

        Thread thread_cal;

        Int64 m_num = 0;
        Int64 m_sum = 0;
        Int64 m_times = 0;


        public MainWindow()
        {
            InitializeComponent();

            //创建mdb文件
            ADOX.Catalog catalog = new Catalog();

            string filePath = "D:\\test.mdb";

            if (!File.Exists(filePath))
            {
                try
                {
                    catalog.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath);
                    //catalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+filePath);

                    //创建表
                    ADODB.Connection cn = new ADODB.Connection();
                    //cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath, null, null, -1);
                    cn.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath, null, null, -1);
                    catalog.ActiveConnection = cn;

                    ADOX.Table table = new ADOX.Table();
                    //创建表
                    table.Name = "result";
                    //创建列
                    ADOX.Column column = new ADOX.Column();
                    //ParentCatalog,指定表、用户或列对象的父目录，以提供对特定于访问接口的属性的访问。
                    column.ParentCatalog = catalog;
                    column.Name = "id";
                    column.Type = DataTypeEnum.adInteger;
                    column.DefinedSize = 9;
                    //属性为自增，每次追加自动增加
                    column.Properties["AutoIncrement"].Value = true;

                    table.Columns.Append(column, DataTypeEnum.adInteger, 9);
                    table.Keys.Append("FirstTablePrimaryKey", KeyTypeEnum.adKeyPrimary, column, null, null);
                    table.Columns.Append("input", DataTypeEnum.adBigInt, 50);
                    table.Columns.Append("sum", DataTypeEnum.adBigInt, 50);
                    catalog.Tables.Append(table);

                }
                catch (System.Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            }


        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void bttn_close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void setProgressBar()
        {

            for (int i = 0; i <= m_num; i++)
            {
                //当前进度，最大值默认100
                pBar_run.Dispatcher.BeginInvoke((ThreadStart)delegate { this.pBar_run.Value = i; });
                //pBar_run.Value = i;
                Thread.Sleep(10);
            }
        }

        private void calSum()
        {

            for (Int64 i = 0; i < m_num; i++)
            {
                m_sum += i;
            }

            string filePath = "D:\\test.mdb";

           //INSERT INTO table_name
            //VALUES(value1, value2, value3,...);

            //INSERT INTO table_name(column1, column2, column3,...)
            //VALUES(value1, value2, value3,...);

            string strInsert = " INSERT INTO result VALUES (" + m_times.ToString() + "," + m_num.ToString() + "," + m_sum.ToString() + ")";

            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + filePath);
            OleDbCommand cmd = new OleDbCommand(strInsert, con); //定义Command对象
            con.Open(); //打开数据库连接
            cmd.ExecuteNonQuery(); //执行Command命令
            con.Close(); //关闭数据库连接
            txtBlk_result.Dispatcher.BeginInvoke((ThreadStart)delegate { this.txtBlk_result.Text= m_sum.ToString(); });


            m_times++;

        }

        private void bttn_run_Click(object sender, RoutedEventArgs e)
        {

            Int64 num = 0;
            //Int32 num = 0;

            //num = Convert.ToInt32(txtBx_input.Text);

            try
            {
                //num = Convert.ToInt32(txtBx_input.Text);
                m_num = Convert.ToInt64(txtBx_input.Text);
                //num = Convert.ToInt64(txtBx_input.Text);
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString());

            }

            thread_pBar = new Thread(setProgressBar);
            thread_pBar.IsBackground = true;
            thread_pBar.Start();

            thread_cal = new Thread(calSum);
            thread_cal.IsBackground = true;
            thread_cal.Start();

            //Int64 sum = 0;
            //for(Int64 i = 0; i < num; i++)
            //{
            //    sum+=i;
            //}

            //txtBlk_result.Text = sum.ToString(); 
        }

        private void bttn_clear_Click(object sender, RoutedEventArgs e)
        {
            txtBx_input.Text = "0";
            txtBlk_result.Text = "0";
            m_sum = 0;
            thread_cal.Abort();
            thread_pBar.Abort();

        }
    }
}
