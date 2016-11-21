using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
using Excel = Microsoft.Office.Interop.Excel; 
 

namespace WpfApplication12
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // test connection to database
            Connection instanta = new Connection();
            instanta.user = UserTxt.Text;
            instanta.pass = PassTxt.Text;
            instanta.host = HostTxt.Text;
            instanta.port = PortTxt.Text;
            instanta.service = ServiceTxt.Text;
            instanta.OpenDb();
            instanta.CloseDb();

            
           // instanta.CloseDb();
        }

        public void Button_Click_1(object sender, RoutedEventArgs e)
        {
            //create a new connection
            string x = SqlTxt.Text;
            Connection instanta = new Connection();
            instanta.user = UserTxt.Text;
            instanta.pass = PassTxt.Text;
            instanta.host = HostTxt.Text;
            instanta.port = PortTxt.Text;
            instanta.service = ServiceTxt.Text;
           // if sql text starts with select take to datagrid
            instanta.OpenDb();
            if (x.Substring(0,6)=="select")
            {
            try
            {
                dataGridView1.ItemsSource = instanta.FillDataGrid(x).Tables[0].DefaultView;
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString());
            }
            }

            else {
                //else update sqltxt 
                MessageBox.Show("Only Select commands");
                SqlTxt.Text = instanta.Sql(x);
            }
                 //instanta.Sql(x);
            instanta.CloseDb();

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            //create a new object excel application
            Excel.Application xlapp = new Excel.Application();
            xlapp.Visible = true;
            //new workbook and new worksheet
            Excel.Workbook wb;
            Excel.Worksheet ws;
            object misValue = System.Reflection.Missing.Value;
           
            //assign object
           wb = xlapp.Workbooks.Add(misValue);
           ws = wb.Worksheets.get_Item(1);
           //ws.Cells[1,1] = "kosmin";

           copyAlltoClipboard();

           Excel.Range CR = (Excel.Range)ws.Cells[1, 1];
           CR.Select();
           ws.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
           Excel.Range cr1 = ws.Cells[1, 1];
           cr1.Select();
          
        }

        //copy to clipboard all data from gridview
        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();

            dataGridView1.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
            ApplicationCommands.Copy.Execute(null, dataGridView1);
            dataGridView1.UnselectAllCells();

            
        }

        
    

    
    }
}
