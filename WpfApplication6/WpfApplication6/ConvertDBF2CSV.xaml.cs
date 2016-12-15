using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
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


namespace WpfApplication6
{
    /// <summary>
    /// Interaction logic for ConvertDBF2CSV.xaml
    /// </summary>
    public partial class ConvertDBF2CSV : Window
    {
        public ConvertDBF2CSV()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Select file";
            fdlg.InitialDirectory = @"c:\";
            fdlg.FileName = txtFileName.Text;
            fdlg.Filter = "DBF Files(*.dbf)|*.dbf|All Files(*.*)|*.*";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;
            //GetDataTableDBF(@"T:\RAImport\SWS\ul_ogval\model\ogval_v2\app.dbf");
            if (fdlg.ShowDialog() == true)
            {
                txtFileName.Text = fdlg.FileName;
                dg.ItemsSource = GetDataTableDBF(txtFileName.Text).DefaultView;
            }
        }
        public static DataTable GetDataTableDBF(string strFileName)
        {
            System.Data.Odbc.OdbcConnection conn = new System.Data.Odbc.OdbcConnection("Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" + System.IO.Path.GetFullPath(strFileName).Replace(System.IO.Path.GetFileName(strFileName), "") + ";Exclusive=No");
            conn.Open();
            string strQuery = "SELECT * FROM [" + System.IO.Path.GetFileName(strFileName) + "]";
            System.Data.Odbc.OdbcDataAdapter adapter = new System.Data.Odbc.OdbcDataAdapter(strQuery, conn);
            System.Data.DataSet ds = new System.Data.DataSet();
            adapter.Fill(ds);
            return ds.Tables[0];
        }

        private void button1_Click(object sender, RoutedEventArgs e)
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
            dg.SelectAll();

            dg.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
            ApplicationCommands.Copy.Execute(null, dg);
            dg.UnselectAllCells();


        }
    }
    }

