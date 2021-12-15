using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SYstem.Windows.Forms;

using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;


namespace ControlDatabase
{
  public partial class Form1 : Form
  {
    public Form1()
    {
      InitializeComponent();
    }
    
    private void dbConnection_Button_Click(object sender, EventArgs e)
    {
      // DB 정보
      string myConnection = @"datasource = localhost;
                              port       = 3306;
                              username   = root;
                              password   = eunhwan0!!;
                              charset    = utf8";
      
      // DB 연결 객체 생성
      MySqlConnection myConn =  new MySqlConnection(myConnection);
      
      //DB 연결
      try
      {
        //MySqlDataAdapter myDataAdapter = enw MySqlDataAdapter();
        //myDataAdapter.SelectCommand = new MySqlCommand("Select * database.edata;", myConn);
        //MySqlCommandBuilder cb = new MySqlCommandBuilder(myDataAdapter);
        //Dataset ds = new Dataset();
        
        myConn.Open();
        
        MessageBox.Show("데이터베이스와 연결되었습니다.");
      }
      catch (Exception ex)
      {
        MessageBox.Show("데이터베이스 연결에 실패했습니다.\n" + ex.Message);
      }
      finally
      {
        myConn.Close();
      }  
    }
    
    private void inputData_Button_Click(object sender, EventArgs e)
    {
      // LP 데이터 경로
      System.IO.DirectoryInfo currentPath = new System.IO.DirectoryInfo(Application.StartupPath);
      string lpData_Path = Path.GetFullPath(Path.Combine(currentPath.ToString(), @"..\..\..\KRX데이터_LP체결네역\"));
      string[] fileList = Directory.GetFiles(lpData_Path, "*.csv");
      
      string[] fileNameList = fileList.Select(x => Path.GetFileName(x)).ToArray<string>();
      
      MessageBox.Show(String.Format("경로상 LP 데이터 파일의 수는 {0}개 입니다.", fileNameList.Length.ToString()));
      
      foreach (string fileName in fileNameList)
      {
        //1. 각 파일을 열어서 필요한 데이터 추출
        get_Data_from_EachFile(fileName);
        
        //2. 빼낸 데이터를 DB에 넣는 작업 필요
        //List<Array> insertData = new List<Array>();
        
        //input_EachFile_to_DB(fileName);
      }
      
      // DB 연결
      string myConnection = @"datasource = localhost;
                              port       = 3306;
                              username   = root;
                              password   = eunhwan0!!;
                              charset    = utf8;
                              
      // Query 실행
      using (MySqlConnection myConn = new MySqlConnection(myConnection))
      {
        myConn.Open();
        
        string sqlCommandString = @"INSERT INTO test_lp_data VALUES ()";
        
        MySqlCommand sqlCommand = new MySqlCommand(sqlCommandString);
        sqlCommand.ExecuteNonQuery();
      }
    }
    
    private void get_Data_from_EachFile(string fileName)
    {
      //OpenFileDialog openFileDialog = new OpenFileDialog();
      //openFileDialog.Title = "엑셀 파일을 선택하세요.";
      //openFileDialog.Filter = "Excel Files|*csv";
      //DialogResult result = openFiledialog.ShowDialog();
      
      //if (result == DialogResult.OK)
      
      Excel.Application app = new Excel.Application();
      app.DisplayAlerts = false;
      app.Visible = false;
      app.ScreenUpdating = false;
      app.DisplayStatusBar = false;
      app.EnableEvents = false;
      
      Excel.Workbooks workbooks = app.Workbooks;
      
      // 엑셀 워크북 (파일 경로 읽어서)
      Excel.Workbook workbook = workbooks.Open(fileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
      
      // 엑셀 파일 이름 확ㅇ니
      string excelFileName = workbook.Name;
      string[] str = excelFileName.Split('.');
      
      // 엑셀 워크시트 객체
      Excel.Sheets sheets = workbook.Worksheets;
      Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(1);
      //Excel.Worksheet worksheet = sheets.get_Item(1) as Excel.Worksheet;
      
      // 워크시트 첫 번째 이름
      string workSheetName = worksheet.Name;
      
      try
      {
        if (str[1].Equals("csv"))
        {
          // 연결 string
          string xlConn = "Provider = Microsoft.Jet.OLEDB.4.0;" +
            "Data Source = " + fileName + ";" +
            "Extended Properties = 'Excel 12.0 XML; HDR = YES; IMEX = 1';";
            
          OleDbConnection conn = new OleDbConnection(xlConn);
          
          string sqlQuery = @"SELECT * FROM [" + workSheetName + "$";
          OleDbCommand cmd = enw OleDbCommand(sqlQuery, conn);
          
          conn.Open()
          
          OleDbDataAdapter sda = new OleDbDataAdapter(cmd);
          DataTable dt = new DataTable();
          sda.Fill(dt);
          //DataGridiew1.DataSource = dt;
          
          conn.Close();
        }
        catch (Exception ex)
        {
          MessageBox.Show(ex.ToString());
          workbook.Close(true, null, null);
          app.Quit();
        }
        finally
        {
          workbook.Close(true, null, null);
          
          //메모리 할당 해제
          deleteObject(worksheet);
          deleteObject(workbook);
          app.Quit();
          deleteObject(app);
        }
      }
      
      // 메모리 해제를 위한 사용자 정의 함수
      private void deleteObject(object obj)
      {
        try
        {
          System.Runtime.InteropServices.Marshal.ReleaseComObejct(obj);
          obj = null;
        }
        catch (Exception ex)
        {
          obj = null;
          MessageBox.Show("메모리 할당을 해제하는 중 문제가 발생하였습니다." + ex.ToString(), "경고!");
        }
        finally
        {
          GC.Collect();
        }
        
        private void input_EachFile_to_DB(string fileName)
        {
        
        }
      }
    }
    
    
    
  }
}
