using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.Script.Services;
using System.Web.Services;
using ChartsGenerator.Model;

namespace ChartsGenerator
{
    public partial class NewDashboard : System.Web.UI.Page
    {
        private static DataTable _cData;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                if (Session["FPath"] == null)
                {
                    Response.Redirect("Home.aspx");
                }
                else
                {
                    var filepath = HttpContext.Current.Session["FPath"].ToString();
                    _cData = ConvertExcelToDataTable(filepath);
                    ImportToGrid();
                }

            }
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public static object GetProjectCount(string sDate, string eDate)
        {
            if (string.IsNullOrWhiteSpace(sDate) || string.IsNullOrWhiteSpace(eDate))
                return _cData.AsEnumerable().Select(r => r.Field<string>("Project")).Distinct();

            var startDate = DateTime.ParseExact(sDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            var endDate = DateTime.ParseExact(eDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            var tempData = _cData.Clone();
            foreach (DataRow row in _cData.Rows)
            {
                var startDateStr = row["StartDate"] != DBNull.Value ? row["StartDate"] : "";
                if (string.IsNullOrWhiteSpace(startDateStr.ToString()))
                    continue;
                var stDate = DateTime.Parse(startDateStr.ToString().Trim());
                
                var endDateStr = row["EndDate"] != DBNull.Value ? row["EndDate"] : "";
                if (string.IsNullOrWhiteSpace(endDateStr.ToString().Trim()))
                    continue;

                var edDate = DateTime.Parse(endDateStr.ToString().Trim());

                if (stDate >= startDate && edDate <= endDate)
                    tempData.ImportRow(row);
            }

            var pData = tempData.AsEnumerable().Select(r => r.Field<string>("Project")).Distinct();


            return pData;
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public static object[] GetChartData(string name, string s_Date, string e_Date)
        {
            var data = new List<ChartData>();
            foreach (DataRow row in _cData.Rows)
            {
                var startDate = row["StartDate"] != DBNull.Value ? row["StartDate"] : "";
                if (string.IsNullOrWhiteSpace(startDate.ToString()))
                    continue;
                var stDate = DateTime.Parse(startDate.ToString().Trim());

                var endDate = row["EndDate"] != DBNull.Value ? row["EndDate"] : "";
                if (string.IsNullOrWhiteSpace(endDate.ToString().Trim()))
                    continue;

                var eDate = DateTime.Parse(endDate.ToString().Trim());

                data.Add(new ChartData
                {
                    Project = row["Project"].ToString(),
                    Phase = row["Phase"].ToString(),
                    Task = row["Task"].ToString(),
                    //Task = "",
                    StartDate = stDate,
                    EndDate = eDate,
                    Fleet = row["Fleet"].ToString()
                });
            }

            List<ChartData> tempData;
            if (string.IsNullOrWhiteSpace(s_Date) || string.IsNullOrWhiteSpace(e_Date))
                tempData = data;
            else
            {
                var start_Date = DateTime.ParseExact(s_Date, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                var end_Date = DateTime.ParseExact(e_Date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                tempData = data.Where(i => i.StartDate >= start_Date && i.EndDate <= end_Date).ToList();
            }

            var date = tempData.OrderBy(x => x.StartDate).Select(x => x.StartDate).Distinct().FirstOrDefault();

            var newdata = new List<ChartData>();
            foreach (var project in tempData.Select(x => x.Project).Distinct().ToList())
            {
                var projectData = tempData.Where(x => x.Project == project).ToList();
                if(projectData.Count<=0)
                    continue;

                newdata.Add(new ChartData()
                {
                    StartDate = date,
                    EndDate = date.AddHours(4),
                    Phase = project+".......................",
                    Project = project,
                    Task = ""
                });
                foreach (var chartData1 in projectData)
                {
                    chartData1.Phase = chartData1.Fleet + " " + chartData1.Phase;
                    newdata.Add(chartData1);
                }
            }
            
            int j = 0;


            var chartData = new object[newdata.Count + 1];
            chartData[0] = new object[]{
                "Project",    
                "Phase",
                "Task",
                "StartDate",
                "EndDate"
                };
            foreach (var i in newdata)
            {
                j++;
                chartData[j] = new object[] { i.Project, i.Phase, i.Task, i.StartDate, i.EndDate };
            }
            return chartData;
        }

        public static DataTable ConvertExcelToDataTable(string fileName)
        {
            DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }


        public void ImportToGrid()
        {
            var connString = "";
            //var path = HostingEnvironment.MapPath("~/input/template.xlsx");
            var path = HttpContext.Current.Session["FPath"].ToString();
            const string strFileType = ".xlsx";
            //Connection String to Excel Workbook
            if (strFileType.Trim() == ".xls")
            {
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else if (strFileType.Trim() == ".xlsx")
            {
                //connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';";
            }
            const string query = "SELECT [Project], [Phase], [Task], [Duration], [StartDate], [EndDate]  FROM [Sheet1$]";
            var conn = new OleDbConnection(connString);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            var cmd = new OleDbCommand(query, conn);
            var da = new OleDbDataAdapter(cmd);
            var ds = new DataSet();
            da.Fill(ds);
            //grvExcelData.DataSource = ds.Tables[0];
            //grvExcelData.DataBind();
            da.Dispose();
            conn.Close();
            conn.Dispose();
        }

        protected void PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            ImportToGrid();
            //grvExcelData.PageIndex = e.NewPageIndex;
            //grvExcelData.DataBind();
        }
    }
}