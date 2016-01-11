using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Hosting;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Script.Services;
using System.Web.Services;
using ChartsGenerator.Model;

namespace ChartsGenerator
{
    public partial class Dashboard : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                //DataTable cData = ConvertExcelToDataTable("C:/Users/Deepak 1/Downloads/TimelineChart.xlsx");
            }
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public static object GetProjectCount()
        {
            var filepath = HostingEnvironment.MapPath("~/input/template.xlsx");
            var cData = ConvertExcelToDataTable(filepath);
            var pData = cData.AsEnumerable().Select(r => r.Field<string>("Project")).Distinct();
            return pData;
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public static object[] GetChartData(string name)
        {
            var filepath = HostingEnvironment.MapPath("~/input/template.xlsx");
            DataTable cData = ConvertExcelToDataTable(filepath);
            var data = new List<ChartData>();
            data = new List<ChartData>();
            foreach (DataRow row in cData.Rows)
            {
                var startDate = row["StartDate"] != DBNull.Value ? row["StartDate"] : "";
                if (string.IsNullOrWhiteSpace(startDate.ToString().Trim()))
                    continue;
                var stDate = DateTime.Parse(startDate.ToString().Trim());

                if (string.IsNullOrWhiteSpace(startDate.ToString().Trim()))
                    continue;

                var endDate = row["EndDate"] != DBNull.Value ? row["EndDate"] : "";
                if (string.IsNullOrWhiteSpace(endDate.ToString().Trim()))
                    continue;

                var eDate = DateTime.Parse(endDate.ToString().Trim());

                data.Add(new ChartData
                {
                    Project = row["Project"].ToString(),
                    Phase = row["Phase"].ToString(),
                    Task = row["Task"].ToString(),
                    StartDate = stDate,
                    EndDate = eDate
                });
            }

            var newdata = data.Where(x => x.Project == name).ToList();
            
            var chartData = new object[newdata.Count + 1];
                chartData[0] = new object[]{
                "Project",    
                "Phase",
                "Task",
                "StartDate",
                "EndDate"
                };
            int j = 0;
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

        protected void UploadBtn_Click(object sender, EventArgs e)
        {

        }
    }
}