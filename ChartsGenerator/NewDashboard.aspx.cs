﻿using System;
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
        private static List<ChartData> _chartData;

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
                    _chartData = new List<ChartData>();
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

                        _chartData.Add(new ChartData
                        {
                            Project = row["Project"].ToString(),
                            Phase = row["Phase"].ToString(),
                            Task = row["Task"].ToString(),
                            //Task = "",
                            StartDate = stDate,
                            EndDate = eDate,
                            Fleet = row["Fleet"].ToString(),
                            Color = row["Color"].ToString(),
                            Vendor = row["Vendor"].ToString()
                        });
                    }
                    var projects = _chartData.Select(x => x.Project).Distinct();
                    foreach (var project in projects)
                    {
                        lstFleet.Items.Add(new ListItem(project));
                    }

                    var phases = _chartData.Select(x => x.Phase).Distinct();
                    foreach (var phase in phases)
                    {
                        lstPhase.Items.Add(phase);
                    }

                    var tasks = _chartData.Select(x => x.Task).Distinct();
                    foreach (var task in tasks)
                    {
                        lstTasks.Items.Add(task);
                    }

                    var vendor = _chartData.Select(x => x.Vendor).Distinct();
                    foreach (var item in vendor)
                    {
                        lstVendor.Items.Add(item);
                    }

                    ImportToGrid();
                }

            }
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public static object GetProjectCount(string sDate, string eDate, string fleet, string phase, string task, string vendor)
        {
            fleet = fleet ?? "";
            phase = phase ?? "";
            vendor = vendor ?? "";
            task = task ?? "";

            var fleets = fleet.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var phases = phase.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var vendors = vendor.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var tasks = task.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            if ((string.IsNullOrWhiteSpace(sDate) || string.IsNullOrWhiteSpace(eDate))
                && fleets.Length == 0 && phases.Length == 0 && tasks.Length == 0 && vendors.Length ==0)
                return _chartData.Select(x => x.Project).Distinct();

            var tempData = _chartData;

            if (!string.IsNullOrWhiteSpace(sDate) && !string.IsNullOrWhiteSpace(eDate))
            {
                var startDate = DateTime.ParseExact(sDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                var endDate = DateTime.ParseExact(eDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                tempData = tempData.Where(x => x.StartDate >= startDate && x.EndDate <= endDate).ToList();
            }

            if (fleets.Length > 0)
                tempData = tempData.Where(x => fleets.Contains(x.Project)).ToList();

            if (phases.Length > 0)
                tempData = tempData.Where(x => phases.Contains(x.Phase)).ToList();

            if (vendors.Length > 0)
                tempData = tempData.Where(x => vendors.Contains(x.Vendor)).ToList();

            if (tasks.Length > 0)
                tempData = tempData.Where(x => !tasks.Contains(x.Task)).ToList();

            var pData = tempData.AsEnumerable().Select(r => r.Project).Distinct();

            return pData;
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public static object[] GetChartData(string sDate, string eDate, string fleet, string phase, string task, string vendor)
        {
            fleet = fleet ?? "";
            phase = phase ?? "";
            vendor = vendor ?? "";
            task = task ?? "";

            var fleets = fleet.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var phases = phase.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var vendors = vendor.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var tasks = task.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            List<ChartData> tempData;

            if ((string.IsNullOrWhiteSpace(sDate) || string.IsNullOrWhiteSpace(eDate))
                && fleets.Length == 0 && phases.Length == 0 && vendors.Length == 0)
                tempData = _chartData;
            else
            {
                tempData = _chartData;
                if (!string.IsNullOrWhiteSpace(sDate) && !string.IsNullOrWhiteSpace(eDate))
                {
                    var startDate = DateTime.ParseExact(sDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                    var endDate = DateTime.ParseExact(eDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                    tempData = tempData.Where(x => x.StartDate >= startDate && x.EndDate <= endDate).ToList();
                }

                if (fleets.Length > 0)
                    tempData = tempData.Where(x => fleets.Contains(x.Project)).ToList();

                if (phases.Length > 0)
                    tempData = tempData.Where(x => phases.Contains(x.Phase)).ToList();

                if (vendors.Length > 0)
                    tempData = tempData.Where(x => vendors.Contains(x.Vendor)).ToList();
                
                if (tasks.Length > 0)
                    tempData = tempData.Where(x => !tasks.Contains(x.Task)).ToList();
            }

            // 
            var date = tempData.OrderBy(x => x.StartDate).Select(x => x.StartDate).Distinct().FirstOrDefault();

            var newdata = new List<ChartData>();
            foreach (var project in tempData.Select(x => x.Project).Distinct().ToList())
            {
                var projectData = tempData.Where(x => x.Project == project).ToList();
                if (projectData.Count <= 0)
                    continue;

                newdata.Add(new ChartData()
                {
                    StartDate = date,
                    EndDate = date.AddHours(4),
                    Phase = project + ".......................",
                    Project = project,
                    Task = "",
                    Fleet = "",
                    Color = "#aaaaaa",
                    Vendor = ""
                });
                newdata.AddRange(projectData);
            }

            int j = 0;


            var chartData = new object[newdata.Count + 1];
            chartData[0] = new object[]{
                "Project",    
                "Phase",
                "Task",
                "StartDate",
                "EndDate",
                "Fleet",
                "Color",
                "Vendor"
                };
            foreach (var i in newdata)
            {
                j++;
                chartData[j] = new object[] { i.Project, i.Phase, i.Task, i.StartDate, i.EndDate, i.Fleet, i.Color, i.Vendor };
            }
            return chartData;
        }

        protected void FleetSelected(object sender, EventArgs e)
        {
            var project = lstFleet.SelectedValue;
            if (project == "--SELECT--") return;

            lstPhase.Items.Clear();
            lstPhase.Items.Add("--SELECT--");

            var phases = _chartData.Where(x => x.Project == project).Select(x => x.Phase).Distinct();
            foreach (var phase in phases)
            {
                lstPhase.Items.Add(phase);
            }
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
            const string query = "SELECT [Project], [Phase], [Task], [Duration], [StartDate], [EndDate], [Vendor]  FROM [Sheet1$]";
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