using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.Script.Services;
using System.Web.Services;
using ChartsGenerator.Model;
using DataTable = System.Data.DataTable;

namespace ChartsGenerator
{
    public partial class NewDashboard : System.Web.UI.Page
    {
        private static List<ChartData> _chartData;
        private static List<ColorData> _colorData;


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
                    try
                    {
                        var filepath = HttpContext.Current.Session["FPath"].ToString();
                        var cData = ConvertExcelToDataTable(filepath, "Data");

                        var colorDataTable = ConvertExcelToDataTable(filepath, "ColorCodes");

                        _colorData = (from DataRow row in colorDataTable.Rows
                                      select new ColorData()
                                      {
                                          Color = row["Color"].ToString(),
                                          Task = row["Task"].ToString(),
                                      }).ToList();

                        _chartData = new List<ChartData>();
                        foreach (DataRow row in cData.Rows)
                        {
                            var startDate = row["StartDate"] != DBNull.Value ? row["StartDate"] : "";
                            if (string.IsNullOrWhiteSpace(startDate.ToString()))
                                continue;
                            var stDate = DateTime.Parse(startDate.ToString().Trim());

                            var endDate = row["EndDate"] != DBNull.Value ? row["EndDate"] : "";
                            if (string.IsNullOrWhiteSpace(endDate.ToString().Trim()))
                                continue;

                            var eDate = DateTime.Parse(endDate.ToString().Trim());

                            var chartData = new ChartData
                            {
                                Project = row["Project"].ToString(),
                                Phase = row["Phase"].ToString(),
                                Task = row["Task"].ToString(),
                                StartDate = stDate,
                                EndDate = eDate,
                                Fleet = row["Fleet"].ToString(),
                                Vendor = row["Vendor"].ToString(),
                                Color = row["Color"].ToString()
                            };

                            if (chartData.Task.Contains("EOI"))
                            {
                                Response.Write("");
                            }
                            var color = _colorData.FirstOrDefault(x => x.Task == chartData.Task);
                            if (color != null)
                                chartData.Color = color.Color;
                            else
                            {
                                Response.Write("");
                            }
                            _chartData.Add(chartData);
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
                        Error.Visible = false;
                    }
                    catch (Exception ex)
                    {
                        Error.Visible = true;
                        Error.Text = ex.Message;
                    }
                }

            }
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public static object GetProjectCount(string sDate, string eDate, string fleet, string phase, string task, string vendor)
        {
            try
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
                    && fleets.Length == 0 && phases.Length == 0 && tasks.Length == 0 && vendors.Length == 0)
                    return _chartData.Select(x => x.Project).Distinct();

                var tempData = _chartData;

                if (!string.IsNullOrWhiteSpace(sDate) && !string.IsNullOrWhiteSpace(eDate))
                {
                    var startDate = DateTime.ParseExact(sDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                    var endDate = DateTime.ParseExact(eDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                    tempData = tempData.Where(x => x.EndDate >= startDate && x.StartDate <= endDate).ToList();
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
            catch (Exception ex)
            {
                return null;
            }
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public static object[] GetChartData(string sDate, string eDate, string fleet, string phase, string task, string vendor)
        {
            fleet = fleet ?? "";
            phase = phase ?? "";
            vendor = vendor ?? "";
            task = task ?? "";

            var startDate = DateTime.MinValue;
            var endDate = DateTime.MinValue;

            var fleets = fleet.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var phases = phase.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var vendors = vendor.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            var tasks = task.Replace("null", "").Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            List<ChartData> tempData;

            if ((string.IsNullOrWhiteSpace(sDate) || string.IsNullOrWhiteSpace(eDate))
                && fleets.Length == 0 && phases.Length == 0 && tasks.Length == 0 && vendors.Length == 0)
                tempData = _chartData;
            else
            {
                tempData = _chartData;
                if (!string.IsNullOrWhiteSpace(sDate) && !string.IsNullOrWhiteSpace(eDate))
                {
                    startDate = DateTime.ParseExact(sDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                    endDate = DateTime.ParseExact(eDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                    tempData = tempData.Where(x => x.EndDate >= startDate && x.StartDate <= endDate).ToList();
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

            if (date < startDate)
                date = startDate;

            var newdata = new List<ChartData>();

            var maxLength = tempData.Select(x => x.Phase).Distinct().ToList().Select(x => x.Length).Concat(new[] { 0 }).Max();
            var sb = new StringBuilder();
            for (var i = 0; i < maxLength; i++)
                sb.Append("..");

            foreach (var project in tempData.Select(x => x.Project).Distinct().ToList())
            {
                var projectData = tempData.Where(x => x.Project == project).ToList();
                if (projectData.Count <= 0)
                    continue;

                newdata.Add(new ChartData()
                {
                    StartDate = date.AddDays(-1),
                    EndDate = date.AddDays(-1).AddHours(4),
                    Phase = project + sb + "....",
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
                "Vendor",
                "Tooltip"
                };

            //foreach (var project in newdata.Select(x => x.Project).Distinct())
            //{
            //    var tempPhases = newdata.Where(x => x.Project == project).Select(x => x.Phase).Distinct().ToList();
            //    foreach (var tempPhase in tempPhases)
            //    {
            //        var phase1 = tempPhase;
            //        var dictDuplicateTasks = newdata.Where(x => x.Project == project && x.Phase == phase1).Select(x => x.Task).GroupBy(x => x)
            //            .Where(group => group.Count() > 1)
            //            .ToDictionary(group => group.Key, x => x.ToList());
            //        foreach (var key in dictDuplicateTasks.Keys)
            //        {
            //            var duplicateTasks = newdata.Where(x => x.Project == project && x.Phase == phase1 && x.Task == key).ToList();
            //            for (var i = 0; i < duplicateTasks.Count; i++)
            //            {
            //                if (i == 0)
            //                    continue;
            //                duplicateTasks[i].Task = duplicateTasks[i].Task + new string('*', i);
            //            }
            //        }
            //    }
            //}

            foreach (var data in newdata)
            {
                var stDate = data.StartDate;
                var enDate = data.EndDate;
                var duration = data.EndDate - data.StartDate;
                var tooltip = "";
                if (stDate.Date != enDate.Date || enDate.Subtract(stDate).Hours > 4)
                {
                    if (startDate != DateTime.MinValue && stDate < startDate)
                        stDate = startDate;

                    if (endDate != DateTime.MinValue && enDate > endDate)
                        enDate = endDate;

                    tooltip = string.Format("<div ><span style='width:300px; white-space: nowrap;'><br/>&nbsp;&nbsp;<b>{0}</b><br/><br/><hr style='border-style: inset;  color: #fff; background-color: #fff;' />&nbsp;&nbsp;<b>Date Range : </b>{1}&nbsp; to &nbsp;{2}&nbsp;&nbsp;<br/>&nbsp;&nbsp;<b>Duration : </b>{3}&nbsp;&nbsp;<br /><br/></span>",
                        data.Task, data.StartDate.ToString("MM/dd/yy"), data.EndDate.ToString("MM/dd/yy"), duration.ToString("dd") + " days");

                }

                j++;


                chartData[j] = new object[] { data.Project, data.Phase, data.Task, stDate, enDate, data.Fleet, data.Color, data.Vendor, tooltip };
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

        public static DataTable ConvertExcelToDataTable(string fileName, string sheetName)
        {
            using (var objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                var cmd = new OleDbCommand();
                var dtable = new DataTable();
                var dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    var totalSheet = dt.Rows.Count; //No of sheets on excel file  
                    for (var i = 0; i < totalSheet; i++)
                    {
                        var name = dt.Rows[i]["TABLE_NAME"].ToString();

                        if (name.Contains(sheetName))
                        {
                            sheetName = name;
                            break;
                        }
                    }
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                var oleda = new OleDbDataAdapter(cmd);
                oleda.FillSchema(dtable, SchemaType.Source);

                foreach (DataColumn cl in dtable.Columns)
                {
                    if (cl.DataType == typeof(double))
                        cl.DataType = typeof(decimal);
                    if (cl.ColumnName.ToUpper().Contains("AMT"))
                        cl.DataType = typeof(decimal);
                    if (cl.ColumnName.ToUpper().Contains("DATE"))
                        cl.DataType = typeof(DateTime);
                }

                foreach (DataRow dr in dtable.Rows)
                {
                    foreach (DataColumn dc in dtable.Columns)
                    {
                        if (dc.ColumnName.ToUpper().Contains("DATE"))
                        {
                            dr[dc.ColumnName] = DateTime.ParseExact(dr[dc.ColumnName].ToString(), "dd-MM-yyyy", new CultureInfo(ConfigurationManager.AppSettings["CultureInfo"].ToString()));
                        }
                    }
                }

                oleda.Fill(dtable);
                objConn.Close();
                return dtable;
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

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public static object GenerateLegends(string val, string sDate, string eDate, string fleet, string phase, string task, string vendor)
        {
            var html = "";
            var i = 0;

            var newdata = _colorData.Select(x => x.Task).Distinct();
            html = html + "<table class='' style='border-collapse: collapse;'><tbody>";
            html = html + "<tr>";
            i = 1;
            const int noOfCols = 5;
            foreach (var data in newdata)
            {
                if (string.IsNullOrWhiteSpace(data))
                    continue;

                if (i <= noOfCols)
                {
                    var etask = data;
                    var color = _colorData.Where(x => x.Task == etask).Select(x => x.Color).FirstOrDefault();
                    html = html + "<td style='border-top: none !important;padding-right:25px;'><span style='display:inline-block;white-space: nowrap;text-overflow: ellipsis;' title='" + etask +
                           "'><svg width='15' height='15'><rect  width='15' height='15' style='fill:" + color +
                           "' /></svg> " + etask + "&nbsp;&nbsp;</span> </td>";
                    i = i + 1;
                }
                else
                {
                    var etask = data;
                    var color = _colorData.Where(x => x.Task == etask).Select(x => x.Color).FirstOrDefault();
                    html = html + "<td style='border-top: none !important;padding-right:25px;'><span style='display:inline-block;white-space: nowrap;text-overflow: ellipsis;' title='" + etask +
                           "'><svg width='15' height='15'><rect  width='15' height='15' style='fill:" + color +
                           "' /></svg> " + etask + "&nbsp;&nbsp;</span> </td>";

                    html = html + "</tr><tr>";
                    i = 1;
                }


            }
            if (i < noOfCols)
            {
                for (var j = i; j == noOfCols; j++)
                {
                    html = html + "<td style='border-top: none !important;'></td>";
                }
            }
            html = html + "</tr></tbody>";
            html = html + "</table>";

            html = html + "";
            return html;
        }
    }
}