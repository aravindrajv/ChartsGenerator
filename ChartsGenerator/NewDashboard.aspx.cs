﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.Script.Services;
using System.Web.Services;
using ChartsGenerator.Model;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ChartsGenerator
{
    public partial class NewDashboard : System.Web.UI.Page
    {
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
                    try
                    {
                        var filepath = HttpContext.Current.Session["FPath"].ToString();
                        var _cData = ConvertExcelToDataTable(filepath, "Data");

                        var colorDataTable = ConvertExcelToDataTable(filepath, "ColorCodes");
                        var colors = (from DataRow row in colorDataTable.Rows
                                      select new ColorData()
                                      {
                                          Color = row["Color"].ToString(),
                                          Task = row["Task"].ToString(),
                                      }).ToList();

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

                            var chartData = new ChartData
                            {
                                Project = row["Project"].ToString(),
                                Phase = row["Phase"].ToString(),
                                Task = row["Task"].ToString(),
                                //Task = "",
                                StartDate = stDate,
                                EndDate = eDate,
                                Fleet = row["Fleet"].ToString(),
                                Vendor = row["Vendor"].ToString(),
                                Color = row["Color"].ToString()
                            };
                            var color = colors.FirstOrDefault(x => x.Task == chartData.Task);
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
                    StartDate = date,
                    EndDate = date.AddHours(4),
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

        private DataTable ConvertExcelToDataTable(string fileName, string sheetName)
        {
            DataTable dt = null;
            try
            {
                object rowIndex = 1;
                dt = new DataTable();
                DataRow row;
                var app = new Application();
                var workBook = app.Workbooks.Open(fileName, 0, true, 5, "", "", true,
                    XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                var workSheet = (Worksheet)app.Worksheets[sheetName];
                int temp = 1;
                while (((Range)workSheet.Cells[rowIndex, temp]).Value2 != null)
                {
                    dt.Columns.Add(Convert.ToString(((Range)workSheet.Cells[rowIndex, temp]).Value2));
                    temp++;
                }
                rowIndex = Convert.ToInt32(rowIndex) + 1;
                int columnCount = temp;
                temp = 1;
                while (((Range)workSheet.Cells[rowIndex, temp]).Value2 != null)
                {
                    row = dt.NewRow();
                    for (int i = 1; i < columnCount; i++)
                    {
                        if (i == 6 || i == 7)
                        {

                            DateTime datetime = DateTime.Parse(ConvertToDateTime(Convert.ToString(((Range)workSheet.Cells[rowIndex, i]).Value2)));
                            var date = datetime.Date;
                            row[i - 1] = date;
                        }
                        else
                        {
                            row[i - 1] = Convert.ToString(((Range)workSheet.Cells[rowIndex, i]).Value2);
                        }
                    }
                    dt.Rows.Add(row);
                    rowIndex = Convert.ToInt32(rowIndex) + 1;
                    temp = 1;
                }
                app.Workbooks.Close();
            }
            catch (Exception ex)
            {
                //lblError.Text = ex.Message;
            }
            return dt;
        }

        public static string ConvertToDateTime(string strExcelDate)
        {
            double excelDate;
            try
            {
                excelDate = Convert.ToDouble(strExcelDate);
            }
            catch
            {
                return strExcelDate;
            }
            if (excelDate < 1)
            {
                throw new ArgumentException("Excel dates cannot be smaller than 0.");
            }
            DateTime dateOfReference = new DateTime(1900, 1, 1);
            if (excelDate > 60d)
            {
                excelDate = excelDate - 2;
            }
            else
            {
                excelDate = excelDate - 1;
            }
            return dateOfReference.AddDays(excelDate).ToShortDateString();
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
                && fleets.Length == 0 && phases.Length == 0 && tasks.Length == 0 && vendors.Length == 0)
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

            var html = "<div >";
            var i = 0;
            if (val == "")
            {
                var newdata = tempData.Select(x => x.Task).Distinct();
                html = html + "<table>";
                html = html + "<tr>";
                i = 1;
                foreach (var data in newdata)
                {
                    if (i <= 6)
                    {
                        var etask = data;
                        var color = tempData.Where(x => x.Task == etask).Select(x => x.Color).FirstOrDefault();
                        html = html + "<td><span style='display:inline-block;' title='" + etask +
                               "'><svg width='15' height='15'><rect  width='15' height='15' style='fill:" + color +
                               "' /></svg> " + etask + " </span> <span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>";
                        i = i + 1;
                    }
                    else
                    {
                        html = html + "</tr><tr>";
                        i = 1;
                    }
                    

                }
                if (i < 6)
                {
                    for (var j = i; j == 6; j++)
                    {
                        html = html + "<td></td>";
                    }
                }
                html = html + "</tr>";
                html = html + "</table>";
            }
            else
            {
                var selectedval = val.Split(',');
                var newdata = tempData.Where(x => !selectedval.Contains(x.Task)).Select(x => x.Task).Distinct();
                html = html + "<table>";
                html = html + "<tr>";
                i = 1;
                foreach (var data in newdata)
                {
                    if (i <= 6)
                    {
                        var etask = data;
                        var color = tempData.Where(x => x.Task == etask).Select(x => x.Color).FirstOrDefault();
                        html = html + "<td><span style='display:inline-block;' title='" + etask +
                               "'><svg width='15' height='15'><rect  width='15' height='15' style='fill:" + color +
                               "' /></svg> " + etask + " </span> <span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>";
                        i = i + 1;
                    }
                    else
                    {
                        html = html + "</tr><tr>";
                        i = 1;
                    }
                    
                }

                if (i < 6)
                {
                    for (var j = i; j == 6; j++)
                    {
                        html = html + "<td></td>";
                    }
                }
                html = html + "</tr>";
                html = html + "</table>";


            }
            html = html + "</div>";
            return html;
        }
    }
}