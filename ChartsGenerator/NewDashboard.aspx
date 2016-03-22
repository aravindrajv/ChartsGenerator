<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="NewDashboard.aspx.cs" Inherits="ChartsGenerator.NewDashboard" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Dashboard</title>
    <script src="Scripts/jquery-2.2.0.min.js"></script>
    <script src="Scripts/bootstrap.js"></script>
    <link href="Content/bootstrap.css" rel="stylesheet" />
    <link href="Content/font-awesome.css" rel="stylesheet" />
    <link href="Content/site.css" rel="stylesheet" />
    <link href="Content/bootstrap-datepicker.css" rel="stylesheet" />
    <script src="Scripts/bootstrap-datepicker.js"></script>
    <%--<script type="text/javascript" src="https://www.google.com/jsapi?autoload={'modules':[{'name':'visualization','version':'1.1','packages':['timeline']}]}"></script>--%>
    <script type="text/javascript" src="Scripts/loader.js"></script>
    <style>
          .charts {
             overflow-y: hidden;
             /*min-height: 400px;*/
         }

        .fields {
            width: 130px;
            height: 25px;
        }

        .ListBoxCssClass {
            width: 150px;
        }

        .btncustom {
            width: 100px;
        }
    </style>
    <script type="text/javascript">
        var cData;
        var valnew;
        google.charts.load('current', { 'packages': ['timeline'] });

        function onLoad() {
            cData = null;
            valnew = null;
            $("#Charts").html('');
            var sDate = $('#strtDate').val();
            var eDate = $('#endDate').val();
            //var selected = $("[id*=ListBox1] option:selected");
            var fleet = $('#lstFleet').val();
            var phase = $('#lstPhase').val();
            var vendor = $('#lstVendor').val();
            var task = $('#lstTasks').val();

            if (sDate.empty || eDate.empty) {
                sDate = "";
                eDate = "";
            }

            var pData = "";
            $.ajax({
                url: "NewDashboard.aspx/GetProjectCount",
                data: '{"sDate":"' + sDate + '", "eDate":"' + eDate + '","fleet":"' + fleet + '", "phase":"' + phase + '", "task":"' + task + '", "vendor":"' + vendor + '"}',
                dataType: "json",
                type: "POST",
                contentType: "application/json; chartset=utf-8",
                success: function (json) {
                    pData = json.d;
                    CreateChart("chart", sDate, eDate, phase, fleet, task, vendor);

                },
                error: function () {
                    alert("Error loading data! Please try again.");
                }
            }).done(function () {

            });
            GenerateLegends();
        }

        function clearfilters() {
            $("#Charts").html('');

            $('#strtDate').val("");
            $('#endDate').val("");
            $("#lstFleet").val([]);
            $("#lstPhase").val([]);
            $("#lstVendor").val([]);
            $("#lstTasks").val([]);


            var sDate = $('#strtDate').val();
            var eDate = $('#endDate').val();
            //var selected = $("[id*=ListBox1] option:selected");
            var fleet = $('#lstFleet').val();
            var phase = $('#lstPhase').val();
            var vendor = $('#lstVendor').val();
            var task = $('#lstTasks').val();

            if (sDate.empty || eDate.empty) {
                sDate = "";
                eDate = "";
            }

            var pData = "";
            $.ajax({
                url: "NewDashboard.aspx/GetProjectCount",
                data: '{"sDate":"' + sDate + '", "eDate":"' + eDate + '","fleet":"' + fleet + '", "phase":"' + phase + '", "task":"' + task + '", "vendor":"' + vendor + '"}',
                dataType: "json",
                type: "POST",
                contentType: "application/json; chartset=utf-8",
                success: function (json) {
                    pData = json.d;
                    CreateChart("chart", sDate, eDate, phase, fleet, task, vendor);

                },
                error: function () {
                    alert("Error loading data! Please try again.");
                }
            }).done(function () {

            });
            GenerateLegends();
        }

        function GenerateLegends() {

            var sDate = $('#strtDate').val();
            var eDate = $('#endDate').val();
            //var selected = $("[id*=ListBox1] option:selected");
            var fleet = $('#lstFleet').val();
            var phase = $('#lstPhase').val();
            var vendor = $('#lstVendor').val();
            var task = $('#lstTasks').val();

            var selectedVal = [];
            $('#lstTasks :selected').each(function (i, selected) {
                selectedVal[i] = $(selected).text();
                //alert(selectedVal[i]);
            });

            $("#LegendsDiv").html('');
            $.ajax({
                url: "NewDashboard.aspx/GenerateLegends",
                data: '{"val":"' + selectedVal + '", "sDate":"' + sDate + '", "eDate":"' + eDate + '","fleet":"' + fleet + '", "phase":"' + phase + '", "task":"' + task + '", "vendor":"' + vendor + '"}',
                dataType: "json",
                type: "POST",
                contentType: "application/json; chartset=utf-8",
                success: function (json) {
                    $("#LegendsDiv").html(json.d);
                    //  CreateChart("chart", sDate, eDate, phase, fleet, task, vendor);

                },
                error: function () {
                    alert("Error loading data! Please try again.");
                }
            }).done(function () {

            });
        }


        function CreateChart(val, sDate, eDate, phase, fleet, task, vendor) {
            valnew = val;
            $.ajax({
                url: "NewDashboard.aspx/GetChartData",
                //data: "",
                data: '{"sDate":"' + sDate + '", "eDate":"' + eDate + '","fleet":"' + fleet + '", "phase":"' + phase + '", "task":"' + task + '", "vendor":"' + vendor + '"}',
                dataType: "json",
                type: "POST",
                contentType: "application/json; chartset=utf-8",
                success: function (json) {
                    cData = json.d;
                    //AddData(cData, val);
                    google.charts.setOnLoadCallback(AddData);
                },
                error: function () {
                    alert("Error loading data! Please try again.");
                }
            }).done(function () {

            });
        }


        function AddData() {
            var val = valnew;

            var divName = val.replace(" ", "");
            $("#Charts").append("<br /><div id =" + divName + " class='charts' ></div><br />");
            
            var container = document.getElementById(divName);
            var chart = new google.visualization.Timeline(container);
            var dataTable = new google.visualization.DataTable();
            //dataTable.addColumn({ type: 'string', id: 'Project' });
            dataTable.addColumn({ type: 'string', id: 'Phase' });
            dataTable.addColumn({ type: 'string', id: 'Task' });
            dataTable.addColumn({ 'type': 'string', 'role': 'tooltip', 'p': { 'html': true } });
            dataTable.addColumn({ type: 'date', id: 'Start Date' });
            dataTable.addColumn({ type: 'date', id: 'End Date' });
            

            var colors = '';
            var first = true;

            jQuery.each(cData, function (i, val) {
                if (i != 0) {
                    var remove = /-?\d+/;
                    var clearsDate = remove.exec(val[3]);
                    var sDate = new Date(parseInt(clearsDate[0]));
                    var syear = sDate.getFullYear();
                    var smonth = sDate.getMonth();
                    var sday = sDate.getDate();

                    var cleareDate = remove.exec(val[4]);
                    var eDate = new Date(parseInt(cleareDate[0]));
                    var eyear = eDate.getFullYear();
                    var emonth = eDate.getMonth();
                    var eday = eDate.getDate();
                    dataTable.addRows([[val[5] + ' ' + val[1], val[2], val[8], new Date(syear, smonth, sday), new Date(eyear, emonth, eday)]]);
                    if (first) {
                        colors = val[6];
                        first = false;
                    } else
                        colors = colors + ',' + val[6];
                }
            });


            var options = {
                title: divName,
                curveType: 'function',
                height: 800,
                colors: colors.split(','),
                timeline: { showBarLabels: false },
                //tooltip : {isHtml:true}
            };

            chart.draw(dataTable, options);

        }

        $(document).ready(function () {
            onLoad();
        });

    </script>

    <%--Filter--%>
    <script>
        $(document).ready(function () {
            strtDate();
            endDate();
        });

        function strtDate() {
            var nowTemp = new Date();
            var now = new Date(nowTemp.getFullYear(), nowTemp.getMonth(), nowTemp.getDate(), 0, 0, 0, 0);


            var checkout = $('#strtDate').datepicker({
                onRender: function (date) {
                    return date.valueOf() > now.valueOf() ? 'disabled' : '';
                }
            }).on('changeDate', function (ev) {
                checkout.hide();
            }).data('datepicker');
        }

        function endDate() {
            var nowTemp = new Date();
            var now = new Date(nowTemp.getFullYear(), nowTemp.getMonth(), nowTemp.getDate(), 0, 0, 0, 0);


            var checkout = $('#endDate').datepicker({
                onRender: function (date) {
                    return date.valueOf() > now.valueOf() ? 'disabled' : '';
                }
            }).on('changeDate', function (ev) {
                checkout.hide();
            }).data('datepicker');

        }


    </script>
</head>
<body>
    <form runat="server">
        <div class="container ">
            <div class="navbar navbar-default" >
                <table style="width: 100%">
                    <tr>
                        <td style="width: 30px; font-weight: normal; font-size: 13px;"><a href="Home.aspx">Back</a></td>
                        <td style="width: 1110px; text-align: center;">Dashboard</td>
                    </tr>
                </table>
            </div>
            <%--<a href="Home.aspx">Back</a>--%>

            <%--<div>
            <asp:GridView ID="grvExcelData" runat="server" OnPageIndexChanging="PageIndexChanging" AllowPaging="true" Width="100%" Style="text-align: left; border-color: gray;">
                    <HeaderStyle BackColor="#158CBA" Font-Bold="true" ForeColor="White" />
                </asp:GridView>
            </div>
        <br/>--%>

            
            <div >
                <div class="table-responsive" style="border-color: #E6E6E6;background-color: #E6E6E6;border-radius: 10px;">
                    <table class="table" style="border-collapse: collapse;">
                      <tbody>
                        <tr>
                        <td style="" >
                            <br />
                            <span style="font-weight: bold;white-space: nowrap;">Start Date :</span>
                            <br />
                            <br />
                            <span style="font-weight: bold;white-space: nowrap;">End Date &nbsp;&nbsp;:</span>
                        </td>
                        <td >
                            <br />
                            <input class="fields" id="strtDate" type="text" class="datepicker " />
                            <br />
                            <br />
                            <input class="fields" id="endDate" type="text" class="datepicker " />
                        </td>
                        <td >
                            <span style="font-weight: bold;">Vendor : </span>
                            <br />
                            <asp:ListBox CssClass="ListBoxCssClass"  ID="lstVendor" runat="server" ClientIDMode="Static" SelectionMode="Multiple" />
                        </td>
                        <td >
                            <span style="font-weight: bold;">Fleet : </span>
                            <br />
                            <asp:ListBox  CssClass="ListBoxCssClass" ID="lstFleet" runat="server" ClientIDMode="Static" SelectionMode="Multiple" />
                        </td>
                        <td >
                            <span style="font-weight: bold;">Release : </span>
                            <br />
                            <asp:ListBox  CssClass="ListBoxCssClass" ID="lstPhase" runat="server" ClientIDMode="Static" SelectionMode="Multiple" />
                        </td>
                        <td >
                            <span style="font-weight: bold;">Exclude Task : </span>
                            <br />
                            <asp:ListBox  CssClass="ListBoxCssClass" ID="lstTasks" runat="server" ClientIDMode="Static" SelectionMode="Multiple" />
                        </td>
                        <td >
                            <br />
                            <input id="BtnSubmit" type="button" class="btn-primary btncustom" value="Create Chart" onclick="onLoad();" />
                            <br />
                            <br />
                            <input id="ClearBtn" type="button" class="btn-primary btncustom" value="Clear Filters" onclick="clearfilters();" />
                        </td>
                    </tr>
                    </tbody>
                </table>
                </div>
            </div>
            <br />
            <div >
                <div id="LegendsDiv"   class="table-responsive" style="border-color: #E6E6E6;background-color: #E6E6E6;border-radius: 10px;padding-top: 5px;padding-bottom: 5px;padding-left: 5px;padding-right: 5px;">
                </div>
            </div>
            <div>
                <div id="Charts"></div>
            </div>
            
            <div>
                <asp:Label Style="color: red;" ID="Error" runat="server"></asp:Label>
            </div>
            <br />

        </div>
    </form>
    </body>
</html>
