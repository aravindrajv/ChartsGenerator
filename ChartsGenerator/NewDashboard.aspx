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
    <%--<script type="text/javascript" src="https://www.google.com/jsapi"></script>--%>
    <script type="text/javascript" src="https://www.google.com/jsapi?autoload={'modules':[{'name':'visualization','version':'1.1','packages':['timeline']}]}"></script>
    <style>
        .charts {
            overflow-y: hidden;
            /*min-height: 400px;*/
        }

        .FilterTable td {
            text-align: center;
            vertical-align: middle;
            padding: 2px;
            width: 100px;
        }
    </style>
    <script type="text/javascript">
        function onLoad() {
            $("#Charts").html('');
            var sDate = $('#strtDate').val();
            var eDate = $('#endDate').val();

            if (sDate.empty || eDate.empty) {
                sDate = "";
                eDate = "";
            }

            var pData = "";
            $.ajax({
                url: "NewDashboard.aspx/GetProjectCount",
                data: '{"sDate":"' + sDate + '", "eDate":"' + eDate + '"}',
                dataType: "json",
                type: "POST",
                contentType: "application/json; chartset=utf-8",
                success: function (json) {
                    pData = json.d;
                    CreateChart("chart", sDate, eDate);

                },
                error: function () {
                    alert("Error loading data! Please try again.");
                }
            }).done(function () {

            });
        }

        function CreateChart(val, sDate, eDate) {
            var cData;
            $.ajax({
                url: "NewDashboard.aspx/GetChartData",
                //data: "",
                data: '{"name":"' + val + '", "s_Date":"' + sDate + '", "e_Date":"' + eDate + '"}',
                dataType: "json",
                type: "POST",
                contentType: "application/json; chartset=utf-8",
                success: function (json) {
                    cData = json.d;
                    AddData(cData, val);
                },
                error: function () {
                    alert("Error loading data! Please try again.");
                }
            }).done(function () {

            });

        }


        function AddData(cData, val) {

            var divName = val.replace(" ", "");
            $("#Charts").append("<br /><div id =" + divName + " class='charts' ></div><br />");

            var options = {
                title: divName,
                curveType: 'function',
                legend: { position: 'bottom' },
                height:800
            };

            var container = document.getElementById(divName);
            var chart = new google.visualization.Timeline(container);
            var dataTable = new google.visualization.DataTable();
            //dataTable.addColumn({ type: 'string', id: 'Project' });
            dataTable.addColumn({ type: 'string', id: 'Phase' });
            dataTable.addColumn({ type: 'string', id: 'Task' });
            dataTable.addColumn({ type: 'date', id: 'Start Date' });
            dataTable.addColumn({ type: 'date', id: 'End Date' });

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

                    dataTable.addRows([
                  [val[1], val[2], new Date(syear, smonth, sday), new Date(eyear, emonth, eday)]]);
                }
            });
            chart.draw(dataTable, options);

        }

        $(document).ready(function () {
            onLoad();
        });

    </script>

    <%--Filter--%>
    <script>
        $(document).ready(function () {
            var nowTemp = new Date();
            var now = new Date(nowTemp.getFullYear(), nowTemp.getMonth(), nowTemp.getDate(), 0, 0, 0, 0);
            var checkout = $('#strtDate').datepicker({
                onRender: function (date) {
                    return date.valueOf() > now.valueOf() ? 'disabled' : '';
                }
            }).on('changeDate', function (ev) {
                checkout.hide();
            }).data('datepicker');

            checkout = $('#endDate').datepicker({
                onRender: function (date) {
                    return date.valueOf() > now.valueOf() ? 'disabled' : '';
                }
            }).on('changeDate', function (ev) {
                checkout.hide();
            }).data('datepicker');
        });


    </script>
</head>
<body>
    <form runat="server">
        <div class="container ">
            <div class="navbar navbar-default" style="text-align: center;">
                <div>
                    &nbsp;Dashboard
                </div>
            </div>
            <a href="Home.aspx">Back</a>
            <%--<div>
            <asp:GridView ID="grvExcelData" runat="server" OnPageIndexChanging="PageIndexChanging" AllowPaging="true" Width="100%" Style="text-align: left; border-color: gray;">
                    <HeaderStyle BackColor="#158CBA" Font-Bold="true" ForeColor="White" />
                </asp:GridView>
            </div>
        <br/>--%>
            <div>
                <table class="FilterTable">
                    <tr>
                        <td>
                            <span style="font-weight: bold;">Start Date :</span>
                        </td>
                        <td>
                            <input id="strtDate" type="text" class="datepicker form-control" />
                        </td>
                        <td>
                            <span style="font-weight: bold;">End Date : </span>
                        </td>
                        <td>
                            <input id="endDate" type="text" class="datepicker form-control" />
                        </td>
                        <td>
                            <input id="BtnSubmit" type="button" class="btn-primary" value="CreateChart" onclick="onLoad();" />
                        </td>
                    </tr>

                </table>
            </div>

            <div>
                <div id="Charts"></div>
            </div>
            
        </div>
    </form>
</body>
</html>
