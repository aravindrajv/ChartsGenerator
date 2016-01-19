<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Dashboard.aspx.cs" Inherits="ChartsGenerator.Dashboard" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Dashboard</title>
    <script src="Scripts/jquery-2.2.0.min.js"></script>
    <script src="Scripts/bootstrap.js"></script>
    <link href="Content/bootstrap.css" rel="stylesheet" />
    <link href="Content/font-awesome.css" rel="stylesheet" />
    <link href="Content/site.css" rel="stylesheet" />
    <%--<script type="text/javascript" src="https://www.google.com/jsapi"></script>--%>
    <script type="text/javascript" src="https://www.google.com/jsapi?autoload={'modules':[{'name':'visualization','version':'1.1','packages':['timeline']}]}"></script>
    <style>
        .charts {
            overflow-y: hidden;
            /*min-height: 400px;*/
        }
    </style>
    <script type="text/javascript">
        function onLoad() {
            var pData= "";
            $.ajax({
                url: "Dashboard.aspx/GetProjectCount",
                data: "",
                dataType: "json",
                type: "POST",
                contentType: "application/json; chartset=utf-8",
                success: function (json) {
                    pData = json.d;
                    jQuery.each(pData, function (i, val) {
                        CreateChart(val);
                    });
                },
                error: function () {
                    alert("Error loading data! Please try again.");
                }
            }).done(function () {

            });
        }

        function CreateChart(val) {
            var cData;
            $.ajax({
                url: "Dashboard.aspx/GetChartData",
                //data: "",
                data: '{"name":"' + val + '"}',
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
            $("#Charts").append("<h3 align='center'>" + val + "</h3><div id =" + divName + " class='charts' ></div><br />");

            var options = {
                title: divName,
                curveType: 'function',
                legend: { position: 'bottom' }
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
</head>
<body>
    <div class="container ">
        <div class="navbar navbar-default" style="text-align: center;">
            <div>
                &nbsp;Dashboard
            </div>
        </div>

        <div>
            <div id="Charts"></div>
        </div>
        <a href="Home.aspx">Back</a>
    </div>
</body>
</html>
