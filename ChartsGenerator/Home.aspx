<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Home.aspx.cs" Inherits="ChartsGenerator.Home" %>


<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Home</title>
    <script src="Scripts/jquery-2.2.0.min.js"></script>
    <script src="Scripts/bootstrap.js"></script>
    <link href="Content/bootstrap.css" rel="stylesheet" />
    <link href="Content/font-awesome.css" rel="stylesheet" />
    <link href="Content/site.css" rel="stylesheet" />

</head>
<body>
    <div class="container ">
        <div class="navbar navbar-default" style="text-align: center;">
            <div>
                &nbsp;Home
            </div>
        </div>
        <form runat="server">
            <table align="center">
                <tr>
                    <td colspan="2" align="center">
                        <a href="sample/template.xlsx">Sample File</a>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <br/>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <span><b>Please Upload a file</b></span>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:FileUpload ID="FileUploadXL" runat="server" />
                        <asp:RequiredFieldValidator ErrorMessage="Required" ControlToValidate="FileUploadXL"
                            runat="server" Display="Dynamic" ForeColor="Red" />
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator1" ValidationExpression="([a-zA-Z0-9\s_\\.\-:])+(.xlsx)$"
                            ControlToValidate="FileUploadXL" runat="server" ForeColor="Red" ErrorMessage="Please select a valid xlsx file."
                            Display="Dynamic" />
                    </td>
                    <td>
                        <asp:Button ID="UploadBtn" runat="server" Text="Button" OnClick="UploadBtn_Click" />
                    </td>
                </tr>
            </table>
        </form>
    </div>
</body>
</html>
