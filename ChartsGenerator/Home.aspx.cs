using System;
using System.IO;

namespace ChartsGenerator
{
    public partial class Home : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Session["FPath"] = null;
            }
        }

        protected void UploadBtn_Click(object sender, EventArgs e)
        {
            //var fileName = Path.GetFileName(FileUploadXL.PostedFile.FileName);
            if (File.Exists(Server.MapPath(Path.Combine("~/input/", "template.xlsx"))))
            {
                File.Delete(Server.MapPath(Path.Combine("~/input/", "template.xlsx")));
            }
            FileUploadXL.PostedFile.SaveAs(Server.MapPath("~/input/") + "template.xlsx");
            Session["FPath"] = (Server.MapPath(Path.Combine("~/input/", "template.xlsx")));
            Response.Redirect("Dashboard.aspx");
        }
    }
}