using System;
using System.IO;
using System.Text;

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
            //if (File.Exists(Server.MapPath(Path.Combine("~/input/", "template.xlsx"))))
            //{
            //    File.Delete(Server.MapPath(Path.Combine("~/input/", "template.xlsx")));
            //}
            var fName = RandomHexString(5)+ ".xlsx";
            FileUploadXL.PostedFile.SaveAs(Server.MapPath("~/input/") + fName);
            Session["FPath"] = (Server.MapPath(Path.Combine("~/input/", fName)));
            Response.Redirect("Dashboard.aspx");
        }

        public string RandomHexString(int length)
        {
            var mySb = new StringBuilder();
            var myRandom = new Random((int)DateTime.Now.Ticks);
            while (mySb.Length < length)
            {
                int nextValue = (int)(myRandom.NextDouble() * 100000);
                string doubleString = nextValue.ToString("X");
                int dsLength = doubleString.Length;
                if (dsLength > 4)
                    dsLength = 4;

                mySb.Append(doubleString.Substring(doubleString.Length - dsLength, dsLength));
            }
            return mySb.ToString();
        }
    }
}