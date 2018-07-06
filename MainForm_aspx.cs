using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Web.UI.Design.WebControls;
using System.DirectoryServices;
using System.Security.Principal;
using System.Collections;
using System.Data.OleDb;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Windows.Forms;

namespace ********
{
    public partial class MainForm : System.Web.UI.Page
    {
        DAL dal = new DAL();//класс, где реализованы процедуры и функции для работы с бд

        public int idxCurr, idxPred, idxCurr_Pr;
        
        
        private void UpdateGrid()
        {
            string U = Context.User.Identity.Name;
            Title += String.Format(" ({0})", U.Substring(U.IndexOf(@"\") + 1));

            string p = DivisionDropDownList.SelectedValue;

            if (p != "")
            {
                if (dal.GetLastUser(p) != "")
                {
                    ClientGrid.EmptyDataText = "Оценки выставлены пользователем " + dal.GetLastUser(p);
                    ProviderGrid.EmptyDataText = "Оценки выставлены пользователем " + dal.GetLastUser(p);
                }
                else
                {
                    ClientGrid.EmptyDataText = "Нет оценок";
                    ProviderGrid.EmptyDataText = "Нет оценок";
                }
            }
            else
            {
                ClientGrid.EmptyDataText = "Нет оценок";
                ProviderGrid.EmptyDataText = "Нет оценок";
            }

            ClientGrid.DataSource = dal.GetMarksC(p);
            ProviderGrid.DataSource = dal.GetMarksP(p);
            ScalaGrid.DataSource = dal.GetScala();
            
            Page.DataBind();
        }

        private void SelectData()
        {
            string U = Context.User.Identity.Name;
            string UU = U.Substring(U.IndexOf(@"\") + 1);
            int t = 0;

            if (ClientGrid.Rows.Count != 0)
                for (int i = 0; i < ClientGrid.Rows.Count; i++)
                {
                    string DU_ID = DivisionDropDownList.SelectedValue;
                    string DM_ID = ClientGrid.DataKeys[i].Values["Division_ID"].ToString();
                    string P_ID = ClientGrid.DataKeys[i].Values["Product_ID"].ToString();
                    string PER_ID = ClientGrid.DataKeys[i].Values["Period_ID"].ToString();

                    DropDownList ddl = ClientGrid.Rows[i].FindControl("DropDownMarks") as DropDownList;
                    string M = ddl.SelectedValue;
                    string C = ClientGrid.Rows[i].Cells[4].Text;

                    dal.InsertMark(UU, DU_ID, DM_ID, P_ID, PER_ID, M, C, 1);

                    if (t == 0)
                    {
                        dal.InsertUserLog(DU_ID, UU, PER_ID);
                        t = 1;
                    }
                }

            if (ProviderGrid.Rows.Count != 0)
                for (int i = 0; i < ProviderGrid.Rows.Count; i++)
                {
                    string DU_ID = DivisionDropDownList.SelectedValue;
                    string DM_ID = ProviderGrid.DataKeys[i].Values["Division_ID"].ToString();
                    string P_ID = ProviderGrid.DataKeys[i].Values["Product_ID"].ToString();
                    string PER_ID = ProviderGrid.DataKeys[i].Values["Period_ID"].ToString();

                    DropDownList ddl = ProviderGrid.Rows[i].FindControl("DropDownMarks") as DropDownList;
                    string M = ddl.SelectedValue;
                    string C = ProviderGrid.Rows[i].Cells[4].Text;

                    dal.InsertMark(UU, DU_ID, DM_ID, P_ID, PER_ID, M, C, 2);

                    if (t == 0)
                    {
                        dal.InsertUserLog(DU_ID, UU, PER_ID);
                        t = 1;
                    }
                }

            UpdateGrid();
        }

        protected void ClientGridView_OnRowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType.Equals(DataControlRowType.DataRow))
            {
                foreach (DataControlFieldCell cell in e.Row.Cells)
                {
                    foreach (Control control in cell.Controls)
                    {
                        DropDownList ddl = control as DropDownList;
                        if (ddl != null)
                            ddl.Attributes.Add("onclick", "SetScrollEvent()");
                    }
                }
            }
        }

        protected void bOK_Click(object sender, EventArgs e)
        {
            if ((PanelComment.Visible == true) && (Comment.Text == String.Empty))
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Пожалуйста, оставьте комментарий');", true);
            }
            else
                if ((PanelComment.Visible == true) && (Comment.Text.Length < 20) && (Comment.Text != String.Empty))
                {
                    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Пожалуйста, оставьте более развернутый комментарий');", true);
                }
                else
                {
                    if ((Session["grView"] as GridView).ID == "ClientGrid")
                    {
                        ClientGrid.Rows[Convert.ToInt16(Session["Row"])].Cells[4].Text = Comment.Text;
                        PanelComment.Visible = false;
                        Comment.Text = String.Empty;
                        ClientGrid.Enabled = true;
                        ProviderGrid.Enabled = true;
                        SaveButton.Enabled = true;
                    }
                    else
                    {
                        ProviderGrid.Rows[Convert.ToInt16(Session["Row"])].Cells[4].Text = Comment.Text;
                        PanelComment.Visible = false;
                        Comment.Text = String.Empty;
                        ClientGrid.Enabled = true;
                        ProviderGrid.Enabled = true;
                        SaveButton.Enabled = true;
                    }
                }
         }

        //комментарий при выборе ddl меньше 6
        protected void ddlCompany_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (IsPostBack)
            {
                DropDownList ddl = (DropDownList)sender;
                GridViewRow row = (GridViewRow)ddl.Parent.Parent;
                GridView gv = (GridView)ddl.Parent.Parent.Parent.Parent;
                Session["ddl"] = ddl;
                Session["grView"] = gv;
                Session["Row"] = row.RowIndex;

                int res = Convert.ToInt32(ddl.SelectedValue.ToString());
                if (((res < 6) && (res > 0)) || (res > 8))
                {
                    ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Пожалуйста, оставьте комментарий');", true);
                    PanelComment.Visible = true;
                    Comment.Text = String.Empty;
                    ClientGrid.Enabled = false;
                    ProviderGrid.Enabled = false;
                    SaveButton.Enabled = false;
                }
                else
                {
                    if ((Session["grView"] as GridView).ID == "ClientGrid")
                    {
                        ClientGrid.Rows[Convert.ToInt16(Session["Row"])].Cells[4].Text = String.Empty;
                    }
                    else
                    {
                        ProviderGrid.Rows[Convert.ToInt16(Session["Row"])].Cells[4].Text = String.Empty;
                    }
                }
                
            }
         }
        
         private void IsAdmin()
        {
            int n1 = 0;
            int n2 = 0;
            int n3 = 0;
            string U = Context.User.Identity.Name;
            string UU = U.Substring(U.IndexOf(@"\") + 1);

            
            string sUserDomain = "********";
            DirectoryEntry entry = new DirectoryEntry(string.Format("LDAP://{0}", sUserDomain));
            DirectorySearcher mySearcher = new DirectorySearcher(entry);
            mySearcher.Filter = string.Format("(&(objectClass=user) (sAMAccountName= {0}))", UU);
            mySearcher.PropertiesToLoad.Add("memberOf"); SearchResult searchresult = mySearcher.FindOne();

                if (!(searchresult == null))
                {
                    foreach (string dn in searchresult.Properties["memberOf"])
                    {
                        DirectoryEntry group = new DirectoryEntry(string.Format("LDAP://{0}/{1}", sUserDomain, dn));
                        SecurityIdentifier sid = new SecurityIdentifier(group.Properties["objectSid"][0] as byte[], 0);
                        if (sid.Value == "*********") n1 = 1;
                        if (sid.Value == "*********") n2 = 1;
                        if (sid.Value == "*********") n3 = 1;
                    }

                    if (n1 == 1)
                    {
                        AdminButton.Visible = true;
                    }

                    if (n2 == 1)
                    {
                        UpdateButton.Enabled = false;
                        SaveButton.Enabled = false;
                    }

                    if (n3 == 1)
                    {
                        UpdateButton.Enabled = true;
                        SaveButton.Enabled = true;
                    }
                }
        }

        private void UserDivision()
        {
            string U = Context.User.Identity.Name;
            string UU = U.Substring(U.IndexOf(@"\") + 1);

            int dv, dp, b;

            dal.UserDivision(UU, out dv, out dp, out b);

            if (dv != 0)
            {
                BlockDropDownList.DataSource = dal.GetBlock();
                BlockDropDownList.DataBind();
                BlockDropDownList.SelectedValue = b.ToString();

                DepartmentDropDownList.DataSource = dal.GetDepartment(BlockDropDownList.SelectedValue);
                DepartmentDropDownList.DataBind();
                DepartmentDropDownList.SelectedValue = dp.ToString();

                DivisionDropDownList.DataSource = dal.GetDivision(DepartmentDropDownList.SelectedValue);
                DivisionDropDownList.DataBind();
                DivisionDropDownList.SelectedValue = dv.ToString();
            }
            else
            {
                BlockDropDownList.DataSource = dal.GetBlock();
                BlockDropDownList.DataBind();

                DepartmentDropDownList.DataSource = dal.GetDepartment(BlockDropDownList.SelectedValue);
                DepartmentDropDownList.DataBind();

                DivisionDropDownList.DataSource = dal.GetDivision(DepartmentDropDownList.SelectedValue);
                DivisionDropDownList.DataBind();
            }
        
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                idxCurr = -1;
                idxCurr_Pr = -1;
                Session["Cl_idxPred"] = -1;
                Session["Cl_ddl"] = null;
                Session["Pr_idxPred"] = -1;
                Session["Pr_ddl"] = null;
                Session["grView"] = ClientGrid;

                IsAdmin();

                UserDivision();

                UpdateGrid();
            }
        }

        protected void BlockDropDownList_SelectedIndexChanged(object sender, EventArgs e)
        {
            DepartmentDropDownList.DataSource = dal.GetDepartment(BlockDropDownList.SelectedValue);
            DepartmentDropDownList.DataBind();

            DivisionDropDownList.DataSource = dal.GetDivision(DepartmentDropDownList.SelectedValue);
            DivisionDropDownList.DataBind();

            UpdateGrid();
        }


        protected void DepartmentDropDownList_SelectedIndexChanged(object sender, EventArgs e)
        {
            DivisionDropDownList.DataSource = dal.GetDivision(DepartmentDropDownList.SelectedValue);
            DivisionDropDownList.DataBind();

            UpdateGrid();
        }

        protected void DivisionDropDownList_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateGrid();
        }

        protected void SaveButton_Click(object sender, EventArgs e)
        {
            DropDownList ddl;
            int t = 0;

            if (ClientGrid.Rows.Count != 0)
                for (int i = 0; i < ClientGrid.Rows.Count; i++)
                {
                    ddl = ClientGrid.Rows[i].FindControl("DropDownMarks") as DropDownList;
                    if (ddl.SelectedValue != "") t++;
                    else
                    {
                        ddl.Focus();
                        break;
                    }
                }

            if (ProviderGrid.Rows.Count != 0)
                for (int i = 0; i < ProviderGrid.Rows.Count; i++)
                {
                    ddl = ProviderGrid.Rows[i].FindControl("DropDownMarks") as DropDownList;
                    if (ddl.SelectedValue != "") t++;
                    else
                    {
                        ddl.Focus();
                        break;
                    }
                }

            string p = DivisionDropDownList.SelectedValue;
            string per = "0";
            int m = 0;
            if (ClientGrid.Rows.Count != 0) per = ClientGrid.DataKeys[0].Values["Period_ID"].ToString();
            else if (ProviderGrid.Rows.Count != 0) per = ProviderGrid.DataKeys[0].Values["Period_ID"].ToString();
            m = dal.IsMarked(p, per); //проверка на повторную оценку

            if (m == 0)
            {
                if (t == (ClientGrid.Rows.Count + ProviderGrid.Rows.Count))
                {
                    SelectData();
                }
                else ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Невозможно сохранить. Не все оценки заполнены.');", true);
            }
            else
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Невозможно сохранить. Оценки уже внесены');", true);
                UpdateGrid();
            }
        }

        protected void UpdateButton_Click(object sender, EventArgs e)
        {
            UpdateGrid();
        }

        protected void MarksButton_Click(object sender, EventArgs e)
        {
            Response.Redirect("*******.aspx");
        }

        protected void AdminButton_Click(object sender, EventArgs e)
        {
            Response.Redirect("********.aspx");
        }

        protected void ExcelButton_Click(object sender, EventArgs e)
        {
            string p = DivisionDropDownList.SelectedValue;
            DataTable dtc = dal.GetMarksCExcel(p);
            DataTable dtp = dal.GetMarksPExcel(p);

            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("********");
                ws.Cells[1, 1].Value = DivisionDropDownList.SelectedItem;
                ws.Cells[3, 1].LoadFromDataTable(dtc, true);
                ws.Cells[dtc.Rows.Count + 6, 1].LoadFromDataTable(dtp, true);

                ws.Column(1).Width = 50;
                ws.Column(2).Width = 50;

                ws.Cells[3, 1, dtc.Rows.Count + 6 + dtp.Rows.Count + 1, 2].Style.WrapText = true;

                ws.Row(1).Style.Font.Bold = true;
                ws.Row(3).Style.Font.Bold = true;
                ws.Row(dtc.Rows.Count + 6).Style.Font.Bold = true;


                ws.Cells[3, 1, dtc.Rows.Count + 3, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[dtc.Rows.Count + 6, 1, dtc.Rows.Count + 6 + dtp.Rows.Count, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Row(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Row(dtc.Rows.Count + 6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                ws.Cells[3, 1, dtc.Rows.Count + 3, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[3, 1, dtc.Rows.Count + 3, 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[3, 1, dtc.Rows.Count + 3, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[3, 1, dtc.Rows.Count + 3, 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws.Cells[dtc.Rows.Count + 6, 1, dtc.Rows.Count + 6 + dtp.Rows.Count, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[dtc.Rows.Count + 6, 1, dtc.Rows.Count + 6 + dtp.Rows.Count, 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[dtc.Rows.Count + 6, 1, dtc.Rows.Count + 6 + dtp.Rows.Count, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[dtc.Rows.Count + 6, 1, dtc.Rows.Count + 6 + dtp.Rows.Count, 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                Response.Clear();
                Response.AddHeader("content-disposition", "attachment;  filename=*******.xlsx");
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.BinaryWrite(pck.GetAsByteArray());
                Response.End();
            }
        }

        protected void ImportExcelButton_Click(object sender, EventArgs e)
        {
            string excelConnectionString = @"**************";
            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = excelConnectionString;
            OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
            connection.Open();
            DbDataReader dr = command.ExecuteReader();
            string sqlConnectionString = @"*************";
            SqlBulkCopy bulkInsert = new SqlBulkCopy(sqlConnectionString);
            bulkInsert.DestinationTableName = "********";
            bulkInsert.WriteToServer(dr);
            connection.Close();
        }

        protected void ClientGrid_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ClientGrid.Columns[4].Visible)
            {
                ClientGrid.Columns[4].ItemStyle.Width = 0;
                ClientGrid.Columns[4].Visible = false;
            }
        }
    }
}