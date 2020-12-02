using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Collections;
using Newtonsoft.Json;
using System.Web;
using System.Runtime.InteropServices;
using System.Web.Http.Cors;
using System.IO;
using AES_Cryptography;
using Classes;
using DAL;
using System.Configuration;

namespace WebAPI.Controllers
{
    //[EnableCors(origins: "localhost:4200", headers: "*", methods: "*")]
    [RoutePrefix("api/usermgt")]
    public class UserMgtController : ApiController
    {
        String DBPath = ConfigurationManager.AppSettings["DBPath"].ToString();
        String DBPwd = ConfigurationManager.AppSettings["DBPwd"].ToString();

        Generic objGeneric = new Generic();
        ResponseView objResponse = new ResponseView();

        [HttpPost]
        [Route("auth")]
        public HttpResponseMessage Auth(csUserMgt obj)
        {
            ResponseObject res = new ResponseObject();
            try
            {

                string ErrString = string.Empty;
                DataSet ds = new DataSet();

               AES objpwd = new AES();

                DAL_UserMgt obj_UserMgt = new DAL_UserMgt();
                DataTable dtUserDetails = new DataTable();
                int cid = obj.int_SiteID;
                string username = obj.str_UserName;
                string ADDomain = string.Empty;
                bool ADLogin = false;

                int errno = 0;
                string errstring = string.Empty;

                string pwd = objpwd.AES_Encrypt(obj.str_Password);
                //var result = obj_UserMgt.GetUserDetails(new Tuple<int, string, string, string, bool>(cid, username, pwd, ADDomain, ADLogin));
                obj_UserMgt.GetUserDetails(DBPath, DBPwd, cid, username, pwd, ADDomain, ADLogin, ref errno, ref errstring, ref dtUserDetails);
                //dtUserDetails = result.dsUserdetails.Tables[0];
                if (dtUserDetails != null && dtUserDetails.Rows.Count > 0)
                {
                    if (pwd == dtUserDetails.Rows[0]["Password"].ToString())
                    {
                        Hashtable ht = new Hashtable();
                        ht.Add("cid", cid);
                        ht.Add("userid", dtUserDetails.Rows[0]["UserID"].ToString());
                        ht.Add("username", username);
                        ht.Add("ledgerid", dtUserDetails.Rows[0]["LedgerID"]);
                        ht.Add("password", dtUserDetails.Rows[0]["Password"].ToString());
                        ht.Add("groupid", dtUserDetails.Rows[0]["GroupID"].ToString());
                        ht.Add("groupname", dtUserDetails.Rows[0]["GroupName"].ToString());
                        string encrypttoken = JsonConvert.SerializeObject(ht);
                        encrypttoken = objpwd.AES_Encrypt(encrypttoken);

                        DataTable dtConfigParam = new DataTable();
                        dtConfigParam = obj_UserMgt.GetConfigParam(DBPath, DBPwd, cid);
                        int salesmanid = obj_UserMgt.GetSalesmanIDByLedgerID(DBPath, DBPwd, cid, Convert.ToInt32(dtUserDetails.Rows[0]["ledgerid"]));
                        res.respdata = new User() { userid = Convert.ToInt32(dtUserDetails.Rows[0]["UserID"]), username = username, ledgerid = Convert.ToInt32(dtUserDetails.Rows[0]["ledgerid"]), groupid = Convert.ToInt16(dtUserDetails.Rows[0]["GroupID"]), token = encrypttoken, configparam = dtConfigParam, salesmanid = salesmanid };
                    }
                    else
                    {
                        res.errno = 1;
                        res.errdesc = "Wrong password";
                    }
                   
                }
                else
                {
                    res.errno = 1;
                    res.errdesc = "Login failed";
                }
                //}
                return Request.CreateResponse(HttpStatusCode.OK, res);
            }
            catch (Exception e)
            {
                //throw e;
                res.errno = 1;
                res.errdesc = e.Message;
                return Request.CreateResponse(HttpStatusCode.ExpectationFailed, res);
            }

        }

        [HttpGet]
        [Route("buildsidemenu")]
        public HttpResponseMessage GetBuildSideMenu(int cid, int uniqid, int groupid, string flag)
        {

            DataTable dtGroupMgtSub = new DataTable();
            DAL_UserMgt objGrpMgt = new DAL_UserMgt();

            //var result = objGrpMgt.GetGrpMgt(new Tuple<int, int, string>(cid, groupid, "getGroupMgtSub"));
            //dtGroupMgtSub = result.dtGroupmgt;


            //var result1 = objGrpMgt.GetMenuGrouping(new Tuple<int, int, string>(cid, uniqid, flag));
            //DataTable dt_Form = new DataTable();
            //dt_Form = result1.dtGroupmgt;

            dtGroupMgtSub = getGroupMgtSub();

            DataTable dt_Form = new DataTable();
            dt_Form = getForm();

            dt_Form.Rows.RemoveAt(0);
            List<menu> objmenulist = new List<menu>();
            Dictionary<string, menu> dic_MenuObject = new Dictionary<string, menu>();
            menu ParentMenuObject;
            foreach (DataRow drow in dt_Form.Rows)
            {
                menu objmenu = new menu();
                objmenu.name = drow["FormName"].ToString();
                objmenu.menuid = drow["MenuID"].ToString();
                objmenu.parentid = drow["Parent"].ToString();
                objmenu.icon = drow["WebIcon"].ToString();
                objmenu.visible = false;

                if (drow["parameters"].ToString() != "" && drow["parameters"].ToString() != "[]")
                {
                    DataTable dt = (DataTable)JsonConvert.DeserializeObject(drow["parameters"].ToString(), (typeof(DataTable)));
                    string[] menuparam = dt.Rows[0][0].ToString().Split('=');

                    objmenu.parameters = menuparam[1];
                }
                else
                {
                    objmenu.parameters = "";
                }

                dic_MenuObject.Add(drow["MenuID"].ToString(), objmenu);

                if (drow["Parent"].ToString() != "1")
                {

                    if (dic_MenuObject.ContainsKey(drow["Parent"].ToString()))
                    {
                        dic_MenuObject.TryGetValue(drow["Parent"].ToString(), out ParentMenuObject);
                        try
                        {
                            if (drow["TYPE"].ToString() == "FORM")
                            {
                                if (dtGroupMgtSub.Rows.Count > 0)
                                {
                                    DataRow[] resultMenu = dtGroupMgtSub.Select("MenuID ='" + drow["MenuID"].ToString() + "'");
                                    if (resultMenu.Count() > 0)
                                    {
                                        objmenu.visible = true;
                                        ParentMenuObject.children.Add(objmenu);
                                        enableParentNode(objmenu, ref dic_MenuObject);
                                    }

                                }

                            }
                            else
                            {
                                objmenu.visible = false;
                                ParentMenuObject.children.Add(objmenu);
                            }
                        }
                        catch (Exception e)
                        {
                            string msg = e.Message.ToString();
                        }
                    }

                }
                else
                {
                    objmenulist.Add(objmenu);
                }
            }


            return Request.CreateResponse(HttpStatusCode.OK, objmenulist);
        }

        // infuture we need to remove below two hardcoded datatable, instead read it from DB
        private DataTable getForm()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("FormName");
            dt.Columns.Add("MenuID");
            dt.Columns.Add("Parent");
            dt.Columns.Add("WebIcon");
            dt.Columns.Add("parameters");
            dt.Columns.Add("TYPE");

            DataRow drow;
            
            drow= dt.NewRow();
            drow["FormName"] = "MenuGroup";
            drow["MenuID"] = "1";
            drow["Parent"] = "0";
            drow["WebIcon"] = "av_timer";
            drow["parameters"] = "";
            drow["TYPE"] = "GROUP";
            dt.Rows.Add(drow);

            drow = dt.NewRow();
            drow["FormName"] = "File";
            drow["MenuID"] = "30";
            drow["Parent"] = "1";
            drow["WebIcon"] = "av_timer";
            drow["parameters"] = "";
            drow["TYPE"] = "GROUP";
            dt.Rows.Add(drow);

            drow = dt.NewRow();
            drow["FormName"] = "New Order";
            drow["MenuID"] = "ERP_155";
            drow["Parent"] = "30";
            drow["WebIcon"] = "av_timer";
            drow["parameters"] = "";
            drow["TYPE"] = "FORM";
            dt.Rows.Add(drow);

            return dt;
        }

        private DataTable getGroupMgtSub()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("MenuID");
            dt.Columns.Add("Options");

            DataRow drow;

            drow = dt.NewRow();
            drow["MenuID"] = "ERP_155";
            drow["Options"] = "Add";
            dt.Rows.Add(drow);

            drow = dt.NewRow();
            drow["MenuID"] = "ERP_155";
            drow["Options"] = "Edit";
            dt.Rows.Add(drow);

            drow = dt.NewRow();
            drow["MenuID"] = "ERP_155";
            drow["Options"] = "Delete";
            dt.Rows.Add(drow);

            drow = dt.NewRow();
            drow["MenuID"] = "ERP_155";
            drow["Options"] = "View";
            dt.Rows.Add(drow);

            return dt;
        }

        private void enableParentNode(menu objmenu, ref Dictionary<string, menu> objdict)
        {
            menu ParentMenuObject;
            objdict.TryGetValue(objmenu.parentid, out ParentMenuObject);
            if (ParentMenuObject != null)
            {
                ParentMenuObject.visible = true;
                enableParentNode(ParentMenuObject, ref objdict);
            }
        }

        //Get All Groups/Languages/Other company access details by passing the company id    

        //[HttpGet]
        //[Route("loaddetails")]
        //public HttpResponseMessage LoadDetails(int cid, string tablename, string type, string condition, string menuid, [Optional] bool withinactive)
        //{

        //    type = string.IsNullOrEmpty(type) ? "" : "";
        //    ResponseObject res = new ResponseObject();
        //    try
        //    {
        //        DataSet ds;
        //        DAL_UserMgt obj = new DAL_UserMgt();
        //        ds = obj.getDataSourceForUserMgt(cid);
        //        if (ds.Tables.Count > 0)
        //        {
        //            ds.Tables[0].TableName = "groups";
        //            ds.Tables[1].TableName = "languages";
        //            ds.Tables[2].TableName = "companies";

        //            DAL_General objGen = new DAL_General();
        //            DataTable dt_employees = new DataTable();
        //            dt_employees = objGen.LoadMCCBWithLedger(new Tuple<int, string>(cid, tablename), type, condition, menuid, withinactive);
        //            dt_employees.TableName = "employees";
        //            ds.Tables.Add(dt_employees);

        //            res.respdata = ds;

        //        }
        //        return Request.CreateResponse(HttpStatusCode.OK, res);
        //    }
        //    catch (Exception e)
        //    {
        //        res.errno = 1;
        //        res.errdesc = e.Message;
        //        return Request.CreateResponse(HttpStatusCode.ExpectationFailed, res);
        //    }
        //}

    }
}
