using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using Newtonsoft.Json;
using AES_Cryptography;
using System.Collections;
using Classes;
using DAL;

namespace WebAPI
{
    public class Generic
    {
        public string DataSetToString(DataSet _ds)
        {
            string result = string.Empty;
            result = JsonConvert.SerializeObject(_ds, Formatting.Indented);
            return result;
        }

        public DataTable StringToDataTable(string json)
        {
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(json, (typeof(DataTable)));
            return dt;
        }

        public bool ValidateToken(string DBPath, string DBPwd, string encryptedtoken, ref string message)
        {
            bool Validate = true;
            try
            {
                AES objpwd = new AES();
                //DAL_LoginForm obj_DALLoginForm = new DAL_LoginForm();
                //DAL_UserMgt obj_UserMgt = new DAL_UserMgt();
                DAL_UserMgt obj_UserMgt = new DAL_UserMgt();

                Hashtable ht = new Hashtable();
                encryptedtoken = objpwd.AES_Decrypt(encryptedtoken);
                ht = (Hashtable)JsonConvert.DeserializeObject((encryptedtoken), (typeof(Hashtable)));
                int cid = Convert.ToInt16(ht["cid"]);
                string username = ht["username"].ToString();
                string password = ht["password"].ToString();
                int errno = 0;
                string errstring = string.Empty;

                string ADDomain = string.Empty;
                bool ADLogin = false;
                DataTable dtUserDetails = new DataTable();
                //obj_UserMgt.GetUserDetails(ref DBPath, ref DBPwd, ref cid, ref username, ref password, ref ADDomain, ref ADLogin, ref dtUserDetails, ref _ErrNo, ref ErrString);
                //var result = obj_UserMgt.GetUserDetails(new Tuple<int, string, string, string, bool>(cid, username, password, ADDomain, ADLogin));
                obj_UserMgt.GetUserDetails(DBPath, DBPwd, cid, username, password, ADDomain, ADLogin, ref errno, ref errstring, ref dtUserDetails);
                if (dtUserDetails.Rows.Count == 0)
                {
                    Validate = false;
                    message = "Invalid Token";
                }
            }
            catch
            {
                Validate = false;
                message = "Invalid Token";
            }

            return Validate;
        }
    }
}
