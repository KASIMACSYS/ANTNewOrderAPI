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
using System.Collections.Generic;
using System.Collections;

namespace WebAPI
{


    public class ResponseObject
    {
        public int errno = 0;
        public string errdesc = string.Empty;
        public Hashtable errorlist;
        public object respdata;
    }

    public class menu
    {
        public string name { get; set; }
        public string type { get; set; }
        public string icon { get; set; }
        public string menuid { get; set; }
        public string parentid { get; set; }
        public string parameters { get; set; }
        public bool visible { get; set; }
        public List<menu> children = new List<menu>();
    }

    public class User
    {
        public int userid { get; set; }
        public string username { get; set; }
        public int ledgerid { get; set; }
        public int groupid { get; set; }
        public string token { get; set; }
        public DataSet configparam { get; set; }
        public int salesmanid { get; set; }
    }
    
}