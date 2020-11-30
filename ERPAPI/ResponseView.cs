using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WebAPI
{
    public class ResponseView
    {

        public string DataSetToString(Int32 errNo, string errMsg, DataSet _ds)
        {

            ERPMAResponse objERPMAResponse = new ERPMAResponse
            {
                ErrorNo = errNo,
                ErrorMsg = errMsg,
                // Response = JsonConvert.SerializeObject(_ds.Tables[0])
                Response = _ds
            };

            string result = string.Empty;
            //result = JsonConvert.SerializeObject(_ds, Formatting.Indented);            
            result = JsonConvert.SerializeObject(objERPMAResponse, Formatting.Indented);
            return result;
        }

        public class ERPMAResponse
        {
            public Int32 ErrorNo { get; set; }
            public string ErrorMsg { get; set; }
            public DataSet Response { get; set; }
            // public IList<string> Response { get; set; }
        }
    }
}