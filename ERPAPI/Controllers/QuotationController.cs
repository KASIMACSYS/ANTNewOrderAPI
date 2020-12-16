using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Data;
using System.Web;
using System.IO;
using Classes;
using DAL;
using System.Configuration;
using WebAPI;

namespace ERPAPI.Controllers
{
    [RoutePrefix("api/quotation")]
    public class QuotationController : ApiController
    {
        String DBPath = ConfigurationManager.AppSettings["DBPath"].ToString();
        String DBPwd = ConfigurationManager.AppSettings["DBPwd"].ToString();
        DAL_Quotation obj;
        DAL_General objGen;
        
        [HttpGet]
        [Route("getcustomersalesman")]
        public HttpResponseMessage CustomerAndSalesman(int cid)
        {
            ResponseObject res = new ResponseObject();
            try
            {
                DataSet ds = new DataSet();
                DataTable dtCustomer = new DataTable();
                //DataTable dtSalesman = new DataTable();
                objGen = new DAL_General(cid.ToString());

                dtCustomer = objGen.GetCustomer(DBPath, DBPwd, cid);
                if (dtCustomer.Rows.Count>0)
                {
                    dtCustomer.TableName = "Customers";
                    ds.Tables.Add(dtCustomer);
                }

                //dtSalesman = objGen.GetSalesmanList(DBPath, DBPwd, cid);
                //if (dtSalesman.Rows.Count > 0)
                //{
                //    dtSalesman.TableName = "Salesman";
                //    ds.Tables.Add(dtSalesman);
                //}

                res.respdata = ds;
                return Request.CreateResponse(HttpStatusCode.OK, res);
            }
            catch (Exception e)
            {
                res.errno = 1;
                res.errdesc = e.Message;
                return Request.CreateResponse(HttpStatusCode.ExpectationFailed, res);
            }
        }

        [HttpGet]
        [Route("getsalesmanquotation")]
        public HttpResponseMessage GetSalesmanQuotation(int cid, int salesmanid, DateTime fromdate, DateTime todate, string qtnstatus)
        {
            ResponseObject res = new ResponseObject();
            try
            {
                int errno = 0;
                string errstring = string.Empty;
                DataSet ds = new DataSet();
                DataTable dtSalesmanQuotation = new DataTable();
                obj = new DAL_Quotation();

                dtSalesmanQuotation = obj.GetSalesmanQuotation(DBPath, DBPwd, cid, salesmanid, fromdate, todate, qtnstatus, ref errno, ref errstring);
                if (dtSalesmanQuotation.Rows.Count > 0)
                {
                    dtSalesmanQuotation.TableName = "SalesmanQuotation";
                }
                                
                res.respdata = dtSalesmanQuotation;
                return Request.CreateResponse(HttpStatusCode.OK, res);
            }
            catch (Exception e)
            {
                res.errno = 1;
                res.errdesc = e.Message;
                return Request.CreateResponse(HttpStatusCode.ExpectationFailed, res);
            }
        }


        [HttpPost]
        [Route("neworder")]
        public HttpResponseMessage NewOrder()
        {
            string imagename = null;
            string imagetype = null;
            var httprequest = HttpContext.Current.Request;

            //upload image
            var postedfile = httprequest.Files["Image"];
            //create custom filename
            imagename = new string(Path.GetFileNameWithoutExtension(postedfile.FileName).Take(10).ToArray()).Replace(" ", "-");
            imagename = imagename + DateTime.Now.ToString("yymmssfff") + Path.GetExtension(postedfile.FileName);

            //var filepath = HttpContext.Current.Server.MapPath("~/Image/" + imagename);
            //postedfile.SaveAs(filepath);

            string imageCaption = httprequest["ImageCaption"];
            string customerledger = httprequest["CustomerLedger"];
            string customername = httprequest["CustomerName"];
            string salesmanid = httprequest["SalesmanID"];
            string username = httprequest["UserName"];

            Stream fs = postedfile.InputStream;
            BinaryReader br = new BinaryReader(fs);
            byte[] bytes = br.ReadBytes((Int32)fs.Length);
            string result = createQuoation(Convert.ToInt16(customerledger), customername, Convert.ToInt32(salesmanid), username, bytes, imageCaption, imagetype);
            return Request.CreateResponse(HttpStatusCode.Created);
        }

        private string createQuoation(int customerledger, string customername, int salesmanid, string username, Byte[] image, string imagename, string imagetype)
        {
            string qtnNo = string.Empty;
            string outsms = string.Empty, outemail = string.Empty;
            string errstring = string.Empty;
            int revno = 0;
            int errno = 0;
            obj = new DAL_Quotation();
            csQuotation objcsqtn = CreateQtnObject(customerledger, customername, salesmanid, username);

            errstring = obj.Update_Quotation(DBPath, DBPwd, ref qtnNo, ref revno, objcsqtn, ref outsms, ref outemail, ref errno);
            //errstring = obj.Update_Quotation(ref qtnNo, ref revno, objcsqtn, ref outsms, ref outemail, ref errno);
            if (errstring == "" && qtnNo != "")
                UpdateImage(objcsqtn.str_CID, objcsqtn.objQuotationMain.str_FormPrefix + qtnNo, image, imagename, imagetype, username);

            return qtnNo;
        }

        private csQuotation CreateQtnObject(int customerledger, string customername, int salesmanid, string username)
        {
            Dictionary<string, string> objproj = new Dictionary<string, string>();
            
            csQuotation objqtn = new csQuotation(objproj);
            objqtn.str_CID = "101";
            objGen = new DAL_General(objqtn.str_CID);
            objqtn.objQuotationMain.str_QtnNo = "";
            objqtn.objQuotationMain.int_BusinessPeriodID = objGen.GetLatestBusinessPeriodID(DBPath, DBPwd,101);
            objqtn.objQuotationMain.str_Flag = "ADD";
            objqtn.objQuotationMain.int_RevNo = 0;
            objqtn.objQuotationMain.Str_QtnStatus = "Open";
            objqtn.objQuotationMain.str_MenuID = "ERP_155";
            objqtn.objQuotationMain.str_FormPrefix = "QTN/";
            objqtn.objQuotationMain.int_LedgerID = customerledger;
            objqtn.objQuotationMain.str_Alias = customername;
            objqtn.objQuotationMain.str_EstNo = "";
            objqtn.objQuotationMain.dtp_QtnDate = DateTime.Now;
            objqtn.objQuotationMain.int_Aging = 0;
            objqtn.objQuotationMain.str_PayTerm = "";
            objqtn.objQuotationMain.str_IndRef = "";
            objqtn.objQuotationMain.dbl_TCAmount = 0;
            objqtn.objQuotationMain.dbl_TCDisAmount = "0";
            objqtn.objQuotationMain.dbl_TCDiscountAmount = 0;
            objqtn.objQuotationMain.dbl_TCNetAmount = 0;
            objqtn.objQuotationMain.dbl_TCItemTaxAmount = 0;
            objqtn.objQuotationMain.dbl_TCInvoiceTaxAmount = 0;
            objqtn.objQuotationMain.dbl_TCMiscAmount = 0;
            objqtn.objQuotationMain.dbl_TCMiscPercentage = "0";
            objqtn.objQuotationMain.dbl_TCAdjAmount = 0;
            objqtn.objQuotationMain.str_Comment = "";
            objqtn.objQuotationMain.str_Contact = "";
            objqtn.objQuotationMain.int_StatusCancel = 0;
            objqtn.objQuotationMain.str_DeliverIn = "";
            objqtn.objQuotationMain.str_QtnValidity = "";

            objqtn.objQuotationMain.str_Desc1 = "";
            objqtn.objQuotationMain.str_Desc2 = "";
            objqtn.objQuotationMain.str_Desc3 = "";
            objqtn.objQuotationMain.str_Desc4 = "";
            objqtn.objQuotationMain.str_Desc5 = "";
            objqtn.objQuotationMain.str_Desc6 = "";
            objqtn.objQuotationMain.str_Desc7 = "";
            objqtn.objQuotationMain.str_Desc8 = "";

            objqtn.objQuotationMain.dbl_ItemDiscPercentage = 0;
            objqtn.objQuotationMain._XMLCustomData = "";

            objqtn.objQuotationMain.str_ExpiryDays = "";
            objqtn.DTItemExtraDetails = ItemExtraDT();
            objqtn.objQuotationMain.str_SalesManID = salesmanid.ToString();

            objqtn.objproject.str_ProjectID = "";
            objqtn.objproject.str_ProjectLocation = "";
            objqtn.objproject.str_WorkOrderNo = "";

            objqtn.str_CreatedBy = username;
            objqtn.dtp_CreatedDate = DateTime.Now;
            objqtn.str_LastUpdatedBy = "";
            objqtn.dtp_LastUpdatedDate = DateTime.Now;
            objqtn.str_ApprovedBy = "";
            objqtn.dtp_ApprovedDate = DateTime.Now;
            objqtn.bool_ApprovedStatus = 1;
            objqtn.ApprovedHigherLevel = true;
            objqtn.ApprovedComment = "";

            objqtn.objQuotationMain.str_UserComment = "";
            objqtn.objQuotationMain.str_ApproverComment = "";

            objqtn.objQuotationMain.dbl_LCNetAmount = 0;
            objqtn.objQuotationMain.str_TCCurrency = "AED";
            objqtn.objQuotationMain.dbl_ExchangeRate = 1;
            objqtn.objQuotationMain.str_MiscText = "";
            objqtn.objQuotationMain.str_DiscText = "";
            objqtn.objQuotationMain.int_LanguageCode = 0;


            objqtn.objQuotationSub.dt_Quotation = DBTemplate();

            objqtn.objQuotationMain.str_ItemTaxCode = "";
            objqtn.objQuotationMain.str_InvoiceTaxCode = "";
            objqtn.objQuotationMain.str_InvoiceTaxXML = ConvertDatatableToXML(SingleItemTaxDetails());

            objqtn.objQuotationMain.dt_TaxItemDetails = TaxItemDetails();
            objqtn.objQuotationMain.str_RTF_Description = "";

            return objqtn;
        }

        private DataTable ItemExtraDT()
        {
            DataTable DT_ItemExtraDetails = new DataTable();
            DT_ItemExtraDetails.Columns.Add("SortNo");
            DT_ItemExtraDetails.Columns.Add("SlNo");
            DT_ItemExtraDetails.Columns.Add("ItemCode");
            DT_ItemExtraDetails.Columns.Add("ImagePath");
            DT_ItemExtraDetails.Columns.Add("Description");
            return DT_ItemExtraDetails;
        }

        private DataTable TaxItemDetails()
        {
            DataTable dt_TaxItemDetails = new DataTable();
            dt_TaxItemDetails.Columns.Add("TaxCode");
            dt_TaxItemDetails.Columns.Add("TaxableAmount", System.Type.GetType("System.Double"));
            dt_TaxItemDetails.Columns.Add("TaxAmount", System.Type.GetType("System.Double"));
            dt_TaxItemDetails.Columns.Add("NonClaimableTax", System.Type.GetType("System.Double"));
            dt_TaxItemDetails.Columns.Add("ReverseChargeAmount", System.Type.GetType("System.Double"));
            dt_TaxItemDetails.Columns.Add("Type");
            return dt_TaxItemDetails;
        }

        private DataTable DBTemplate()
        {
            DataTable DT_QTNTemplate = new DataTable();
            DT_QTNTemplate.Columns.Add("SortNo", System.Type.GetType("System.Int32"));
            DT_QTNTemplate.Columns.Add("SlNo", System.Type.GetType("System.Int32"));
            DT_QTNTemplate.Columns.Add("BarCodeNo");
            DT_QTNTemplate.Columns.Add("Alias1");
            DT_QTNTemplate.Columns.Add("Alias2");

            DT_QTNTemplate.Columns.Add("ItemCode");
            DT_QTNTemplate.Columns.Add("Package", System.Type.GetType("System.Double"));
            DT_QTNTemplate.Columns.Add("Pieces", System.Type.GetType("System.Double"));
            DT_QTNTemplate.Columns.Add("Unit");
            DT_QTNTemplate.Columns.Add("PriceType");

            DT_QTNTemplate.Columns.Add("BaseUnit", System.Type.GetType("System.Double"));
            DT_QTNTemplate.Columns.Add("VouQty", System.Type.GetType("System.Double"));
            DT_QTNTemplate.Columns.Add("PrimaryQty", System.Type.GetType("System.Double"));
            DT_QTNTemplate.Columns.Add("Price", System.Type.GetType("System.Double"));
            DT_QTNTemplate.Columns.Add("BaseUnitPrice", System.Type.GetType("System.Double"));

            DT_QTNTemplate.Columns.Add("DiscType");
            DT_QTNTemplate.Columns.Add("DiscPercentage", System.Type.GetType("System.Decimal"));
            DT_QTNTemplate.Columns.Add("TCDiscountAmount", System.Type.GetType("System.Decimal"));
            DT_QTNTemplate.Columns.Add("Amount", System.Type.GetType("System.Decimal"));
            DT_QTNTemplate.Columns.Add("Tax");
            DT_QTNTemplate.Columns.Add("TaxPercentage", System.Type.GetType("System.Double"));
            DT_QTNTemplate.Columns.Add("TaxAmount", System.Type.GetType("System.Decimal"));
            DT_QTNTemplate.Columns.Add("NonClaimableTaxAmount", System.Type.GetType("System.Decimal"));
            DT_QTNTemplate.Columns.Add("NetAmount", System.Type.GetType("System.Decimal"));
            DT_QTNTemplate.Columns.Add("LCAmount", System.Type.GetType("System.Decimal"));
            DT_QTNTemplate.Columns.Add("LCCostPrice", System.Type.GetType("System.Double"));
            DT_QTNTemplate.Columns.Add("DeliveredTotQty", System.Type.GetType("System.Double"));
            DT_QTNTemplate.Columns.Add("PartNo");
            DT_QTNTemplate.Columns.Add("Comment");
            DT_QTNTemplate.Columns.Add("Desc1");
            DT_QTNTemplate.Columns.Add("Desc2");
            DT_QTNTemplate.Columns.Add("Desc3");
            DT_QTNTemplate.Columns.Add("Desc4");
            DT_QTNTemplate.Columns.Add("Desc5");
            DT_QTNTemplate.Columns.Add("Desc6");
            DT_QTNTemplate.Columns.Add("Desc7");
            DT_QTNTemplate.Columns.Add("Desc8");
            DT_QTNTemplate.Columns.Add("MinSellPrice", System.Type.GetType("System.Decimal"));
            DT_QTNTemplate.Columns.Add("ItemTaxDetails");

            DataRow drow = DT_QTNTemplate.NewRow();
            drow["SortNo"] = 1;
            drow["SlNo"] = 1;
            drow["BarCodeNo"] = "";
            drow["Alias1"] = "SERVICES";
            drow["Alias2"] = "SERVICES";

            drow["ItemCode"] = "service";
            drow["Package"] = 0;
            drow["Pieces"] = 0;
            drow["Unit"] = "Nos";
            drow["PriceType"] = "";

            drow["BaseUnit"] = 1;
            drow["VouQty"] = 1;
            drow["PrimaryQty"] = 1;
            drow["Price"] = 1;
            drow["BaseUnitPrice"] = 1;
            drow["DiscType"] = "";
            drow["DiscPercentage"] = 0;
            drow["TCDiscountAmount"] = 0;
            drow["Amount"] = 0;
            drow["Tax"] = "";
            drow["TaxPercentage"] = 0;
            drow["TaxAmount"] = 0;
            drow["NonClaimableTaxAmount"] = 0;
            drow["NetAmount"] = 0;
            drow["LCAmount"] = 0;
            drow["LCCostPrice"] = 0;
            drow["DeliveredTotQty"] = 0;
            drow["PartNo"] = "";
            drow["Comment"] = "";
            drow["Desc1"] = "";
            drow["Desc2"] = "";
            drow["Desc3"] = "";
            drow["Desc4"] = "";
            drow["Desc5"] = "";
            drow["Desc6"] = "";
            drow["Desc7"] = "";
            drow["Desc8"] = "";
            drow["MinSellPrice"] = 0;
            drow["ItemTaxDetails"] = "";
            DT_QTNTemplate.Rows.Add(drow);
            return DT_QTNTemplate;
        }


        private DataTable SingleItemTaxDetails()
        {
            DataTable dt_SingleItemTax = new DataTable();
            dt_SingleItemTax.Columns.Add("TaxCode");
            dt_SingleItemTax.Columns.Add("TaxableAmount", System.Type.GetType("System.Double"));
            dt_SingleItemTax.Columns.Add("PurchaseTaxPercentage", System.Type.GetType("System.Double"));
            dt_SingleItemTax.Columns.Add("SalesTaxPercentage", System.Type.GetType("System.Double"));
            dt_SingleItemTax.Columns.Add("TaxAmount", System.Type.GetType("System.Double"));
            dt_SingleItemTax.Columns.Add("NonClaimableTax", System.Type.GetType("System.Double"));
            dt_SingleItemTax.Columns.Add("ReverseChargeAmount", System.Type.GetType("System.Double"));
            return dt_SingleItemTax;
        }

        private string ConvertDatatableToXML(DataTable dt)
        {
            string xmlstr = string.Empty;

            if (dt.Rows.Count > 0)
            {
                MemoryStream str = new MemoryStream();
                dt.WriteXml(str, true);
                str.Seek(0, SeekOrigin.Begin);
                StreamReader sr = new StreamReader(str);
                xmlstr = sr.ReadToEnd();
            }

            return (xmlstr);
        }

        private void UpdateImage(string cid, string qtnno, Byte[] image, string imagename, string imagetype, string updatedby)
        {
            string errstring = string.Empty;
            int errno = 0;
            DateTime updateddate = DateTime.Now;
            string slno = "1";
            DAL_General obj = new DAL_General(cid);
            obj.FileUpload(cid, DBPath, DBPwd, 101, qtnno, "Quotation", "ERP_155", "", "", imagename, ".jpg", updatedby, image,ref slno, "ADD", ref errno, ref errstring);
            //errstring = obj.FileUpload(cid, qtnno, "Quotation", "ERP_155", "ORDER", "", imagename, imagetype, updatedby, image, "1", "ADD", ref errno, ref errstring);

        }

    }
}
