using Classes;
using DAL;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using WebAPI;

namespace ERPAPI.Controllers
{
    [RoutePrefix("api/salesorder")]
    public class SalesOrderController : ApiController
    {
        String DBPath = ConfigurationManager.AppSettings["DBPath"].ToString();
        String DBPwd = ConfigurationManager.AppSettings["DBPwd"].ToString();
        DAL_SalesOrder obj;
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
                objGen = new DAL_General(cid.ToString());

                dtCustomer = objGen.GetCustomer(DBPath, DBPwd, cid);
                if (dtCustomer.Rows.Count > 0)
                {
                    dtCustomer.TableName = "Customers";
                    ds.Tables.Add(dtCustomer);
                }
                               
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
               
        //[HttpGet]
        //[Route("getsalesmanquotation")]
        //public HttpResponseMessage GetSalesmanQuotation(int cid, int salesmanid, DateTime fromdate, DateTime todate, string qtnstatus)
        //{
        //    ResponseObject res = new ResponseObject();
        //    try
        //    {
        //        int errno = 0;
        //        string errstring = string.Empty;
        //        DataSet ds = new DataSet();
        //        DataTable dtSalesmanQuotation = new DataTable();
        //        obj = new DAL_Quotation();

        //        dtSalesmanQuotation = obj.GetSalesmanQuotation(DBPath, DBPwd, cid, salesmanid, fromdate, todate, qtnstatus, ref errno, ref errstring);
        //        if (dtSalesmanQuotation.Rows.Count > 0)
        //        {
        //            dtSalesmanQuotation.TableName = "SalesmanQuotation";
        //        }

        //        res.respdata = dtSalesmanQuotation;
        //        return Request.CreateResponse(HttpStatusCode.OK, res);
        //    }
        //    catch (Exception e)
        //    {
        //        res.errno = 1;
        //        res.errdesc = e.Message;
        //        return Request.CreateResponse(HttpStatusCode.ExpectationFailed, res);
        //    }
        //}

        [HttpGet]
        [Route("getsalesmanorder")]
        public HttpResponseMessage GetSalesmanOrder(int cid, int salesmanid, string qtnstatus)
        {
            ResponseObject res = new ResponseObject();
            try
            {
                int errno = 0;
                string errstring = string.Empty;
                DataSet ds = new DataSet();
                DataTable dtSalesmanQuotation = new DataTable();
                obj = new DAL_SalesOrder();

                dtSalesmanQuotation = obj.GetSalesmanOrder(DBPath, DBPwd, cid, salesmanid, qtnstatus, ref errno, ref errstring);
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

        [HttpGet]
        [Route("getorderdashboard")]
        public HttpResponseMessage GetQuotationDashboard(int cid, int salesmanid)
        {
            ResponseObject res = new ResponseObject();
            try
            {
                int errno = 0;
                string errstring = string.Empty;
                int[] arr_values = new int[4];
                DataSet ds = new DataSet();
                DataTable dtSalesmanQuotation = new DataTable();
                DataTable dtMonthQuotation = new DataTable();
                obj = new DAL_SalesOrder();

                ds = obj.MA_OrderDashboard(DBPath, DBPwd, cid, salesmanid, ref errno, ref errstring);
                dtSalesmanQuotation = ds.Tables[0];
                for (int i = 0; i < dtSalesmanQuotation.Rows.Count; i++)
                {
                    arr_values[i] = Convert.ToInt32(ds.Tables[0].Rows[i]["Cnt"]);
                }
                //if (dtSalesmanQuotation.Rows.Count > 0)
                //{
                //    dtSalesmanQuotation.TableName = "series";
                //}
                var pie = new { series = arr_values };

                //barchart data
                dtMonthQuotation = ds.Tables[1]; //GetQuotationDashboardDT();
                string[] arr_month = new string[5];
                List<int[]> monthQtn = new List<int[]>();
                int[] arr_status = new int[5];
                for (int i = 0; i < dtMonthQuotation.Rows.Count; i++)
                {
                    arr_month[i] = dtMonthQuotation.Rows[i]["Month"].ToString();
                    arr_status = new int[4];
                    arr_status[0] = Convert.ToInt32(dtMonthQuotation.Rows[i]["Open"]);
                    arr_status[1] = Convert.ToInt32(dtMonthQuotation.Rows[i]["Close"]);
                    arr_status[2] = Convert.ToInt32(dtMonthQuotation.Rows[i]["Partial"]);
                    arr_status[3] = Convert.ToInt32(dtMonthQuotation.Rows[i]["ManualClose"]);
                    monthQtn.Add(arr_status);
                }

                var barLabels = new { labels = arr_month };
                var barSeries = new { series = monthQtn };

                Dictionary<string, object> dic_QtnObject = new Dictionary<string, object>();
                dic_QtnObject.Add("Count", dtSalesmanQuotation);
                dic_QtnObject.Add("ChartData", pie);
                dic_QtnObject.Add("BarLabels", barLabels);
                dic_QtnObject.Add("BarSeries", barSeries);
                res.respdata = dic_QtnObject;
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
        [Route("getitempricelist")]
        public HttpResponseMessage GetItemPriceList(int cid, string itemcode, int salesmanid, string barcode)
        {
            ResponseObject res = new ResponseObject();
            try
            {
                //int errno = 0;
                //string errstring = string.Empty;
                DataSet ds = new DataSet();
                objGen = new DAL_General(cid.ToString());

                ds = objGen.MA_ItemPriceList(DBPath, DBPwd, cid, itemcode, salesmanid, barcode);
                ds.Tables[0].TableName = "ItemMaster";
                ds.Tables[1].TableName = "ItemPriceList";
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
        [Route("getitemlist")]
        public HttpResponseMessage GetItemList(int cid)
        {
            ResponseObject res = new ResponseObject();
            try
            {
                //int errno = 0;
                //string errstring = string.Empty;
                DataTable dt = new DataTable();
                objGen = new DAL_General(cid.ToString());

                dt = objGen.GetItemList(DBPath, DBPwd, cid);
                //dt.TableName = "ItemList";
                res.respdata = dt;
                return Request.CreateResponse(HttpStatusCode.OK, res);
            }
            catch (Exception e)
            {
                res.errno = 1;
                res.errdesc = e.Message;
                return Request.CreateResponse(HttpStatusCode.ExpectationFailed, res);
            }
        }

        private DataTable GetQuotationDashboardDT()
        {
            DataTable GetQuotationDashboardDT = new DataTable();
            GetQuotationDashboardDT.Columns.Add("Month");
            GetQuotationDashboardDT.Columns.Add("Open");
            GetQuotationDashboardDT.Columns.Add("Close");
            GetQuotationDashboardDT.Columns.Add("Partial");

            DataRow drow;
            drow = GetQuotationDashboardDT.NewRow();
            drow["Month"] = "Dec";
            drow["Open"] = 2;
            drow["Close"] = 1;
            drow["Partial"] = 1;
            GetQuotationDashboardDT.Rows.Add(drow);

            drow = GetQuotationDashboardDT.NewRow();
            drow["Month"] = "Nov";
            drow["Open"] = 2;
            drow["Close"] = 1;
            drow["Partial"] = 0;
            GetQuotationDashboardDT.Rows.Add(drow);

            drow = GetQuotationDashboardDT.NewRow();
            drow["Month"] = "Oct";
            drow["Open"] = 3;
            drow["Close"] = 0;
            drow["Partial"] = 1;
            GetQuotationDashboardDT.Rows.Add(drow);

            return GetQuotationDashboardDT;
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

            string imageCaption = httprequest["ImageCaption"];
            string customerledger = httprequest["CustomerLedger"];
            string customername = httprequest["CustomerName"];
            string salesmanid = httprequest["SalesmanID"];
            string username = httprequest["UserName"];

            Stream fs = postedfile.InputStream;
            BinaryReader br = new BinaryReader(fs);
            byte[] bytes = br.ReadBytes((Int32)fs.Length);
            string result = createSalesOrder(Convert.ToInt16(customerledger), customername, Convert.ToInt32(salesmanid), username, bytes, imageCaption, imagetype);
            return Request.CreateResponse(HttpStatusCode.Created);
        }

        private string createSalesOrder(int customerledger, string customername, int salesmanid, string username, Byte[] image, string imagename, string imagetype)
        {
            string soNo = string.Empty;
            string outsms = string.Empty, outemail = string.Empty;
            string errstring = string.Empty;
            int revno = 0;
            int errno = 0;
            obj = new DAL_SalesOrder();
            csSalesOrder objcsSO = CreateSOObject(customerledger, customername, salesmanid, username);

            errstring = obj.Update_SalesOrder(DBPath, DBPwd, ref soNo, ref revno, objcsSO, ref outsms, ref outemail, ref errno);
            //errstring = obj.Update_Quotation(ref qtnNo, ref revno, objcsqtn, ref outsms, ref outemail, ref errno);
            if (errstring == "" && soNo != "")
                UpdateImage(objcsSO.int_CID.ToString(), objcsSO.objSalesOrderMain.str_FormPrefix + soNo, image, imagename, imagetype, username);

            return soNo;
        }

        private csSalesOrder CreateSOObject(int customerledger, string customername, int salesmanid, string username)
        {
            Dictionary<string, string> objproj = new Dictionary<string, string>();

            csSalesOrder objSO = new csSalesOrder(objproj);
            objSO.int_CID = 101;
            objGen = new DAL_General(objSO.int_CID.ToString());
            objSO.objSalesOrderMain.str_SalOrd = "";
            objSO.objSalesOrderMain.int_BusinessPeriodID = objGen.GetLatestBusinessPeriodID(DBPath, DBPwd, 101);
            objSO.objSalesOrderMain.str_Flag = "ADD";
            objSO.objSalesOrderMain.str_FormPrefix = "PER/";
            objSO.objSalesOrderMain.str_MenuID = "ERP_156";
            objSO.objSalesOrderMain.int_RevNo = 0;
            objSO.objSalesOrderMain.dtp_SODate = DateTime.Now;
            objSO.objSalesOrderMain.str_QtnNum = "";
            objSO.objSalesOrderMain.int_LedgerID = customerledger;
            objSO.objSalesOrderMain.str_Alias = customername;
            objSO.objSalesOrderMain.int_Aging = 0;
            objSO.objSalesOrderMain.str_PayTerm = "";
            objSO.objSalesOrderMain.str_Indref = "";
            objSO.objSalesOrderMain.str_Comment = "";
            objSO.objSalesOrderMain.str_Contact = "";
            objSO.objSalesOrderMain.str_SOStatus = "Open";
            objSO.objSalesOrderMain.str_MerchantRef = "";
            objSO.objSalesOrderMain.str_SalesManID = salesmanid.ToString();
            objSO.objSalesOrderMain.str_TCCurrency = "AED";
            objSO.objSalesOrderMain.dbl_ExchangeRate = 1;
            objSO.objSalesOrderMain.int_StatusCancel = 2;
            objSO.objSalesOrderMain.str_DeliveryAddress = "";
            objSO.objSalesOrderMain.str_ContactPerson = "";

            objSO.objSalesOrderMain.str_Desc1 = "";
            objSO.objSalesOrderMain.str_Desc2 = "";
            objSO.objSalesOrderMain.str_Desc3 = "";
            objSO.objSalesOrderMain.str_Desc4 = "";
            objSO.objSalesOrderMain.str_Desc5 = "";
            objSO.objSalesOrderMain.str_Desc6 = "";
            objSO.objSalesOrderMain.str_Desc7 = "";
            objSO.objSalesOrderMain.str_Desc8 = "";

            objSO.objSalesOrderMain.dbl_TCAmount = 0;
            objSO.objSalesOrderMain.dbl_TCDisAmount = "0";
            objSO.objSalesOrderMain.dbl_TCDiscountAmount = 0;
            objSO.objSalesOrderMain.dbl_TCAdjAmount = 0;
            objSO.objSalesOrderMain.dbl_TCNetAmount = 0;
            objSO.objSalesOrderMain.dbl_TCMiscPercentage = "0";
            objSO.objSalesOrderMain.dbl_TCMiscAmount = 0;
            objSO.objSalesOrderMain.dbl_LCNetAmount = 0;

            objSO.objSalesOrderMain.str_ExpiryDays = "";
            objSO.objSalesOrderMain.str_MiscText = "Misc";
            objSO.objSalesOrderMain.str_DiscText = "Discount";
            objSO.objSalesOrderMain.str_UserComment = "";
            objSO.objSalesOrderMain.str_ApproverComment = "";
            objSO.objSalesOrderMain.str_ItemTaxCode = "TAX";
            objSO.objSalesOrderMain.str_InvoiceTaxCode = "";
            objSO.objSalesOrderMain.str_InvoiceTaxXML = ConvertDatatableToXML(SingleItemTaxDetails());
            objSO.objSalesOrderMain.dbl_TCItemTaxAmount = 0;
            objSO.objSalesOrderMain.dbl_TCInvoiceTaxAmount = 0;
            
            objSO.objSalesOrderMain.dbl_ItemDiscPercentage = 0;
            objSO.objSalesOrderMain.str_WHID = "";
            objSO.objSalesOrderMain.str_Consignee = "";
            objSO.objSalesOrderMain.str_SalesType = "";
            objSO.objSalesOrderMain.str_DeliveryCountry = "";
            objSO.objSalesOrderMain.int_LanguageCode = 0;
            objSO.objSalesOrderMain.str_RTF_Description = "";

            
            objSO.objproject.str_ProjectID = "";
            objSO.objproject.str_ProjectLocation = "";
            objSO.objproject.str_WorkOrderNo = "";

            objSO.str_CreatedBy = username;
            objSO.dtp_CreatedDate = DateTime.Now;
            objSO.str_LastUpdatedBy = "";
            objSO.dtp_LastUpdatedDate = DateTime.Now;
            objSO.str_ApprovedBy = "";
            objSO.dtp_ApprovedDate = DateTime.Now;
            objSO.bool_ApprovedStatus = 1;
            objSO.ApprovedHigherLevel = true;
            objSO.ApprovedComment = "";

            objSO.DTItemExtraDetails = ItemExtraDT();
            objSO.objSalesorderSub.dt_SalesOrderItemDetails = DBTemplate();
            objSO.DTBatch = BatchDTTemplate();
            objSO.objSalesOrderMain.dt_TaxItemDetails = TaxItemDetails();
            return objSO;
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

        private DataTable DBTemplate()
        {
            DataTable DT_SOTemplate = new DataTable();
            DT_SOTemplate.Columns.Add("SortNo", System.Type.GetType("System.Int32"));
            DT_SOTemplate.Columns.Add("SlNo", System.Type.GetType("System.Int32"));
            DT_SOTemplate.Columns.Add("BarCodeNo");
            DT_SOTemplate.Columns.Add("Alias1");
            DT_SOTemplate.Columns.Add("Alias2");

            DT_SOTemplate.Columns.Add("ItemCode");
            DT_SOTemplate.Columns.Add("Package", System.Type.GetType("System.Double"));
            DT_SOTemplate.Columns.Add("Pieces", System.Type.GetType("System.Double"));
            DT_SOTemplate.Columns.Add("Unit");
            DT_SOTemplate.Columns.Add("PriceType");

            DT_SOTemplate.Columns.Add("BaseUnit", System.Type.GetType("System.Double"));
            DT_SOTemplate.Columns.Add("VouQty", System.Type.GetType("System.Double"));
            DT_SOTemplate.Columns.Add("PrimaryQty", System.Type.GetType("System.Double"));
            DT_SOTemplate.Columns.Add("Price", System.Type.GetType("System.Double"));
            DT_SOTemplate.Columns.Add("BaseUnitPrice", System.Type.GetType("System.Double"));

            DT_SOTemplate.Columns.Add("DiscType");
            DT_SOTemplate.Columns.Add("DiscPercentage", System.Type.GetType("System.Decimal"));
            DT_SOTemplate.Columns.Add("TCDiscountAmount", System.Type.GetType("System.Decimal"));
            DT_SOTemplate.Columns.Add("Amount", System.Type.GetType("System.Decimal"));
            DT_SOTemplate.Columns.Add("Tax");
            DT_SOTemplate.Columns.Add("TaxPercentage", System.Type.GetType("System.Double"));
            DT_SOTemplate.Columns.Add("TaxAmount", System.Type.GetType("System.Decimal"));
            DT_SOTemplate.Columns.Add("NonClaimableTaxAmount", System.Type.GetType("System.Decimal"));
            DT_SOTemplate.Columns.Add("NetAmount", System.Type.GetType("System.Decimal"));
            DT_SOTemplate.Columns.Add("LCAmount", System.Type.GetType("System.Decimal"));
            DT_SOTemplate.Columns.Add("LCCostPrice", System.Type.GetType("System.Double"));
            DT_SOTemplate.Columns.Add("DeliveredTotQty", System.Type.GetType("System.Double"));
            DT_SOTemplate.Columns.Add("PartNo");
            DT_SOTemplate.Columns.Add("Comment");
            DT_SOTemplate.Columns.Add("OrgSlno", System.Type.GetType("System.Int32"));
            DT_SOTemplate.Columns.Add("Desc1");
            DT_SOTemplate.Columns.Add("Desc2");
            DT_SOTemplate.Columns.Add("Desc3");
            DT_SOTemplate.Columns.Add("Desc4");
            DT_SOTemplate.Columns.Add("Desc5");
            DT_SOTemplate.Columns.Add("Desc6");
            DT_SOTemplate.Columns.Add("Desc7");
            DT_SOTemplate.Columns.Add("Desc8");
            DT_SOTemplate.Columns.Add("MinSellPrice", System.Type.GetType("System.Decimal"));
            DT_SOTemplate.Columns.Add("ItemTaxDetails");

            DataRow drow = DT_SOTemplate.NewRow();
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
            drow["OrgSlno"] = 1;
            drow["MinSellPrice"] = 0;
            drow["ItemTaxDetails"] = "";
            DT_SOTemplate.Rows.Add(drow);
            return DT_SOTemplate;
        }

        private DataTable BatchDTTemplate()
        {
            DataTable DTBatch = new DataTable();
         
            DTBatch.Columns.Add("OrgSlno", System.Type.GetType("System.Int32"));
            DTBatch.Columns.Add("ItemCode");
            DTBatch.Columns.Add("BatchID");
            DTBatch.Columns.Add("BinID");
            DTBatch.Columns.Add("FromBinID");
            DTBatch.Columns.Add("Serial");
            DTBatch.Columns.Add("Qty", System.Type.GetType("System.Double"));

            return DTBatch;
        }


        private void UpdateImage(string cid, string qtnno, Byte[] image, string imagename, string imagetype, string updatedby)
        {
            string errstring = string.Empty;
            int errno = 0;
            DateTime updateddate = DateTime.Now;
            string slno = "1";
            DAL_General obj = new DAL_General(cid);
            obj.FileUpload(cid, DBPath, DBPwd, 101, qtnno, "SalesOrder", "ERP_156", "", "", imagename, ".jpg", updatedby, image, ref slno, "ADD", ref errno, ref errstring);
            //errstring = obj.FileUpload(cid, qtnno, "Quotation", "ERP_155", "ORDER", "", imagename, imagetype, updatedby, image, "1", "ADD", ref errno, ref errstring);

        }

    }
}
