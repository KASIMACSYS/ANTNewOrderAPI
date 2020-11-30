'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_DO
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef IsInvCS As Boolean, ByRef Obj As csDO, ByRef ErrNo As Integer,
                             ByRef ErrStr As String)

        ErrNo = 0
        ErrStr = ""

        If Obj.objDOMain.str_MenuID = "ERP_157" Then
            Try
                BaseConn.Open(_StrDBPath, _StrDBPwd)
                BaseConn.cmd = New SqlClient.SqlCommand("[GetDODetails]", BaseConn.cnn)
                BaseConn.cmd.CommandType = CommandType.StoredProcedure
                BaseConn.cmd.Parameters.AddWithValue("@DoNo", Obj.objDOMain.str_DoNo)
                BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objDOMain.int_BusinessPeriodID)
                BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
                BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objDOMain.str_Flag)
                BaseConn.cmd.Parameters.AddWithValue("@RevisionHistoryNo", Obj.objDOMain.int_RevisionHistoryNo)
                BaseConn.cmd.Parameters.Add("@IsInvCS", SqlDbType.Bit).Direction = ParameterDirection.Output
                BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
                Dim ds As New DataSet
                BaseConn.da.Fill(ds)

                IsInvCS = BaseConn.cmd.Parameters("@IsInvCS").Value.ToString
                Obj.objDOMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()

                If Obj.objDOMain.str_Flag = "INVOICE" Then
                    Obj.objDOMain.int_LedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
                    Obj.objDOMain.str_PayTerm = ds.Tables(0).Rows(0)("PaymentTerm").ToString()
                Else
                    Obj.objDOMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                    Obj.objDOMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                End If


                Obj.objDOMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objDOMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()



                Obj.objDOMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()

                Obj.objDOMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objDOMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objDOMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

                Obj.objDOMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objDOMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objDOMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objDOMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objDOMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objDOMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                Obj.objDOMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
                Obj.objDOMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                'Obj.objDOMain.str_ContactPerson = ds.Tables(0).Rows(0)("ContactPerson").ToString()

                If Obj.objDOMain.str_Flag.ToUpper = "DO" Then
                    Obj.objDOMain.str_QtnNo = ds.Tables(0).Rows(0)("QtnNo").ToString()
                    Obj.objDOMain.str_SalOrd = ds.Tables(0).Rows(0)("SalOrd").ToString()
                    Obj.objDOMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                    Obj.objDOMain.dtp_DODate1 = ds.Tables(0).Rows(0)("DODate1").ToString()
                    Obj.objDOMain.dtp_DoDate2 = ds.Tables(0).Rows(0)("DODate2").ToString()
                    'Obj.objDOMain.dbl_VouNetCostPrice = ds.Tables(0).Rows(0)("VouNetCostPrice").ToString()
                    'Obj.objDOMain.dbl_VouNetProfit = ds.Tables(0).Rows(0)("VouNetProfit").ToString()
                    Obj.objDOMain.str_MerchantRef = ds.Tables(0).Rows(0)("MerchantRef").ToString()
                    Obj.objDOMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel")
                    Obj.objDOMain.str_SIS = ds.Tables(0).Rows(0)("SISNo").ToString()
                    Obj.objDOMain.dbl_SISAmt = ds.Tables(0).Rows(0)("SISAmount").ToString()

                    Obj.objDOMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
                    Obj.objDOMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
                    Obj.objDOMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
                    Obj.objDOMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString() 'TODO
                    Obj.objDOMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
                    Obj.objDOMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                    Obj.objDOMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()
                    Obj.objDOMain.str_Consignee = ds.Tables(0).Rows(0)("Consignee").ToString()

                    Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                    Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                    Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                    Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                    Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                    Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                    Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()

                ElseIf Obj.objDOMain.str_Flag.ToUpper = "SALESORDER" Then
                    Obj.objDOMain.str_QtnNo = "N/A"
                    Obj.objDOMain.str_MerchantRef = ds.Tables(0).Rows(0)("MerchantRef").ToString()
                    Obj.objDOMain.str_SalOrd = ds.Tables(0).Rows(0)("SalOrd").ToString()
                    Obj.objDOMain.dtp_DODate1 = Date.Now
                    Obj.objDOMain.dtp_DoDate2 = Date.Now
                    Obj.objDOMain.int_StatusCancel = 0
                    Obj.objDOMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString
                    Obj.objDOMain.str_Consignee = ds.Tables(0).Rows(0)("Consignee").ToString()
                ElseIf Obj.objDOMain.str_Flag.ToUpper = "QUOTATION" Then
                    Obj.objDOMain.str_QtnNo = ds.Tables(0).Rows(0)("QtnNo").ToString()
                    Obj.objDOMain.str_SalOrd = "N/A"
                    Obj.objDOMain.dtp_DODate1 = Date.Now
                    Obj.objDOMain.dtp_DoDate2 = Date.Now
                    Obj.objDOMain.int_StatusCancel = 0
                    Obj.objDOMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString
                ElseIf Obj.objDOMain.str_Flag.ToUpper = "INVOICE" Then
                    Obj.objDOMain.str_SIS = ds.Tables(0).Rows(0)("SISNO").ToString()
                    Obj.objDOMain.str_SalOrd = "N/A"
                    Obj.objDOMain.str_QtnNo = "N/A"
                    Obj.objDOMain.dtp_DODate1 = Date.Now
                    Obj.objDOMain.dtp_DoDate2 = ds.Tables(0).Rows(0)("DueDate").ToString()
                    Obj.objDOMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()
                    Obj.objDOMain.int_StatusCancel = 0
                    Obj.objDOMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString
                End If
                Obj.objDOMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objDOMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objDOMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objDOMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objDOMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objDOMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objDOMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objDOMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objDOMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

                Obj.objDOMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode").ToString()
                Obj.objDOMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                Obj.objDOMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage")
                Obj.objDOMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")

                Obj.objDOMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()
                Obj.objDOMain.str_ApproverComment = ds.Tables(0).Rows(0)("ApproverComment").ToString()
                Obj.objDOMain.int_LanguageCode = ds.Tables(0).Rows(0)("LanguageCode").ToString()

                If Obj.objDOMain.str_Flag.ToUpper <> "QUOTATION" Then

                    Obj.objDOMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
                    Obj.objDOMain.str_ContactPerson = ds.Tables(0).Rows(0)("ContactPerson").ToString()
                Else
                    Obj.objDOMain.str_DeliveryAddress = ""
                    Obj.objDOMain.str_ContactPerson = ""
                    Obj.objDOMain.str_MerchantRef = ""
                End If
                Obj.objDOSub.dt_DOSub = ds.Tables(1)

                If ds.Tables.Count >= 3 Then
                    If ds.Tables(2).Rows.Count > 0 Then
                        Obj.objProject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                        Obj.objProject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                        Obj.objProject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                    Else
                        Obj.objProject.str_ProjectID = ""
                        Obj.objProject.str_ProjectLocation = ""
                        Obj.objProject.str_WorkOrderNo = ""
                    End If
                End If
                Obj.objDOMain.dt_TaxItemDetails = ds.Tables(3)
                If ds.Tables.Count >= 4 Then
                    Obj.DTBatch = ds.Tables(4)
                End If

                If ds.Tables(5).Rows.Count > 0 Then
                    Obj.objDOMain.str_RTF_Description = ds.Tables(5).Rows(0)("Description").ToString()
                Else
                    Obj.objDOMain.str_RTF_Description = ""
                End If

                Obj.DTItemExtraDetails = ds.Tables(6)

            Catch ex As Exception
                ErrNo = 1
                ErrStr = ex.Message
            Finally
                BaseConn.Close()
            End Try


        Else
            Try
                BaseConn.Open(_StrDBPath, _StrDBPwd)
                BaseConn.cmd = New SqlClient.SqlCommand("[GetDOForPackaging]", BaseConn.cnn)
                BaseConn.cmd.CommandType = CommandType.StoredProcedure
                BaseConn.cmd.Parameters.AddWithValue("@DoNo", Obj.objDOMain.str_DoNo)
                BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objDOMain.int_BusinessPeriodID)
                BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
                BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
                Dim ds As New DataSet
                BaseConn.da.Fill(ds)

                Obj.objDOMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
                Obj.objDOMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objDOMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objDOMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objDOMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objDOMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()

                Obj.objDOMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objDOMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objDOMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

                Obj.objDOMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objDOMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objDOMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objDOMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objDOMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                Obj.objDOMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
                Obj.objDOMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()

                Obj.objDOMain.str_QtnNo = ds.Tables(0).Rows(0)("QtnNo").ToString()
                Obj.objDOMain.str_SalOrd = ds.Tables(0).Rows(0)("SalOrd").ToString()
                Obj.objDOMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objDOMain.dtp_DODate1 = ds.Tables(0).Rows(0)("DODate1").ToString()
                Obj.objDOMain.dtp_DoDate2 = ds.Tables(0).Rows(0)("DODate2").ToString()
                Obj.objDOMain.str_MerchantRef = ds.Tables(0).Rows(0)("MerchantRef").ToString()
                Obj.objDOMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel")
                Obj.objDOMain.str_SIS = ds.Tables(0).Rows(0)("SISNo").ToString()
                Obj.objDOMain.dbl_SISAmt = ds.Tables(0).Rows(0)("SISAmount").ToString()

                'Obj.objDOMain.dbl_TotalTax = ds.Tables(0).Rows(0)("TotalTax").ToString()
                Obj.objDOMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
                Obj.objDOMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString() 'TODO

                Obj.objDOMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()


                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()


                Obj.objDOMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objDOMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objDOMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objDOMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objDOMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objDOMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objDOMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objDOMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                Obj.objDOMain.str_Packaging = ds.Tables(0).Rows(0)("pkg_MainRemarks").ToString()

                Obj.objDOMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
                Obj.objDOMain.str_ContactPerson = ds.Tables(0).Rows(0)("ContactPerson").ToString()

                Obj.objDOSub.dt_DOSub = ds.Tables(1)

                If ds.Tables.Count >= 3 Then
                    If ds.Tables(2).Rows.Count > 0 Then
                        Obj.objProject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                        Obj.objProject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                        Obj.objProject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                    Else
                        Obj.objProject.str_ProjectID = ""
                        Obj.objProject.str_ProjectLocation = ""
                        Obj.objProject.str_WorkOrderNo = ""
                    End If
                End If

            Catch ex As Exception
                ErrNo = 1
                ErrStr = ex.Message
            Finally
                BaseConn.Close()
            End Try
        End If
    End Sub


    Public Function Update_DO(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef VouNo As String, ByRef intRevNo As Integer, ByVal obj As csDO, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0

        If obj.objDOMain.str_MenuID = "ERP_157" Then
            Try
                BaseConn.Open(_StrDBPath, _StrDBPwd)
                BaseConn.cmd = New SqlClient.SqlCommand("DOUpdate", BaseConn.cnn)
                BaseConn.cmd.CommandType = CommandType.StoredProcedure
                BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
                BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objDOMain.str_FormPrefix)
                BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objDOMain.str_Flag)
                BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objDOMain.str_MenuID)
                BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objDOMain.int_BusinessPeriodID)
                BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objDOMain.int_RevNo)
                BaseConn.cmd.Parameters.AddWithValue("@DoNo", obj.objDOMain.str_DoNo)
                BaseConn.cmd.Parameters.AddWithValue("@SalOrd", obj.objDOMain.str_SalOrd)
                BaseConn.cmd.Parameters.AddWithValue("@QtnNo", obj.objDOMain.str_QtnNo)
                BaseConn.cmd.Parameters.AddWithValue("@DODate1", obj.objDOMain.dtp_DODate1)
                BaseConn.cmd.Parameters.AddWithValue("@DODate2", obj.objDOMain.dtp_DoDate2)
                BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objDOMain.int_LedgerID)
                BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objDOMain.str_Alias)
                BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objDOMain.int_Aging)
                BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.objDOMain.str_PayTerm)
                BaseConn.cmd.Parameters.AddWithValue("@MerchantRef", obj.objDOMain.str_MerchantRef)
                BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objDOMain.str_Comment)
                BaseConn.cmd.Parameters.AddWithValue("@SISNo", obj.objDOMain.str_SIS)
                BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.objDOMain.str_SalesManID)
                BaseConn.cmd.Parameters.AddWithValue("@SalesManName", obj.objDOMain.str_SalesManName)
                BaseConn.cmd.Parameters.AddWithValue("@DeliveryAddress", obj.objDOMain.str_DeliveryAddress)
                BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objDOMain.str_TCCurrency)
                BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objDOMain.dbl_ExchangeRate)

                BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objDOMain.dbl_TCAmount)
                BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objDOMain.dbl_TCDisAmount)
                BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objDOMain.dbl_TCDiscountAmount)
                BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objDOMain.dbl_TCAdjAmount)
                BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objDOMain.dbl_TCNetAmount)
                BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.objDOMain.dbl_TCMiscPercentage)
                BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.objDOMain.dbl_TCMiscAmount)

                BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objDOMain.dbl_LCNetAmount)
                BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.objDOMain.str_MiscText)
                BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objDOMain.str_DiscText)
                BaseConn.cmd.Parameters.AddWithValue("@SISAmount", obj.objDOMain.dbl_SISAmt)
                'BaseConn.cmd.Parameters.AddWithValue("@TotalTax", obj.objDOMain.dbl_TotalTax)

                BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.objDOMain.str_WHID)
                BaseConn.cmd.Parameters.AddWithValue("@Consignee", obj.objDOMain.str_Consignee)
                BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objDOMain.str_Desc1)
                BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objDOMain.str_Desc2)
                BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objDOMain.str_Desc3)
                BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objDOMain.str_Desc4)
                BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objDOMain.str_Desc5)
                BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objDOMain.str_Desc6)
                BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objDOMain.str_Desc7)
                BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objDOMain.str_Desc8)

                BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objDOMain.str_InvoiceTaxCode)

                BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objDOMain.str_ItemTaxCode)
                BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objDOMain.dbl_TCItemTaxAmount)
                BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objDOMain.dbl_TCInvoiceTaxAmount)
                BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objDOMain.str_InvoiceTaxXML)
                BaseConn.cmd.Parameters.AddWithValue("@ItemDiscPercentage", obj.objDOMain.dbl_ItemDiscPercentage)

                BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
                BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
                BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
                BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
                BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
                BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
                BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)
                BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", obj.ApprovedLevel)
                BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", obj.ApprovedComment)
                BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", obj.ApprovedHigherLevel)

                BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objProject.str_ProjectID)
                BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objProject.str_WorkOrderNo)
                BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objProject.str_ProjectLocation)

                BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objDOMain.int_StatusCancel)
                BaseConn.cmd.Parameters.AddWithValue("@ContactPerson", obj.objDOMain.str_ContactPerson)

                BaseConn.cmd.Parameters.AddWithValue("@UserComment", obj.objDOMain.str_UserComment)
                BaseConn.cmd.Parameters.AddWithValue("@ApproverComment", obj.objDOMain.str_ApproverComment)
                BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.objDOMain.int_LanguageCode)

                BaseConn.cmd.Parameters.AddWithValue("@RTF_Description", obj.objDOMain.str_RTF_Description)

                BaseConn.cmd.Parameters.AddWithValue("@DOItemDetailsDT", obj.objDOSub.dt_DOSub)
                BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)
                BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objDOMain.dt_TaxItemDetails)
                BaseConn.cmd.Parameters.AddWithValue("@VoucherItemExtraDetailsDT", obj.DTItemExtraDetails)

                BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
                BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
                BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
                BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
                BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
                BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

                BaseConn.cmd.CommandTimeout = 500
                BaseConn.cmd.ExecuteNonQuery()
                VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
                intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
                ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
                _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
                _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
                _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
            Catch ex As Exception
                _ErrString = ex.Message
                ObjDalGeneral = New DAL_General(obj.int_CID)
                ObjDalGeneral.Elog_Insert(obj.int_CID, _StrDBPath, _StrDBPwd, obj.objDOMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "DO", Err.Number, "Error in " & obj.objDOMain.str_Flag & " : " & obj.objDOMain.str_DoNo & "", ex.Message, 5, 3, 1, ErrNo)
                ErrNo = 1
            Finally
                BaseConn.Close()
            End Try

        Else
            Try
                BaseConn.Open(_StrDBPath, _StrDBPwd)
                BaseConn.cmd = New SqlClient.SqlCommand("DOPackagingUpdate", BaseConn.cnn)
                BaseConn.cmd.CommandType = CommandType.StoredProcedure
                BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
                BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objDOMain.int_BusinessPeriodID)
                BaseConn.cmd.Parameters.AddWithValue("@DoNo", obj.objDOMain.str_DoNo)
                BaseConn.cmd.Parameters.AddWithValue("@PkgComment", obj.objDOMain.str_Packaging)
                BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)

                BaseConn.cmd.Parameters.AddWithValue("@DOPackagingDT", obj.objDOSub.dt_DOSub)
                BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
                BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
                BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
                BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
                BaseConn.cmd.CommandTimeout = 500
                BaseConn.cmd.ExecuteNonQuery()
                VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
                intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value
                ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
                _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            Catch ex As Exception
                _ErrString = ex.Message
                ObjDalGeneral = New DAL_General(obj.int_CID)
                ObjDalGeneral.Elog_Insert(obj.int_CID, _StrDBPath, _StrDBPwd, obj.objDOMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "DO PACKAGE", Err.Number, "Error in " & obj.objDOMain.str_Flag & " : " & obj.objDOMain.str_DoNo & "", ex.Message, 5, 3, 1, ErrNo)
                ErrNo = 1
            Finally
                BaseConn.Close()
            End Try
        End If


        Update_DO = _ErrString
    End Function

    Public Function GetMissMatchedItems(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal GivenItems As DataTable, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMismatchedItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ReceivedItemDT", GivenItems)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
End Class
