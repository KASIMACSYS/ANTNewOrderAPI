'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_MerchantMaster
    Private ObjDalGeneral As DAL_General
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()

    Public Sub Get_Structure(ByRef Obj As csMerchantMaster, ByVal objLedger As csLedgerMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMerchantDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Obj.ObjMerMain.str_MerchantID)
            ''BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjMerMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Count", 0)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjMerMain.str_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
            Obj.ObjMerMain.str_MerchantID = ds.Tables(0).Rows(0)("MerchantID").ToString()
            Obj.ObjMerMain.str_MerchantName = ds.Tables(0).Rows(0)("MerchantName").ToString()
            Obj.ObjMerMain.str_Type = ds.Tables(0).Rows(0)("Type").ToString()
            Obj.ObjMerMain.str_Alias1 = ds.Tables(0).Rows(0)("Alias1").ToString()
            Obj.ObjMerMain.str_Alias2 = ds.Tables(0).Rows(0)("Alias2").ToString()

            Obj.ObjMerMain.bool_CusActiveStatus = ds.Tables(0).Rows(0)("CusActiveStatus").ToString()
            Obj.ObjMerMain.int_CusCreditDays = ds.Tables(0).Rows(0)("CusCreditDays").ToString()
            Obj.ObjMerMain.dbl_CusCreditLimitAmount = ds.Tables(0).Rows(0)("CusCreditLimitAmount").ToString()
            Obj.ObjMerMain.int_CusCreditLimitCondition = ds.Tables(0).Rows(0)("CusCreditLimitCondition").ToString()
            Obj.ObjMerMain.int_CusAgingLimitCondition = ds.Tables(0).Rows(0)("CusCreditDaysCondition").ToString()
            Obj.ObjMerMain.bool_CusCreditDaysRemindExpiry = ds.Tables(0).Rows(0)("CusCreditDaysRemindExpiry").ToString()
            Obj.ObjMerMain.str_CusPayTerm = ds.Tables(0).Rows(0)("CusPayTerm").ToString()
            Obj.ObjMerMain.bool_CusCreditDaysPDC = ds.Tables(0).Rows(0)("CusCreditDaysPDC").ToString()
            Obj.ObjMerMain.bool_CusCreditLimitPDC = ds.Tables(0).Rows(0)("CusCreditLimitPDC").ToString()

            Obj.ObjMerMain.int_PurchaseAccLedgerID = ds.Tables(0).Rows(0)("PurchaseAccLedgerID").ToString()
            Obj.ObjMerMain.int_SalesAccLedgerID = ds.Tables(0).Rows(0)("SalesAccLedgerID").ToString()
            Obj.ObjMerMain.int_CashPurchaseAccLedgerID = ds.Tables(0).Rows(0)("CashPurchaseAccLedgerID").ToString()
            Obj.ObjMerMain.int_CashSalesAccLedgerID = ds.Tables(0).Rows(0)("CashSalesAccLedgerID").ToString()
            Obj.ObjMerMain.int_PurchaseRTNAccLedgerID = ds.Tables(0).Rows(0)("PurchaseRTNAccLedgerID").ToString()
            Obj.ObjMerMain.int_SalesRTNAccLedgerID = ds.Tables(0).Rows(0)("SalesRTNAccLedgerID").ToString()

            Obj.ObjMerMain.bool_VenActiveStatus = ds.Tables(0).Rows(0)("VenActiveStatus").ToString()
            Obj.ObjMerMain.int_VenCreditDays = ds.Tables(0).Rows(0)("VenCreditDays").ToString()
            Obj.ObjMerMain.dbl_VenCreditLimitAmount = ds.Tables(0).Rows(0)("VenCreditLimitAmount").ToString()
            Obj.ObjMerMain.int_VenCreditLimitCondition = ds.Tables(0).Rows(0)("VenCreditLimitCondition").ToString()
            Obj.ObjMerMain.int_VenAgingLimitCondition = ds.Tables(0).Rows(0)("VenCreditDaysCondition").ToString()
            Obj.ObjMerMain.bool_VenCreditDaysRemindExpiry = ds.Tables(0).Rows(0)("VenCreditDaysRemindExpiry").ToString()
            Obj.ObjMerMain.str_VenPayTerm = ds.Tables(0).Rows(0)("VenPayTerm").ToString()
            Obj.ObjMerMain.bool_VenCreditDaysPDC = ds.Tables(0).Rows(0)("VenCreditDaysPDC").ToString()

            Obj.ObjMerMain.bool_ReverseCharge = ds.Tables(0).Rows(0)("ReverseCharge").ToString()

            Obj.ObjMerMain.int_SellingPercentage = ds.Tables(0).Rows(0)("SellingPercentage").ToString()
            Obj.ObjMerMain.bool_IsSellingPercentage = ds.Tables(0).Rows(0)("IsSellingPercentage").ToString()
            Obj.ObjMerMain.str_Contact = ds.Tables(0).Rows(0)("Contact").ToString()
            Obj.ObjMerMain.str_Address = ds.Tables(0).Rows(0)("Address1").ToString()
            Obj.ObjMerMain.str_DelivAddress = ds.Tables(0).Rows(0)("Address2").ToString()
            Obj.ObjMerMain.dtp_MerchantSince = ds.Tables(0).Rows(0)("MerchantSince").ToString()
            Obj.ObjMerMain.str_PoBox = ds.Tables(0).Rows(0)("PoBox").ToString()
            Obj.ObjMerMain.str_Tel = ds.Tables(0).Rows(0)("Tel").ToString()
            Obj.ObjMerMain.str_Mobile = ds.Tables(0).Rows(0)("Mobile").ToString()
            Obj.ObjMerMain.str_Fax = ds.Tables(0).Rows(0)("Fax").ToString()
            Obj.ObjMerMain.str_Email = ds.Tables(0).Rows(0)("Email").ToString()
            Obj.ObjMerMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjMerMain.str_PopUpComment = ds.Tables(0).Rows(0)("PopUpComment").ToString()
            Obj.ObjMerMain.str_ChequePrintingName = ds.Tables(0).Rows(0)("ChequePrintingName").ToString()
            Obj.ObjMerMain.bool_IntraCompanyFlag = ds.Tables(0).Rows(0)("IntraCompanyFlag").ToString()
            Obj.ObjMerMain.str_DefaultItemValue = ds.Tables(0).Rows(0)("DefaultItemValue").ToString()
            Obj.ObjMerMain.str_City = ds.Tables(0).Rows(0)("City").ToString
            Obj.ObjMerMain.str_Filter1 = ds.Tables(0).Rows(0)("Filter1").ToString
            Obj.ObjMerMain.str_Filter2 = ds.Tables(0).Rows(0)("Filter2").ToString
            Obj.ObjMerMain.str_Filter3 = ds.Tables(0).Rows(0)("Filter3").ToString
            Obj.ObjMerMain.str_Filter4 = ds.Tables(0).Rows(0)("Filter4").ToString
            Obj.ObjMerMain.str_TimeZone = ds.Tables(0).Rows(0)("CusTimeZone").ToString
            Obj.ObjMerMain.str_SupportMailID = ds.Tables(0).Rows(0)("SupportMailID").ToString
            Obj.ObjMerMain.str_ItemDiscType = ds.Tables(0).Rows(0)("ItemDiscType").ToString
            Obj.ObjMerMain.str_Trn = ds.Tables(0).Rows(0)("Trn").ToString
            Obj.ObjMerMain.str_Country = ds.Tables(0).Rows(0)("Country").ToString
            Obj.ObjMerMain.str_Region = ds.Tables(0).Rows(0)("Region").ToString
            Obj.ObjMerMain.str_Consignee = ds.Tables(0).Rows(0)("Consignee").ToString

            Obj.ObjMerMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString
            Obj.ObjMerMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString
            Obj.ObjMerMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString
            Obj.ObjMerMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString
            Obj.ObjMerMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString
            Obj.ObjMerMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString
            Obj.ObjMerMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString
            Obj.ObjMerMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString
            Obj.ObjMerMain.bool_SendSMS = ds.Tables(0).Rows(0)("SendSMS").ToString
            Obj.ObjMerMain.bool_SendEmail = ds.Tables(0).Rows(0)("SendEmail").ToString

            If ds.Tables(0).Rows(0)("DefaultSalesMan").ToString() <> "" Then
                Obj.ObjMerMain.int_SalesMan = ds.Tables(0).Rows(0)("DefaultSalesMan").ToString()
            Else
                Obj.ObjMerMain.int_SalesMan = 0
            End If

            Obj.ObjMerSub.STSiteID = ds.Tables(0).Rows(0)("SiteID").ToString()
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_MerMaster(ByVal obj As csMerchantMaster, ByVal objLedger As csLedgerMaster, ByRef str_MerchantID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("MerchantMasterUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantID", obj.ObjMerMain.str_MerchantID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantName", obj.ObjMerMain.str_MerchantName)
            BaseConn.cmd.Parameters.AddWithValue("@Contact", obj.ObjMerMain.str_Contact)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantType", obj.ObjMerMain.str_Type)
            BaseConn.cmd.Parameters.AddWithValue("@Alias1", obj.ObjMerMain.str_Alias1)
            BaseConn.cmd.Parameters.AddWithValue("@Alias2", obj.ObjMerMain.str_Alias2)

            BaseConn.cmd.Parameters.AddWithValue("@CusActiveStatus", obj.ObjMerMain.bool_CusActiveStatus)
            BaseConn.cmd.Parameters.AddWithValue("@CusPayTerm", obj.ObjMerMain.str_CusPayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@CusCreditDays", obj.ObjMerMain.int_CusCreditDays)
            BaseConn.cmd.Parameters.AddWithValue("@CusCreditLimitAmount", obj.ObjMerMain.dbl_CusCreditLimitAmount)
            BaseConn.cmd.Parameters.AddWithValue("@CusCreditLimitCondition", obj.ObjMerMain.int_CusCreditLimitCondition)
            BaseConn.cmd.Parameters.AddWithValue("@CusCreditDaysCondition", obj.ObjMerMain.int_CusAgingLimitCondition)
            BaseConn.cmd.Parameters.AddWithValue("@CusCreditDaysRemindExpiry", obj.ObjMerMain.bool_CusCreditDaysRemindExpiry)
            BaseConn.cmd.Parameters.AddWithValue("@CusCreditDaysPDC", obj.ObjMerMain.bool_CusCreditDaysPDC)
            BaseConn.cmd.Parameters.AddWithValue("@CusCreditLimitPDC", obj.ObjMerMain.bool_CusCreditLimitPDC)

            BaseConn.cmd.Parameters.AddWithValue("@VenActiveStatus", obj.ObjMerMain.bool_VenActiveStatus)
            BaseConn.cmd.Parameters.AddWithValue("@VenPayTerm", obj.ObjMerMain.str_VenPayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@VenCreditDays", obj.ObjMerMain.int_VenCreditDays)
            BaseConn.cmd.Parameters.AddWithValue("@VenCreditLimitAmount", obj.ObjMerMain.dbl_VenCreditLimitAmount)
            BaseConn.cmd.Parameters.AddWithValue("@VenCreditLimitCondition", obj.ObjMerMain.int_VenCreditLimitCondition)
            BaseConn.cmd.Parameters.AddWithValue("@VenCreditDaysCondition", obj.ObjMerMain.int_VenAgingLimitCondition)
            BaseConn.cmd.Parameters.AddWithValue("@VenCreditDaysRemindExpiry", obj.ObjMerMain.bool_VenCreditDaysRemindExpiry)
            BaseConn.cmd.Parameters.AddWithValue("@VenCreditDaysPDC", obj.ObjMerMain.bool_VenCreditDaysPDC)
            BaseConn.cmd.Parameters.AddWithValue("@ReverseCharge", obj.ObjMerMain.bool_ReverseCharge)

            BaseConn.cmd.Parameters.AddWithValue("@PurchaseAccLedgerID", obj.ObjMerMain.int_PurchaseAccLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesAccLedgerID", obj.ObjMerMain.int_SalesAccLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@CashPurchaseAccLedgerID", obj.ObjMerMain.int_CashPurchaseAccLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@CashSalesAccLedgerID", obj.ObjMerMain.int_CashSalesAccLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@PurchaseRTNAccLedgerID", obj.ObjMerMain.int_PurchaseRTNAccLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesRTNAccLedgerID", obj.ObjMerMain.int_SalesRTNAccLedgerID)

            BaseConn.cmd.Parameters.AddWithValue("@SellingPercentage", obj.ObjMerMain.int_SellingPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@IsSellingPercentage", obj.ObjMerMain.bool_IsSellingPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@Address1", obj.ObjMerMain.str_Address)
            BaseConn.cmd.Parameters.AddWithValue("@Address2", obj.ObjMerMain.str_DelivAddress)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantSince", obj.ObjMerMain.dtp_MerchantSince)
            BaseConn.cmd.Parameters.AddWithValue("@PoBox", obj.ObjMerMain.str_PoBox)
            BaseConn.cmd.Parameters.AddWithValue("@Tel", obj.ObjMerMain.str_Tel)
            BaseConn.cmd.Parameters.AddWithValue("@Mobile", obj.ObjMerMain.str_Mobile)
            BaseConn.cmd.Parameters.AddWithValue("@Fax", obj.ObjMerMain.str_Fax)
            BaseConn.cmd.Parameters.AddWithValue("@Email", obj.ObjMerMain.str_Email)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjMerMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@PopUpComment", obj.ObjMerMain.str_PopUpComment)
            BaseConn.cmd.Parameters.AddWithValue("@ChequePrintingName", obj.ObjMerMain.str_ChequePrintingName)
            BaseConn.cmd.Parameters.AddWithValue("@IntraCompanyFlag", obj.ObjMerMain.bool_IntraCompanyFlag)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjMerMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjMerMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjMerMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjMerMain.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@DefaultItemValue", obj.ObjMerMain.str_DefaultItemValue)
            BaseConn.cmd.Parameters.AddWithValue("@STSiteID", obj.ObjMerSub.STSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DefaultSalesMan", obj.ObjMerMain.int_SalesMan)

            '----------Ledger---------------
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", objLedger.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Description", objLedger.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerType", objLedger.str_LedgerType)
            BaseConn.cmd.Parameters.AddWithValue("@ParentAccount", objLedger.str_ParentAccount)
            BaseConn.cmd.Parameters.AddWithValue("@Class", objLedger.str_Class)
            BaseConn.cmd.Parameters.AddWithValue("@StartRange", objLedger.str_StartRange)
            BaseConn.cmd.Parameters.AddWithValue("@EndRange", objLedger.str_EndRange)
            BaseConn.cmd.Parameters.AddWithValue("@AccountCode1", objLedger.str_AccountNo1)
            BaseConn.cmd.Parameters.AddWithValue("@AccountCode2", objLedger.str_AccountNo2)
            BaseConn.cmd.Parameters.AddWithValue("@Status", objLedger.bool_InActive)
            BaseConn.cmd.Parameters.AddWithValue("@ReadOnly", objLedger.bool_Readonly)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerComment", objLedger.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@Category", objLedger.str_Category)

            BaseConn.cmd.Parameters.AddWithValue("@Amount", objLedger.dbl_Amount)
            BaseConn.cmd.Parameters.AddWithValue("@Advance", objLedger.dbl_Advance)


            BaseConn.cmd.Parameters.AddWithValue("@Date", obj.ObjMerMain.dtp_date)
            BaseConn.cmd.Parameters.AddWithValue("@City", obj.ObjMerMain.str_City)
            BaseConn.cmd.Parameters.AddWithValue("@Filter1", obj.ObjMerMain.str_Filter1)
            BaseConn.cmd.Parameters.AddWithValue("@Filter2", obj.ObjMerMain.str_Filter2)
            BaseConn.cmd.Parameters.AddWithValue("@Filter3", obj.ObjMerMain.str_Filter3)
            BaseConn.cmd.Parameters.AddWithValue("@Filter4", obj.ObjMerMain.str_Filter4)
            BaseConn.cmd.Parameters.AddWithValue("@TimeZone", obj.ObjMerMain.str_TimeZone)
            BaseConn.cmd.Parameters.AddWithValue("@SupportMailID", obj.ObjMerMain.str_SupportMailID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDiscType", obj.ObjMerMain.str_ItemDiscType)
            BaseConn.cmd.Parameters.AddWithValue("@Trn", obj.ObjMerMain.str_Trn)
            BaseConn.cmd.Parameters.AddWithValue("@Country", obj.ObjMerMain.str_Country)
            BaseConn.cmd.Parameters.AddWithValue("@Region", obj.ObjMerMain.str_Region)
            BaseConn.cmd.Parameters.AddWithValue("@Consignee", obj.ObjMerMain.str_Consignee)

            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.ObjMerMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.ObjMerMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.ObjMerMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.ObjMerMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.ObjMerMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.ObjMerMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.ObjMerMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.ObjMerMain.str_Desc8)
            BaseConn.cmd.Parameters.AddWithValue("@SendSMS", obj.ObjMerMain.bool_SendSMS)
            BaseConn.cmd.Parameters.AddWithValue("@SendEmail", obj.ObjMerMain.bool_SendEmail)


            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjMerMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@MerchIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            str_MerchantID = BaseConn.cmd.Parameters("@MerchIDOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.ObjMerMain.str_CreatedBy, obj.ObjMerMain.dtp_CreatedDate, "", "MerchantMaster", Err.Number, "Error in " & obj.ObjMerMain.str_Flag & " : " & obj.ObjMerMain.str_MerchantName & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_MerMaster = _ErrString
    End Function
    Public Sub GetMerchantDetails(ByVal _StrSiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _LedgerID As String, ByRef _dtMerchantDetails As DataTable)
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMerchantDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", "MERCHANTDETAILS")
            BaseConn.cmd.Parameters.AddWithValue("@Count", 0)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        _dtMerchantDetails = dt
    End Sub


    Public Sub Get_LedgerDetails(ByRef objLedger As csLedgerMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetLedgerMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", objLedger.int_LedgerID)
            ' BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", objLedger.str_SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            objLedger.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
            objLedger.str_Description = ds.Tables(0).Rows(0)("Description").ToString()
            objLedger.str_LedgerType = ds.Tables(0).Rows(0)("LedgerType").ToString()
            objLedger.str_ParentAccount = ds.Tables(0).Rows(0)("ParentAccount").ToString()
            objLedger.str_Class = ds.Tables(0).Rows(0)("Class").ToString()
            objLedger.str_StartRange = ds.Tables(0).Rows(0)("StartRange").ToString()
            objLedger.str_EndRange = ds.Tables(0).Rows(0)("EndRange").ToString()
            objLedger.str_AccountNo1 = ds.Tables(0).Rows(0)("AccountCode1").ToString()
            objLedger.str_AccountNo2 = ds.Tables(0).Rows(0)("AccountCode2").ToString()
            objLedger.bool_InActive = ds.Tables(0).Rows(0)("Status").ToString()
            objLedger.bool_Readonly = ds.Tables(0).Rows(0)("ReadOnly").ToString()
            objLedger.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            objLedger.str_Category = ds.Tables(0).Rows(0)("Category").ToString()

            'objLedger.dbl_Amount = ds.Tables(0).Rows(0)("Amount").ToString()
            'objLedger.dbl_Advance = ds.Tables(0).Rows(0)("Advance").ToString()
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Sub GetMerchantSellPercentage(ByVal _StrSiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _LedgerID As Integer, _
                                         ByRef isSellPercentage As Boolean, ByRef SellPercentage As Double, ByRef DefaultItemValue As String, _
                                         Optional ByRef DefaultItemDiscType As String = "")
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMerchantSellPercentage]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.Add("@IsSellPercentage", SqlDbType.Bit).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@SellPercentage", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@DefaultItemValue", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@DefaultItemDiscType", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            isSellPercentage = BaseConn.cmd.Parameters("@IsSellPercentage").Value
            SellPercentage = BaseConn.cmd.Parameters("@SellPercentage").Value
            DefaultItemValue = BaseConn.cmd.Parameters("@DefaultItemValue").Value.ToString
            DefaultItemDiscType = BaseConn.cmd.Parameters("@DefaultItemDiscType").Value.ToString
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub IsMerchantExists(ByVal _StrSiteID As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _MerchantID As String,
                                ByRef _LedgerID As Integer)

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[IsMerchantExists]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantID", _MerchantID)
            BaseConn.cmd.Parameters.Add("@LedgerID", SqlDbType.Int).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()
            _LedgerID = BaseConn.cmd.Parameters("@LedgerID").Value
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try

    End Sub




End Class
