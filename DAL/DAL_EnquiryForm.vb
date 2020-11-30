'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_EnquiryForm

    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csEnquiryForm, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEnquiryDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@EnquiryNo", Obj.objEnquiryMain.str_EnquiryNo)
            BaseConn.cmd.Parameters.AddWithValue("@IndentNo", Obj.objEnquiryMain.str_IndentNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objEnquiryMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objEnquiryMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.objEnquiryMain.str_Flag = "ENQUIRY" Then
                Obj.objEnquiryMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
                Obj.objEnquiryMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objEnquiryMain.dtp_EnquiryDate1 = ds.Tables(0).Rows(0)("EnqDate1").ToString()
                Obj.objEnquiryMain.dtp_EnquiryDate2 = ds.Tables(0).Rows(0)("EnqDate2").ToString()
                Obj.objEnquiryMain.str_IndentNo = ds.Tables(0).Rows(0)("IndentNo").ToString()
                Obj.objEnquiryMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objEnquiryMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objEnquiryMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objEnquiryMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objEnquiryMain.str_IndRef = ds.Tables(0).Rows(0)("IndRef").ToString()
                Obj.objEnquiryMain.str_DelivAddress = ds.Tables(0).Rows(0)("DelivAddress").ToString()
                Obj.objEnquiryMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objEnquiryMain.str_Contact = ds.Tables(0).Rows(0)("Contact").ToString()
                Obj.objEnquiryMain.str_RefNo = ds.Tables(0).Rows(0)("RefNo").ToString()
                Obj.objEnquiryMain.str_EnquiryStatus = ds.Tables(0).Rows(0)("EnqStatus").ToString()
                Obj.objEnquiryMain.bool_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString

                Obj.objEnquiryMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objEnquiryMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objEnquiryMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objEnquiryMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objEnquiryMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objEnquiryMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objEnquiryMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objEnquiryMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

                Obj.objEnquiryMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objEnquiryMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objEnquiryMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objEnquiryMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                Obj.objEnquiryMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()

                Obj.objEnquiryMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objEnquiryMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                Obj.objEnquiryMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()

                Obj.objEnquiryMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objEnquiryMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()

            Else
                Obj.objEnquiryMain.dbl_TCAmount = 0
                Obj.objEnquiryMain.dbl_TCDisAmount = 0
                Obj.objEnquiryMain.dbl_TCDiscountAmount = 0
                Obj.objEnquiryMain.dbl_TCMiscAmount = 0
                Obj.objEnquiryMain.dbl_TCMiscPercentage = 0
                Obj.objEnquiryMain.dbl_TCNetAmount = 0
                Obj.objEnquiryMain.dbl_LCNetAmount = 0

                Obj.objEnquiryMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objEnquiryMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objEnquiryMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objEnquiryMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objEnquiryMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                'Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()
                Obj.objEnquiryMain.dtp_EnquiryDate1 = ds.Tables(0).Rows(0)("IndentDate").ToString()
                Obj.objEnquiryMain.dtp_EnquiryDate2 = Date.Now
                Obj.dtp_CreatedDate = Date.Now
                Obj.dtp_LastUpdatedDate = Date.Now
                Obj.dtp_ApprovedDate = Date.Now
                Obj.objEnquiryMain.dtp_IndentDate = ds.Tables(0).Rows(0)("IndentDate").ToString()
                Obj.objEnquiryMain.str_ExpiryDays = ds.Tables(0).Rows(0)("ExpiryDays").ToString()
            End If

            If ds.Tables(1).Rows.Count > 0 Then
                Obj.objEnquirySub.dt_Enquiry = ds.Tables(1)
            End If
            If ds.Tables(2).Rows.Count > 0 Then
                Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
            Else
                Obj.objproject.str_ProjectID = ""
                Obj.objproject.str_ProjectLocation = ""
                Obj.objproject.str_WorkOrderNo = ""
            End If
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_Enquiry(ByVal obj As csEnquiryForm, ByRef EnquiryNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_EnquiryUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objEnquiryMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objEnquiryMain.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objEnquiryMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@EnquiryNo", obj.objEnquiryMain.str_EnquiryNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objEnquiryMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@EnquiryDate1", obj.objEnquiryMain.dtp_EnquiryDate1)
            BaseConn.cmd.Parameters.AddWithValue("@EnquiryDate2", obj.objEnquiryMain.dtp_EnquiryDate1)
            BaseConn.cmd.Parameters.AddWithValue("@IndentNo", obj.objEnquiryMain.str_IndentNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objEnquiryMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objEnquiryMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objEnquiryMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.objEnquiryMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@IndRef", obj.objEnquiryMain.str_IndRef)
            BaseConn.cmd.Parameters.AddWithValue("@DelivAddress", obj.objEnquiryMain.str_DelivAddress)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objEnquiryMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@Contact", obj.objEnquiryMain.str_Contact)
            BaseConn.cmd.Parameters.AddWithValue("@RefNo", obj.objEnquiryMain.str_RefNo)
            BaseConn.cmd.Parameters.AddWithValue("@EnquiryStatus", obj.objEnquiryMain.str_EnquiryStatus)
            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objEnquiryMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objEnquiryMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objEnquiryMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.objEnquiryMain.dbl_TCMiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.objEnquiryMain.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objEnquiryMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objEnquiryMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objEnquiryMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objEnquiryMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objEnquiryMain.bool_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.objEnquiryMain.str_MiscText)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)

            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objEnquiryMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objEnquiryMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objEnquiryMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objEnquiryMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objEnquiryMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objEnquiryMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objEnquiryMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objEnquiryMain.str_Desc8)

            BaseConn.cmd.Parameters.AddWithValue("@EnquiryItemDetailsDT", obj.objEnquirySub.dt_Enquiry)
            BaseConn.cmd.Parameters.Add("@EnquiryNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            EnquiryNo = BaseConn.cmd.Parameters("@EnquiryNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.objEnquiryMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "Enquiry", Err.Number, "Error in " & obj.objEnquiryMain.str_Flag & " : " & obj.objEnquiryMain.str_EnquiryNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Update_Enquiry = _ErrString
    End Function


    Public Sub GetEnquiryDetailsForOrder(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _IndentNo As String, ByVal _OrderStatus As Integer, _
                             ByVal _Flag As String, ByRef _dtEnquiryItems As DataTable, ByRef _dtEnquiryItemsWithPrice As DataTable, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEnquiryDetailsForOrder]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@IndentNo", _IndentNo)
            BaseConn.cmd.Parameters.AddWithValue("@OrderStatus", _OrderStatus)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _dtEnquiryItems = ds.Tables(0)
            _dtEnquiryItemsWithPrice = ds.Tables(1)

        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function EnquiryOrderUpdate(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _BusinessPeriodID As Integer, ByVal _IndentNo As String, _
                               ByVal _OrderStatus As Boolean, ByVal _LastUpdatedBy As String, ByVal _LastUpdatedDate As Date, ByVal _EnquiryOrder As DataTable, _
                               ByRef _ErrNo As Integer, ByRef _ErrDesc As String) As String
        Dim _ErrString As String = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_EnquiryOrderUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@IndentNumber", _IndentNo)
            
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", _LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", _LastUpdatedDate)

            BaseConn.cmd.Parameters.AddWithValue("@EnquiryOrder", _EnquiryOrder)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_SiteID)
            ObjDalGeneral.Elog_Insert(_SiteID, _StrDBPath, _StrDBPwd, _BusinessPeriodID, _LastUpdatedBy, _LastUpdatedDate, "", "Enquiry", Err.Number, "Error in update enquiry " & _IndentNo & " ", ex.Message, 5, 3, 1, _ErrNo)
            _ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        EnquiryOrderUpdate = _ErrString
    End Function
    Public Function EnquiryComparisonUpdate(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _BusinessPeriodID As Integer, ByVal _VouNo As String, ByVal _IndentNo As String, _
                               ByVal _OrderStatus As Boolean, ByVal _LastUpdatedBy As String, ByVal _LastUpdatedDate As Date, ByVal _EnquiryOrder As DataTable, _
                               ByRef _ErrNo As Integer, ByRef _ErrDesc As String) As String
        Dim _ErrString As String = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_EnquiryComparisonUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@IndentNumber", _IndentNo)

            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", _LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", _LastUpdatedDate)

            BaseConn.cmd.Parameters.AddWithValue("@EnquiryOrder", _EnquiryOrder)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_SiteID)
            ObjDalGeneral.Elog_Insert(_SiteID, _StrDBPath, _StrDBPwd, _BusinessPeriodID, _LastUpdatedBy, _LastUpdatedDate, "", "Enquiry", Err.Number, "Error in update enquiry " & _IndentNo & " ", ex.Message, 5, 3, 1, _ErrNo)
            _ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        EnquiryComparisonUpdate = _ErrString
    End Function
End Class
