'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_RV
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csRV, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetRVDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.str_MenuID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString
            Obj.str_VouRef = ds.Tables(0).Rows(0)("VouRef").ToString
            Obj.str_LedgerType = ds.Tables(0).Rows(0)("LedgerType").ToString()
            Obj.dtp_RVDate = ds.Tables(0).Rows(0)("RVDate").ToString()
            Obj.str_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
            Obj.str_Alice = ds.Tables(0).Rows(0)("Alias").ToString()
            Obj.str_DstLedgerID = ds.Tables(0).Rows(0)("DstLedgerID").ToString()
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.str_VouForHRMode = ds.Tables(0).Rows(0)("VouForHR").ToString()

            Obj.dbl_TCTotalAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            Obj.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
            Obj.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            Obj.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
            Obj.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            Obj.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
            Obj.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

            If ds.Tables(0).Rows(0)("ApprovedStatus") = True Then 'Temp fix
                Obj.bool_ApprovedStatus = 1
            ElseIf ds.Tables(0).Rows(0)("ApprovedStatus") = False Then
                Obj.bool_ApprovedStatus = 0
            Else
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            End If

            Obj.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel")
            Obj.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()

            Obj.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            Obj.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            Obj.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            Obj.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            Obj.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
            Obj.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
            Obj.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
            Obj.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_RVCash(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal objcsRV As csRV, ByRef _VouNoOut As String, ByRef int_RevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            If objcsRV.str_LedgerType = "MERCHANT" Or objcsRV.str_LedgerType = "GENERAL" Then
                BaseConn.cmd = New SqlClient.SqlCommand("RVCashUpdate", BaseConn.cnn)
            ElseIf objcsRV.str_LedgerType = "EMPLOYEE" Then
                BaseConn.cmd = New SqlClient.SqlCommand("RVCashUpdateHR", BaseConn.cnn)
            End If

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", objcsRV.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", objcsRV.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", objcsRV.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", objcsRV.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", objcsRV.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@VouNo", objcsRV.str_VouNo) ' "RV/1001") '
            BaseConn.cmd.Parameters.AddWithValue("@VouRef", objcsRV.str_VouRef)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", objcsRV.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerType", objcsRV.str_LedgerType)
            BaseConn.cmd.Parameters.AddWithValue("@RVDate", objcsRV.dtp_RVDate)
            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", objcsRV.str_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", objcsRV.str_Alice)
            BaseConn.cmd.Parameters.AddWithValue("@DstLedgerID", objcsRV.str_DstLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@PayType", objcsRV.str_PayType)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", objcsRV.str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", objcsRV.dbl_TCTotalAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", objcsRV.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", objcsRV.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", objcsRV.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", objcsRV.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", objcsRV.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", objcsRV.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", objcsRV.dbl_ExchangeRate)
            'BaseConn.cmd.Parameters.AddWithValue("@COA", objcsRV.str_CashLedger)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", objcsRV.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", objcsRV.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", objcsRV.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", objcsRV.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", objcsRV.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", objcsRV.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", objcsRV.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", objcsRV.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", objcsRV.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", objcsRV.ApprovedHigherLevel)
            BaseConn.cmd.Parameters.AddWithValue("@VouForHR", objcsRV.str_VouForHRMode)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", objcsRV.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", objcsRV.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", objcsRV.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", objcsRV.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", objcsRV.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", objcsRV.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", objcsRV.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", objcsRV.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", objcsRV.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", objcsRV.str_Desc8)

            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", objcsRV.int_LanguageCode)

            If objcsRV.str_LedgerType = "MERCHANT" Then
                BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", objcsRV.dt_RVMatching)
            ElseIf objcsRV.str_LedgerType = "EMPLOYEE" Then
                BaseConn.cmd.Parameters.AddWithValue("@AdvanceDT", objcsRV.DT_Wages)
            End If

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _VouNoOut = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            int_RevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(objcsRV.int_CID)
            ObjDalGeneral.Elog_Insert(objcsRV.int_CID, _DBPath, _DBPwd, objcsRV.int_BusinessPeriodID, objcsRV.str_CreatedBy, objcsRV.dtp_CreatedDate, "", "RVCash", Err.Number, "Error in " & objcsRV.str_Flag & " : " & objcsRV.str_VouNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Update_RVCash = _ErrString
    End Function

End Class


Public Class DAL_RVCheque
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef dt_ChqDet As DataTable, ByRef dt_VouMatDet As DataTable, ByRef ObjcsRV As csRV, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetRVChequeDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", ObjcsRV.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", ObjcsRV.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", ObjcsRV.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", ObjcsRV.str_MenuID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            ObjcsRV.str_VouRef = ds.Tables(0).Rows(0)("VouRef").ToString
            ObjcsRV.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString
            ObjcsRV.str_LedgerType = ds.Tables(0).Rows(0)("LedgerType").ToString()
            ObjcsRV.str_VouForHRMode = ds.Tables(0).Rows(0)("VouForHR").ToString()

            ObjcsRV.dtp_RVDate = ds.Tables(0).Rows(0)("RVDate").ToString()
            ObjcsRV.str_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
            ObjcsRV.str_Alice = ds.Tables(0).Rows(0)("Alias").ToString()
            ObjcsRV.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            ObjcsRV.dbl_TCTotalAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            ObjcsRV.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
            ObjcsRV.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            ObjcsRV.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()

            ObjcsRV.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            ObjcsRV.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
            ObjcsRV.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            ObjcsRV.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
            If ds.Tables(0).Rows(0)("ApprovedStatus") = True Then ' Temp fix
                ObjcsRV.bool_ApprovedStatus = 1
            ElseIf ds.Tables(0).Rows(0)("ApprovedStatus") = False Then
                ObjcsRV.bool_ApprovedStatus = 0
            Else
                ObjcsRV.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            End If
            ObjcsRV.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel")
            ObjcsRV.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()

            ObjcsRV.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            ObjcsRV.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            ObjcsRV.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            ObjcsRV.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            ObjcsRV.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
            ObjcsRV.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
            ObjcsRV.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
            ObjcsRV.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

            dt_ChqDet = ds.Tables(1)
            dt_VouMatDet = ds.Tables(2)
            'dt_getClosedVoucherDT = ds.Tables(3)


        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function Update_RVCheque(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal objcsRVCheque As csRVCheque, ByVal dt_ChqDetails As DataTable, ByVal dt_VouMatching As DataTable, ByRef _VouNoOut As String, ByRef int_RevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            If objcsRVCheque.objRV.str_LedgerType = "MERCHANT" Or objcsRVCheque.objRV.str_LedgerType = "GENERAL" Then
                BaseConn.cmd = New SqlClient.SqlCommand("RVChequeUpdate", BaseConn.cnn)
            ElseIf objcsRVCheque.objRV.str_LedgerType = "EMPLOYEE" Then
                BaseConn.cmd = New SqlClient.SqlCommand("RVChequeUpdateHR", BaseConn.cnn)
            End If

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", objcsRVCheque.objRV.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", objcsRVCheque.objRV.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", objcsRVCheque.objRV.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", objcsRVCheque.objRV.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", objcsRVCheque.objRV.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@VouNo", objcsRVCheque.objRV.str_VouNo) ' "RV/1001") '
            BaseConn.cmd.Parameters.AddWithValue("@VouRef", objcsRVCheque.objRV.str_VouRef)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", objcsRVCheque.objRV.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", objcsRVCheque.objBankCheque.VouType)
            BaseConn.cmd.Parameters.AddWithValue("@RVDate", objcsRVCheque.objRV.dtp_RVDate)
            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", objcsRVCheque.objRV.str_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", objcsRVCheque.objRV.str_Alice)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", objcsRVCheque.objRV.str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", objcsRVCheque.objRV.dbl_TCTotalAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", objcsRVCheque.objRV.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", objcsRVCheque.objRV.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", objcsRVCheque.objRV.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", objcsRVCheque.objRV.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", objcsRVCheque.objRV.dbl_LCNetAmount)


            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", objcsRVCheque.objRV.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", objcsRVCheque.objRV.dbl_ExchangeRate)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", objcsRVCheque.objRV.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", objcsRVCheque.objRV.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", objcsRVCheque.objRV.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", objcsRVCheque.objRV.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", objcsRVCheque.objRV.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", objcsRVCheque.objRV.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", objcsRVCheque.objRV.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", objcsRVCheque.objRV.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", objcsRVCheque.objRV.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", objcsRVCheque.objRV.ApprovedHigherLevel)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerType", objcsRVCheque.objRV.str_LedgerType)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", objcsRVCheque.objRV.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancelPrevious", objcsRVCheque.objRV.int_StatusCancelPrevious)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", objcsRVCheque.objRV.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", objcsRVCheque.objRV.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", objcsRVCheque.objRV.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", objcsRVCheque.objRV.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", objcsRVCheque.objRV.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", objcsRVCheque.objRV.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", objcsRVCheque.objRV.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", objcsRVCheque.objRV.str_Desc8)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", objcsRVCheque.objRV.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@ChequeDT", dt_ChqDetails)

            If objcsRVCheque.objRV.str_LedgerType = "MERCHANT" Then
                BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", dt_VouMatching)
            ElseIf objcsRVCheque.objRV.str_LedgerType = "EMPLOYEE" Then
                BaseConn.cmd.Parameters.AddWithValue("@AdvanceDT", objcsRVCheque.objRV.DT_Wages)
                BaseConn.cmd.Parameters.AddWithValue("@VouForHR", objcsRVCheque.objRV.str_VouForHRMode)
            End If


            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _VouNoOut = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            int_RevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(objcsRVCheque.objRV.int_CID)
            ObjDalGeneral.Elog_Insert(objcsRVCheque.objRV.int_CID, _DBPath, _DBPwd, objcsRVCheque.objRV.int_BusinessPeriodID, objcsRVCheque.objRV.str_CreatedBy, objcsRVCheque.objRV.dtp_CreatedDate, "", "RVCheque", Err.Number, "Error in " & objcsRVCheque.objRV.str_Flag & " : " & objcsRVCheque.objRV.str_VouNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Update_RVCheque = _ErrString
    End Function
End Class