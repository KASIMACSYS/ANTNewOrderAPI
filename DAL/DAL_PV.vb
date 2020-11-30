'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_PV
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csPV, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPVDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.str_VouNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString
            Obj.str_VouRef = ds.Tables(0).Rows(0)("VouRef").ToString
            Obj.str_LedgerType = ds.Tables(0).Rows(0)("LedgerType").ToString()
            Obj.dtp_PVDate = ds.Tables(0).Rows(0)("PVDate").ToString()
            Obj.str_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
            Obj.str_Alice = ds.Tables(0).Rows(0)("Alias").ToString()
            Obj.str_DstLedgerID = ds.Tables(0).Rows(0)("DstLedgerID").ToString()
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.str_PVForHRMode = ds.Tables(0).Rows(0)("VouForHR").ToString()
            Obj.str_LedgerDepartment = ds.Tables(0).Rows(0)("LedgerDepartment").ToString()

            Obj.dbl_TCTotalAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            Obj.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
            Obj.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            Obj.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
            Obj.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            Obj.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
            Obj.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
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
            ErrNo = 0
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function Update_PVCash(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef _VouNoOut As String, ByRef int_RevNo As Integer, ByVal objcsPV As csPV, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            If objcsPV.str_LedgerType = "MERCHANT" Or objcsPV.str_LedgerType = "GENERAL" Then
                BaseConn.cmd = New SqlClient.SqlCommand("PVCashUpdate", BaseConn.cnn)
            ElseIf objcsPV.str_LedgerType = "EMPLOYEE" Then
                BaseConn.cmd = New SqlClient.SqlCommand("PVCashUpdateHR", BaseConn.cnn)
            End If

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", objcsPV.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", objcsPV.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", objcsPV.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", objcsPV.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", objcsPV.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@VouNo", objcsPV.str_VouNo) ' "PV/1001") '
            BaseConn.cmd.Parameters.AddWithValue("@VouRef", objcsPV.str_VouRef)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", objcsPV.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerType", objcsPV.str_LedgerType)
            BaseConn.cmd.Parameters.AddWithValue("@PVDate", objcsPV.dtp_PVDate)
            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", objcsPV.str_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", objcsPV.str_Alice)
            BaseConn.cmd.Parameters.AddWithValue("@DstLedgerID", objcsPV.str_DstLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@PayType", objcsPV.str_PayType)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", objcsPV.str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", objcsPV.dbl_TCTotalAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", objcsPV.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", objcsPV.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", objcsPV.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", objcsPV.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", objcsPV.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", objcsPV.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", objcsPV.dbl_ExchangeRate)

            BaseConn.cmd.Parameters.AddWithValue("@LedgerDepartment", objcsPV.str_LedgerDepartment)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", objcsPV.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", objcsPV.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", objcsPV.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", objcsPV.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", objcsPV.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", objcsPV.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", objcsPV.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", objcsPV.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", objcsPV.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", objcsPV.ApprovedHigherLevel)
            BaseConn.cmd.Parameters.AddWithValue("@VouForHR", objcsPV.str_PVForHRMode)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", objcsPV.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", objcsPV.str_UserComment)

            BaseConn.cmd.Parameters.AddWithValue("@Desc1", objcsPV.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", objcsPV.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", objcsPV.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", objcsPV.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", objcsPV.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", objcsPV.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", objcsPV.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", objcsPV.str_Desc8)

            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", objcsPV.int_LanguageCode)

            If objcsPV.str_LedgerType = "MERCHANT" Or objcsPV.str_LedgerType = "GENERAL" Then
                BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", objcsPV.dt_PVMatching)
            ElseIf objcsPV.str_LedgerType = "EMPLOYEE" Then
                BaseConn.cmd.Parameters.AddWithValue("@HRPaySlipMainDT", objcsPV.DT_PVHRaySlips)
                BaseConn.cmd.Parameters.AddWithValue("@AdvanceDT", objcsPV.DT_Wages)
            End If

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _VouNoOut = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            int_RevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(objcsPV.int_CID)
            ObjDalGeneral.Elog_Insert(objcsPV.int_CID, _DBPath, _DBPwd, objcsPV.int_BusinessPeriodID, objcsPV.str_CreatedBy, objcsPV.dtp_CreatedDate, "", "PVCash", Err.Number, "Error in " & objcsPV.str_Flag & " : " & objcsPV.str_VouNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Update_PVCash = _ErrString
    End Function

End Class


Public Class DAL_PVCheque
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef dt_ChqDet As DataTable, ByRef dt_VouMatDet As DataTable, ByRef ObjcsPV As csPV, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPVChequeDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", ObjcsPV.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", ObjcsPV.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", ObjcsPV.str_VouNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            ObjcsPV.str_VouRef = ds.Tables(0).Rows(0)("VouRef").ToString
            ObjcsPV.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString
            ObjcsPV.str_LedgerType = ds.Tables(0).Rows(0)("LedgerType").ToString()
            ObjcsPV.dtp_PVDate = ds.Tables(0).Rows(0)("PVDate").ToString()
            ObjcsPV.str_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
            ObjcsPV.str_Alice = ds.Tables(0).Rows(0)("Alias").ToString()
            ObjcsPV.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            ObjcsPV.str_PCNo = ds.Tables(0).Rows(0)("PCNo").ToString
            ObjcsPV.str_PVForHRMode = ds.Tables(0).Rows(0)("VouForHR").ToString()

            ObjcsPV.dbl_TCTotalAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            ObjcsPV.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
            ObjcsPV.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            ObjcsPV.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()

            ObjcsPV.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            ObjcsPV.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
            ObjcsPV.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            ObjcsPV.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
            ObjcsPV.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            ObjcsPV.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel")
            ObjcsPV.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()

            ObjcsPV.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            ObjcsPV.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            ObjcsPV.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            ObjcsPV.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            ObjcsPV.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
            ObjcsPV.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
            ObjcsPV.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
            ObjcsPV.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()


            dt_ChqDet = ds.Tables(1)
            dt_VouMatDet = ds.Tables(2)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function Update_PVCheque(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal objcsPVCheque As csPVCheque, ByVal dt_ChqDetails As DataTable, ByVal dt_VouMatching As DataTable,
                                    ByRef _VouNoOut As String, ByRef int_RevNo As Integer, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            If objcsPVCheque.objPV.str_LedgerType = "MERCHANT" Or objcsPVCheque.objPV.str_LedgerType = "GENERAL" Then
                BaseConn.cmd = New SqlClient.SqlCommand("PVChequeUpdate", BaseConn.cnn)
            ElseIf objcsPVCheque.objPV.str_LedgerType = "EMPLOYEE" Then
                BaseConn.cmd = New SqlClient.SqlCommand("PVChequeUpdateHR", BaseConn.cnn)
            End If

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", objcsPVCheque.objPV.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", objcsPVCheque.objPV.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", objcsPVCheque.objPV.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", objcsPVCheque.objPV.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", objcsPVCheque.objPV.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@VouNo", objcsPVCheque.objPV.str_VouNo) ' "PV/1001") '
            BaseConn.cmd.Parameters.AddWithValue("@VouRef", objcsPVCheque.objPV.str_VouRef)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", objcsPVCheque.objPV.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", objcsPVCheque.objBankCheque.VouType)
            BaseConn.cmd.Parameters.AddWithValue("@PVDate", objcsPVCheque.objPV.dtp_PVDate)
            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", objcsPVCheque.objPV.str_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", objcsPVCheque.objPV.str_Alice)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", objcsPVCheque.objPV.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@PCNo", objcsPVCheque.objPV.str_PCNo)


            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", objcsPVCheque.objPV.dbl_TCTotalAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", objcsPVCheque.objPV.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", objcsPVCheque.objPV.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", objcsPVCheque.objPV.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", objcsPVCheque.objPV.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", objcsPVCheque.objPV.dbl_LCNetAmount)


            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", objcsPVCheque.objPV.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", objcsPVCheque.objPV.dbl_ExchangeRate)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", objcsPVCheque.objPV.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", objcsPVCheque.objPV.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", objcsPVCheque.objPV.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", objcsPVCheque.objPV.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", objcsPVCheque.objPV.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", objcsPVCheque.objPV.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", objcsPVCheque.objPV.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", objcsPVCheque.objPV.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", objcsPVCheque.objPV.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", objcsPVCheque.objPV.ApprovedHigherLevel)
            BaseConn.cmd.Parameters.AddWithValue("@VouForHR", objcsPVCheque.objPV.str_PVForHRMode)


            BaseConn.cmd.Parameters.AddWithValue("@Desc1", objcsPVCheque.objPV.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", objcsPVCheque.objPV.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", objcsPVCheque.objPV.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", objcsPVCheque.objPV.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", objcsPVCheque.objPV.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", objcsPVCheque.objPV.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", objcsPVCheque.objPV.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", objcsPVCheque.objPV.str_Desc8)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", objcsPVCheque.objPV.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerType", objcsPVCheque.objPV.str_LedgerType)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", objcsPVCheque.objPV.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancelPrevious", objcsPVCheque.objPV.int_StatusCancelPrevious)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", objcsPVCheque.objPV.int_LanguageCode)

            BaseConn.cmd.Parameters.AddWithValue("@ChequeDT", dt_ChqDetails)
            If objcsPVCheque.objPV.str_LedgerType = "MERCHANT" Or objcsPVCheque.objPV.str_LedgerType = "GENERAL" Then
                BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", dt_VouMatching)
            ElseIf objcsPVCheque.objPV.str_LedgerType = "EMPLOYEE" Then
                BaseConn.cmd.Parameters.AddWithValue("@HRPaySlipMainDT", objcsPVCheque.objPV.DT_PVHRaySlips)
                BaseConn.cmd.Parameters.AddWithValue("@AdvanceDT", objcsPVCheque.objPV.DT_Wages)
            End If

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _VouNoOut = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            int_RevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString

            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(objcsPVCheque.objPV.int_CID)
            ObjDalGeneral.Elog_Insert(objcsPVCheque.objPV.int_CID, _DBPath, _DBPwd, objcsPVCheque.objPV.int_BusinessPeriodID, objcsPVCheque.objPV.str_CreatedBy, objcsPVCheque.objPV.dtp_CreatedDate, "", "PVCheque", Err.Number, "Error in " & objcsPVCheque.objPV.str_Flag & " : " & objcsPVCheque.objPV.str_VouNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Update_PVCheque = _ErrString
    End Function
End Class