Imports Classes

Public Class DAL_LandingCost

    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    'Public Function GetMatchedLC(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As String, ByVal _LCNo As String, ByRef ErrNo As Integer, ByVal ErrStr As String) As Integer
    '    Dim IsMatched As Integer = 0
    '    Try
    '        BaseConn.Open(_DBPath, _DBPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[IsLCMatched]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@LCNo", _LCNo)
    '        BaseConn.cmd.Parameters.Add("@IsMatched", SqlDbType.Int).Direction = ParameterDirection.Output

    '        BaseConn.cmd.ExecuteNonQuery()
    '        IsMatched = BaseConn.cmd.Parameters("@IsMatched").Value

    '    Catch ex As Exception
    '        ErrStr = ex.Message
    '        ErrNo = 0
    '    Finally
    '        BaseConn.Close()
    '    End Try

    '    Return IsMatched
    'End Function

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef _DTLCTypes As DataTable, ByRef Obj As csLandingCost, ByRef ErrNo As Integer, ByVal ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetLandingCostDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objLCMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LCVou", Obj.objLCMain.str_LCNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objLCMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTLCTypes = ds.Tables(0)

            If Obj.objLCMain.str_Flag = "EDIT" Then
                Obj.objLCMain.int_BusinessPeriodID = ds.Tables(1).Rows(0)("BusinessPeriodID").ToString()
                Obj.objLCMain.int_RevNo = ds.Tables(1).Rows(0)("RevNo").ToString()
                Obj.objLCMain.dtp_LCDate = ds.Tables(1).Rows(0)("LCDate").ToString()
                Obj.objLCMain.str_TCCurrency = ds.Tables(1).Rows(0)("TCCurrency").ToString()
                Obj.objLCMain.dbl_ExchangeRate = ds.Tables(1).Rows(0)("ExchangeRate").ToString()
                Obj.objLCMain.bool_AffectInventoryCost = ds.Tables(1).Rows(0)("AffectInventoryCost").ToString()
                Obj.objLCMain.str_Comment = ds.Tables(1).Rows(0)("Comment").ToString()
                Obj.objLCMain.dbl_TCAmount = ds.Tables(1).Rows(0)("TCAmount").ToString()
                Obj.objLCMain.dbl_TCTaxAmount = ds.Tables(1).Rows(0)("TCVATAmount").ToString()
                Obj.objLCMain.dbl_TCItemTaxAmount = ds.Tables(1).Rows(0)("TCItemTaxAmount").ToString()
                Obj.objLCMain.dbl_TCInvoiceTaxAmount = ds.Tables(1).Rows(0)("TCInvTaxAmount").ToString()
                Obj.objLCMain.dbl_TCNetAmount = ds.Tables(1).Rows(0)("TCNetAmount").ToString()
                Obj.objLCMain.dbl_LCNetAmount = ds.Tables(1).Rows(0)("LCNetAmount").ToString()
                Obj.objLCMain.str_VouType = ds.Tables(1).Rows(0)("VouType").ToString()
                Obj.objLCMain.bool_TaxFileReturn = ds.Tables(1).Rows(0)("TaxReturnFiled").ToString()
                Obj.objLCMain.str_ItemTaxCode = ds.Tables(1).Rows(0)("ItemTaxCode")

                Obj.str_CreatedBy = ds.Tables(1).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(1).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(1).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(1).Rows(0)("LastUpdatedDate").ToString()
                Obj.str_ApprovedBy = ds.Tables(1).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(1).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(1).Rows(0)("ApprovedStatus")

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objLCSub.dt_LCSub = ds.Tables(2)
                End If

                If ds.Tables.Count > 3 Then
                    If ds.Tables(3).Rows.Count > 0 Then
                        Obj.objLCSub.dt_LCItemDetails = ds.Tables(3)
                    End If
                End If

                If ds.Tables.Count > 4 Then
                    Obj.objLCMain.dt_TaxItemDetails = ds.Tables(4)
                End If
            End If
        Catch ex As Exception
            ErrNo = 0
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetVoucherItemDetailsForLC(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As String, ByVal _BusinessPeriodID As Integer,
                                               ByVal _VouType As String, ByVal _DTVou As DataTable, ByRef ErrNo As Integer, ByVal ErrStr As String) As DataTable
        GetVoucherItemDetailsForLC = New DataTable
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetVoucherItemDetailsForLC]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", _VouType)
            BaseConn.cmd.Parameters.AddWithValue("@VouNoDT", _DTVou)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                GetVoucherItemDetailsForLC = ds.Tables(0)
            End If
        Catch ex As Exception
            ErrNo = 0
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
        Return GetVoucherItemDetailsForLC
    End Function

    Public Function Update_LandingCost(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csLandingCost, ByRef LCNo As String,
                                  ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("LandingCostUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objLCMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objLCMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objLCMain.str_Prefix)

            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objLCMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LCVou", obj.objLCMain.str_LCNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objLCMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@LCDate", obj.objLCMain.dtp_LCDate)
            'BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.objLCMain.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objLCMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objLCMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@AffectInventoryCost", obj.objLCMain.bool_AffectInventoryCost)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objLCMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objLCMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCVatAmount", obj.objLCMain.dbl_TCTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objLCMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvTaxAmount", obj.objLCMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objLCMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objLCMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", obj.objLCMain.str_VouType)
            'BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objLCMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objLCMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objLCMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objLCMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objLCMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objLCMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objLCMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objLCMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objLCMain.str_Desc8)

            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objLCMain.str_ItemTaxCode)

            BaseConn.cmd.Parameters.AddWithValue("@LCSubDT", obj.objLCSub.dt_LCSub)
            BaseConn.cmd.Parameters.AddWithValue("@LCItemDetailsDT", obj.objLCSub.dt_LCItemDetails)
            BaseConn.cmd.Parameters.AddWithValue("@LCItemSplitValueDT", obj.objLCSub.dt_LCItemSplitValue)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objLCMain.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.NVarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.NVarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            LCNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _strPath, _strPwd, obj.objLCMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "GIP", Err.Number, "Error in " & obj.objLCMain.str_Flag & " : " & obj.objLCMain.str_LCNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Update_LandingCost = _ErrString
    End Function
End Class
