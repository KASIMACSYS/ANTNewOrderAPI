'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_PackingList

    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csPackingList, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, _
                             ByRef ErrStr As String)

        ErrNo = 0
        ErrStr = ""

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GGBC_GetPackingList]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@PkNo", Obj.ObjPackListSub.str_PkNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjPackListMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjPackListSub.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.ObjPackListSub.str_Flag <> "INVOICE" Then
                Obj.ObjPackListMain.dtp_DODate1 = ds.Tables(0).Rows(0)("DODate1").ToString()
            End If

            Obj.ObjPackListCommon.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()

            Obj.ObjPackListMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()

            Obj.ObjPackListMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
            Obj.ObjPackListMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
            Obj.ObjPackListMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
            Obj.ObjPackListMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()

            Obj.ObjPackListMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjPackListMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.ObjPackListMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

            'Obj.objDOMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            'Obj.objDOMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
            'Obj.objDOMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            'Obj.objDOMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            'Obj.objDOMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
            'Obj.objDOMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
            'Obj.objDOMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()

            Obj.ObjPackListMain.str_DoNo = ds.Tables(0).Rows(0)("DoNo").ToString()
            Obj.ObjPackListMain.str_InvNo = ds.Tables(0).Rows(0)("SISNo").ToString()
            Obj.ObjPackListMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            'Obj.ObjPackListMain.dtp_DODate1 = ds.Tables(0).Rows(0)("DODate1").ToString()
            'Obj.ObjPackListMain.dtp_DoDate2 = ds.Tables(0).Rows(0)("DODate2").ToString()
            'Obj.objDOMain.str_MerchantRef = ds.Tables(0).Rows(0)("MerchantRef").ToString()
            Obj.ObjPackListMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel")
            'Obj.objDOMain.str_SIS = ds.Tables(0).Rows(0)("SISNo").ToString()
            'Obj.objDOMain.dbl_SISAmt = ds.Tables(0).Rows(0)("SISAmount").ToString()

            'Obj.objDOMain.dbl_TotalTax = ds.Tables(0).Rows(0)("TotalTax").ToString()
            'Obj.objDOMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
            'Obj.objDOMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString() 'TODO

            Obj.ObjPackListMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()


            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()


            Obj.ObjPackListMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            Obj.ObjPackListMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            Obj.ObjPackListMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            Obj.ObjPackListMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            Obj.ObjPackListMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
            Obj.ObjPackListMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
            Obj.ObjPackListMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
            Obj.ObjPackListMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()


            'Obj.objDOMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
            'Obj.objDOMain.str_ContactPerson = ds.Tables(0).Rows(0)("ContactPerson").ToString()

            Obj.ObjPackListSub.dt_PackListSub = ds.Tables(1)



        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub


    Public Function Update_packing(ByVal obj As csPackingList, ByRef VouNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0




        Try

            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("GGBC_UpdatePackingList", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.ObjPackListCommon.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjPackListSub.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjPackListSub.MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjPackListMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjPackListMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@DoNo", obj.ObjPackListMain.str_DoNo)
            BaseConn.cmd.Parameters.AddWithValue("@SalOrd", "")
            BaseConn.cmd.Parameters.AddWithValue("@PkNo", obj.ObjPackListMain.str_PkNo)
            BaseConn.cmd.Parameters.AddWithValue("@DODate1", obj.ObjPackListMain.dtp_DODate1)
            BaseConn.cmd.Parameters.AddWithValue("@DODate2", obj.ObjPackListMain.dtp_DoDate2)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.ObjPackListCommon.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.ObjPackListMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.ObjPackListMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.ObjPackListMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantRef", "")
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjPackListMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@SISNo", obj.ObjPackListMain.str_InvNo)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.ObjPackListMain.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManName", obj.ObjPackListMain.str_SalesManName)
            BaseConn.cmd.Parameters.AddWithValue("@DeliveryAddress", "")
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.ObjPackListMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.ObjPackListMain.dbl_ExchangeRate)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", 0)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", 0)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", 0)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", 0)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", 0)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", 0)

            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", 0)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", "")

            BaseConn.cmd.Parameters.AddWithValue("@SISAmount", 0)
            BaseConn.cmd.Parameters.AddWithValue("@TotalTax", 0)

            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.ObjPackListMain.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@PackagingComment", obj.ObjPackListMain.str_PackagingComment)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.ObjPackListMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.ObjPackListMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.ObjPackListMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.ObjPackListMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.ObjPackListMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.ObjPackListMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.ObjPackListMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.ObjPackListMain.str_Desc8)

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

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", "")
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", "")
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", "")

            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.ObjPackListMain.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@ContactPerson", "")

            BaseConn.cmd.Parameters.AddWithValue("@UserComment", "")
            BaseConn.cmd.Parameters.AddWithValue("@ApproverComment", "")

            BaseConn.cmd.Parameters.AddWithValue("@PackingDT", obj.ObjPackListSub.dt_PackListSub)
            'BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

            'Catch ex As Exception
            '    _ErrString = ex.Message
            '    ObjDalGeneral = New DAL_General(obj.str_SiteID)
            '    ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.objDOMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "DO", Err.Number, "Error in " & obj.objDOMain.str_Flag & " : " & obj.objDOMain.str_DoNo & "", ex.Message, 5, 3, 1, ErrNo)
            '    ErrNo = 1
            'Finally
            '    BaseConn.Close()
            'End Try

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.ObjPackListMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "DO", Err.Number, "Error in " & obj.ObjPackListSub.str_Flag & " : " & obj.ObjPackListMain.str_DoNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try



        Update_packing = _ErrString
    End Function


End Class
