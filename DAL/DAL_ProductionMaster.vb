Imports Classes

Public Class DAL_ProductionMaster
    Private ObjDalGeneral As DAL_General
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()


    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef Obj As csProductionUnit, ByRef _DTProdDetails As DataTable, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetProductionMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure

            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ProdNo", Obj.str_ProdUnitNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                Obj.str_ProdUnitName = ds.Tables(0).Rows(0)("ProdUnitName").ToString()
                Obj.dtp_Date = ds.Tables(0).Rows(0)("ProdDate").ToString()
                Obj.str_Location = ds.Tables(0).Rows(0)("Location").ToString()
                Obj.str_Contact = ds.Tables(0).Rows(0)("Contact").ToString()
                Obj.str_Telephone = ds.Tables(0).Rows(0)("Telephone").ToString()
                Obj.str_Mobile = ds.Tables(0).Rows(0)("Mobile").ToString()
                Obj.str_Address = ds.Tables(0).Rows(0)("Address").ToString()
            End If
            _DTProdDetails = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_ProductionMaster(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csProductionUnit, ByRef str_ProdUnitNo As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ProductionUnitUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ProdUnitNo", obj.str_ProdUnitNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProdUnitName", obj.str_ProdUnitName)
            BaseConn.cmd.Parameters.AddWithValue("@ProdDate", obj.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@Location", obj.str_Location)
            BaseConn.cmd.Parameters.AddWithValue("@Contact", obj.str_Contact)

            BaseConn.cmd.Parameters.AddWithValue("@Telephone", obj.str_Telephone)
            BaseConn.cmd.Parameters.AddWithValue("@Mobile", obj.str_Mobile)
            BaseConn.cmd.Parameters.AddWithValue("@Address", obj.str_Address)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)

            BaseConn.cmd.Parameters.Add("@ProdNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            str_ProdUnitNo = BaseConn.cmd.Parameters("@ProdNoOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "ProductionMaster", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_ProdUnitName & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_ProductionMaster = _ErrString
    End Function
End Class
