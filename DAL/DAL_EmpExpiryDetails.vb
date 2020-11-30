'======================================================================================
'$Author: Kasim $
'$Rev: 240 $
'$Date: 2013-09-12 20:50:27 +0530 (Thu, 12 Sep 2013) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes
Imports System.Data.SqlClient
Public Class DAL_EmpExpiryDetails
    Dim dt, dt1 As DataTable
    Dim BaseConn As New SQLConn()
    Dim SiteID As String

    Public Sub New(ByVal siteid As String)
        Me.SiteID = siteid
    End Sub

    Public Function GetMasterDetails(ByVal _strPath As String, ByVal _strPwd As String, ByVal str_Flag As String, ByVal _Condition As String, _
                                      ByVal _Category As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEmpExpiryDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Category", _Category)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    'Public Function GetSubTableDetails(ByVal _strPath As String, ByVal _strPwd As String, ByVal str_Flag As String, ByVal Condition As String) As DataTable
    '    Try
    '        dt = New DataTable
    '        BaseConn.Open(_strPath, _strPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEmpExpiryDetails]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@Flag", str_Flag)
    '        BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        dt = ds.Tables(0)
    '    Catch ex As Exception
    '        MsgBox("Error" & ex.Message)
    '    Finally
    '        BaseConn.Close()
    '    End Try
    '    Return dt
    'End Function
End Class
