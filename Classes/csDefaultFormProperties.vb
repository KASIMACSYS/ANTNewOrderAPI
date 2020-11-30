'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports System.Data

Public Class csDefaultFormProperties
    Public str_SiteID As String
    Public str_BusinessPerionID As Integer
    Public dt As New DataTable
    Public MenuID As String
    Public MenuName As String
    Public dt_DefaultForm As DataTable
    Public dt4 As New DataTable
    Public objproject As csProjectDetail
    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        If CustomerSetting.Item("Project").ToString = "True" Then
            objProject = New csProjectDetail
        End If
    End Sub
End Class
