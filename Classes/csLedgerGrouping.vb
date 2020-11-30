'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csLedgerGrouping
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public int_LedgerID As Integer
    Public int_ParentID As Integer
    Public str_Description As String

    Public bool_OnlyControlAC As Boolean
    Public bool_OnlyClassAC As Boolean
    Public bool_OnlyNominalAC As Boolean
    Public bool_ControlACwithClass As Boolean

    Public str_Category As String

    Public str_Flag As String
    Public str_FormType As String

    Public dt_group As New DataTable
    Public dt_LedgerDesc As New DataTable
End Class
