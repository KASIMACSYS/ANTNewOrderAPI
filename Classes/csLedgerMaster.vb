'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csLedgerMaster
    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer

    Public int_LedgerID As Integer
    Public str_Description As String
    Public str_Classification As String
    Public str_LedgerType As String
    Public str_ParentAccount As String
    Public str_Class As String
    Public str_StartRange As String
    Public str_EndRange As String
    Public str_AccountNo1 As String
    Public str_AccountNo2 As String
    Public bool_InActive As Boolean
    Public bool_CostCentre As Boolean
    Public bool_Readonly As Boolean = False

    Public dbl_Amount As Double
    Public dbl_Advance As Double
    Public str_Comment As String

    Public str_Flag As String
    Public str_Category As String
End Class
