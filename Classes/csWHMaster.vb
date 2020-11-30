'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csWHMaster
    Inherits csSignature
    Public str_SiteID As String
    Public str_WHID As String
    Public str_WHDesc As String
    Public str_Address As String
    Public bool_DefaultWH As Boolean
    Public str_Comment As String
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String
    Public DTItems As DataTable
End Class
