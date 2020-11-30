'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csLogistics
    Inherits csSignature
    Public str_SiteID As String
    Public str_Flag As String
    Public str_FormPrefix As String
    Public str_MenuID As String
    Public int_BusinessPerionID As Integer
    Public str_VoucherNo As String
    Public int_RevNo As Integer
    Public dtp_VoucherDate As Date
    Public str_DoNo As String
    Public str_CusName As String
    Public dtp_DoDate As Date
    Public str_DeliveryLocation As String
    Public str_DOCreadtedBy As String
    Public bool_DOStatus As Boolean
    Public str_TruckNo As String
    Public str_DeliverName As String
    Public str_MobileNo As String
    Public dtp_TimePrint As DateTime
    Public dbl_CargoCharges As Decimal
    Public bool_TransportStatus As Boolean
    Public str_GatePass As String
    Public str_CustomRef As String
    Public dbl_CustomCharges As Decimal
    Public str_AirwayBillNo As String
    Public dtp_ExitDate As Date
    Public bool_CustomStatus As Boolean
    Public str_Comment As String

End Class

