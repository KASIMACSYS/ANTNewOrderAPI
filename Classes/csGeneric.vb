'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csSignature

    Public str_CreatedBy As String = String.Empty
    Public dtp_CreatedDate As Date = Date.Now
    Public str_LastUpdatedBy As String = String.Empty
    Public dtp_LastUpdatedDate As Date = Date.Now
    Public str_ApprovedBy As String = String.Empty
    Public dtp_ApprovedDate As Date = Date.Now
    Public bool_ApprovedStatus As Integer
    Public int_ApprovedLevel As Integer
    Public str_UserComment As String = String.Empty
    Public str_ApprovedComment As String = String.Empty
    Public bool_ApprovedHigherLevel As Boolean

    Public Property CreatedBy() As String
        Get
            Return str_CreatedBy
        End Get
        Set(ByVal value As String)
            str_CreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return dtp_CreatedDate
        End Get
        Set(ByVal value As Date)
            dtp_CreatedDate = value
        End Set
    End Property

    Public Property LastUpdatedBy() As String
        Get
            Return str_LastUpdatedBy
        End Get
        Set(ByVal value As String)
            str_LastUpdatedBy = value
        End Set
    End Property

    Public Property LastUpdatedDate() As Date
        Get
            Return dtp_LastUpdatedDate
        End Get
        Set(ByVal value As Date)
            dtp_LastUpdatedDate = value
        End Set
    End Property

    Public Property ApprovedBy() As String
        Get
            Return str_ApprovedBy
        End Get
        Set(ByVal value As String)
            str_ApprovedBy = value
        End Set
    End Property

    Public Property ApprovedDate() As Date
        Get
            Return dtp_ApprovedDate
        End Get
        Set(ByVal value As Date)
            dtp_ApprovedDate = value
        End Set
    End Property

    Public Property ApprovedStatus() As Integer
        Get
            Return bool_ApprovedStatus
        End Get
        Set(ByVal value As Integer)
            bool_ApprovedStatus = value
        End Set
    End Property

    Public Property ApprovedLevel() As Integer
        Get
            Return int_ApprovedLevel
        End Get
        Set(ByVal value As Integer)
            int_ApprovedLevel = value
        End Set
    End Property

    Public Property UserComment() As String
        Get
            Return str_UserComment
        End Get
        Set(ByVal value As String)
            str_UserComment = value
        End Set
    End Property

    Public Property ApprovedComment() As String
        Get
            Return str_ApprovedComment
        End Get
        Set(ByVal value As String)
            str_ApprovedComment = value
        End Set
    End Property

    Public Property ApprovedHigherLevel() As Boolean
        Get
            Return bool_ApprovedHigherLevel
        End Get
        Set(ByVal value As Boolean)
            bool_ApprovedHigherLevel = value
        End Set
    End Property

    'Public Function Get_Signature(ByVal strCreatedBy As String, ByVal dtpCreatedDate As Date, ByVal strLastUpdatedBy As String _
    '                              , ByVal dtpLastUpdatedDate As Date, ByVal strApprovedBy As String, ByVal dtpApprovedDate As Date, ByVal boolApprovedStatus As Boolean) As csSignature
    '    Dim objSignature As New csSignature
    '    objSignature.CreatedBy = strCreatedBy
    '    objSignature.CreatedDate = dtpCreatedDate
    '    objSignature.LastUpdatedBy = strLastUpdatedBy
    '    objSignature.LastUpdatedDate = dtpLastUpdatedDate
    '    objSignature.ApprovedBy = strApprovedBy
    '    objSignature.ApprovedDate = dtpApprovedDate
    '    objSignature.ApprovedStatus = boolApprovedStatus
    '    Return objSignature
    'End Function


End Class


Public Class FormDefaults
    Public VouPrefix As String = String.Empty
    Public IsVouFrcFlag As Boolean = False
    Public IsVouApproveFlag As Boolean = False

End Class

Public Class csCustomerDetails
    Public str_MerchantName As String
    Public str_MerchantID As String
    Public str_Contact As String
    Public str_Address As String
    Public str_Tel As String
    Public str_Mobile As String
    Public str_Aging As String
End Class


