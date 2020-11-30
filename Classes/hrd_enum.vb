'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class hrd_enum

    Public Enum ShortcutEvnet As Integer
        CtrlA = 1
        ctrlS = 19
        CtrlE = 5
        CtrlD = 4
        CtrlQ = 17
        CtrlV = 22
        CtrlP = 16
        CtrlC = 10 'change
    End Enum

    Public Enum LimitStatus As Integer
        NoAction = 0
        CannotContinue = 1
        NotifyAndContinue = 2
        PermissionToContinue = 3
    End Enum

    Public Enum ItemDefaultPrice As Integer
        P
        H
        M
        L
    End Enum
    'Public Property FindShortCut() As ShortcutEvnet
    '    Get
    '        Return ShortCutValue
    '    End Get
    '    Set(ByVal value As ShortcutEvnet)
    '        ShortCutValue = value
    '    End Set
    'End Property
End Class

Public Class ConstantLedgerName
    Public Const Income As Integer = 4
    Public Const DirectIncome As Integer = 172
    Public Const IndirectIncome As Integer = 177

    Public Const Expenses As Integer = 3
    Public Const DirectExpense As Integer = 140
    Public Const IndirectExpense As Integer = 153

    Public Const COGS As Integer = 198
End Class