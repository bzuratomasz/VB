Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Runtime.CompilerServices

Public Module modDefs
    'Globale Konstenten & Objekte
    
    Public Class Konstanten
        Public Class Extensions
            Public Const ZIP As String = ".zip"
            Public Const TXT As String = ".txt"
            Public Const STKL As String = ".stkl"
            Public Const LNK As String = ".lnk"
            Public Const PDF As String = ".pdf"
            Public Const DOC As String = ".doc"
            Public Const DOT As String = ".dot"
            Public Const TIF As String = ".tif"
            Public Const CSV As String = ".csv"
            Public Const MSG As String = ".msg"
            Public Const OFT As String = ".oft"
            Public Const XLS As String = ".xlsx"
            Public Const XML As String = ".xml"
            Public Const DAT As String = ".dat"
            Public Const LOG As String = ".log"
            Public Const BIN As String = ".bin"
            Public Const P As String = ".p"
            Public Const L As String = ".l"
            Public Const ST As String = ".st"
            Public Const DWG As String = ".dwg"
            Public Const VVI As String = ".vvi"
            Public Const VVIEND As String = ".end"
            Public Const DXF As String = ".dxf"
            Public Const JT As String = ".jt"
        End Class

        Public Class ExtensionsOhnePunkt
            Public Const ZIP As String = "zip"
            Public Const TXT As String = "txt"
            Public Const STKL As String = "stkl"
            Public Const LNK As String = "lnk"
            Public Const PDF As String = "pdf"
            Public Const DOC As String = "doc"
            Public Const DOT As String = "dot"
            Public Const TIF As String = "tif"
            Public Const CSV As String = "csv"
            Public Const MSG As String = "msg"
            Public Const OFT As String = "oft"
            Public Const XLS As String = "xlsx"
            Public Const XML As String = "xml"
            Public Const DAT As String = "dat"
            Public Const LOG As String = "log"
            Public Const BIN As String = "bin"
            Public Const P As String = "p"
            Public Const L As String = "l"
            Public Const ST As String = "st"
            Public Const DWG As String = "dwg"
            Public Const VVI As String = "vvi"
            Public Const VVIEND As String = "end"
            Public Const DXF As String = "dxf"
            Public Const JT As String = "jt"
        End Class

        Public Class Objects
            Public Const Excel As String = "Excel.Application"
            Public Const Outlook As String = "Outlook.Application"
            Public Const Word As String = "Word.Application"
        End Class

        Public Class Crypt
            Public Const String4DB As String = "§SE_WCrypt4DB&"
        End Class
    End Class

    '----
    Public Enum EnumRecordSetMode As Integer
        rmAddNew
        rmEdit
        rmChanged
        rmDelte
    End Enum

    Public Enum EnumAktion As Integer
        AktionNew = 1
        AktionEdit = 2
        AktionReadOnly = 3
    End Enum

    Public Enum EnumYesNo As Integer
        Yes = 1
        No = 2
    End Enum
    Public Function PrintEnumYesNo(opt As EnumYesNo) As String
        Select Case opt
            Case EnumYesNo.Yes : Return My.Resources.resGlobal.TextYes
            Case EnumYesNo.No : Return My.Resources.resGlobal.TextNo
            Case Else : Return "???"
        End Select
    End Function
    Public Sub FillEnumYesNo2Cmb(cmb As ComboBox, Optional withAll As Boolean = False)
        cmb.Items.Clear()
        If withAll Then cmb.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextAll, 0))
        cmb.Items.Add(New clsListBoxItem(PrintEnumYesNo(EnumYesNo.Yes), EnumYesNo.Yes))
        cmb.Items.Add(New clsListBoxItem(PrintEnumYesNo(EnumYesNo.No), EnumYesNo.No))
    End Sub

    Public Enum EnumUserType As Integer
        Administrator = 1
        User = 2
    End Enum
    Public Function PrintEnumUserType(type As EnumUserType) As String
        Select Case type
            Case EnumUserType.User : Return My.Resources.resGlobal.TextUser
                'Case EnumUserType.Supervisor : Return My.Resources.resMain.TextSupervisor
            Case EnumUserType.Administrator : Return My.Resources.resGlobal.TextAdministrator
            Case Else : Return "???"
        End Select
    End Function
    Public Sub FillEnumUserType2Cmb(cmb As ComboBox, Optional withAll As Boolean = False)
        cmb.Items.Clear()
        If withAll Then cmb.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextAll, 0))
        cmb.Items.Add(New clsListBoxItem(PrintEnumUserType(EnumUserType.User), EnumUserType.User))
        'cmb.Items.Add(New clsListBoxItem(PrintEnumUserType(EnumUserType.Supervisor), EnumUserType.Supervisor))
        cmb.Items.Add(New clsListBoxItem(PrintEnumUserType(EnumUserType.Administrator), EnumUserType.Administrator))
    End Sub

    Public Enum EnumCmbFlags
        WithAll = 1 '          00000001
        WithPleaseSelect = 2 ' 00000010
        Ordered = 4 '          00000100
        WithNoChoice = 8 '     00001000
    End Enum

End Module
