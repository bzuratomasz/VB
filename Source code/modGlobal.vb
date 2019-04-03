Imports System.IO
Imports System.IO.Compression
Imports System.Collections.Generic
Imports System.Runtime.CompilerServices

Module modGlobal

    Public Function Obj2Lng(ByVal obj As Object) As Integer
        Dim ret As Integer = 0
        If obj IsNot Nothing Then
            If TypeOf (obj) Is String AndAlso obj = "" Then Return ret
            Try
                ret = CType(obj, Integer)
            Catch ex As Exception
                'nix
            End Try
        End If
        Return ret
    End Function

    Public Sub Pflichtfeld(ByRef l As Control, ByRef t As Control)
        l.ForeColor = CType(IIf(t.Text.Trim = "", Color.Red, Color.Black), Color)
    End Sub

    Public Sub Pflichtfeld(ByRef l As Control, ByVal ausgefuellt As Boolean)
        l.ForeColor = CType(IIf(ausgefuellt, Color.Black, Color.Red), Color)
    End Sub

    Public Sub Pflichtfeld(ByRef l As Control, ByVal liste As ListView)
        l.ForeColor = CType(IIf(liste.Items.Count > 0, Color.Black, Color.Red), Color)
    End Sub

    Public Function GetMandatoryColor(ByVal empty As Boolean) As Color
        If empty Then Return Color.Red Else Return Drawing.SystemColors.ControlText
    End Function

    Public Function CheckField(label As Control, field As Control, errorMsg As String) As Boolean
        If label.ForeColor.Equals(GetMandatoryColor(True)) Then
            clsShow.ErrorMsg(errorMsg)
            field.Focus()
            Return False
        End If
        Return True
    End Function

    Public Function CheckSelItem(lvw As ListView, Optional errorInfo As String = "", Optional ShowInfo As Boolean = True) As Boolean
        If lvw.SelectedItems.Count <> 0 Then Return True
        If ShowInfo Then clsShow.ErrorMsg(If(errorInfo = "", My.Resources.resGlobal.MsgPleaseSelectEntry, errorInfo))
        Return False
    End Function

    Public Sub OpenUrl(ByVal url As String)
        System.Diagnostics.Process.Start(url)
    End Sub

    Public Sub OpenFile(ByVal File As String)
        If Not FileExist(File) Then
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgFile_DoesNotExist.Replace("{0}", File))
        Else
            System.Diagnostics.Process.Start(File)
        End If
    End Sub

    Public Function RCIf(ByVal bed As Boolean, ByVal ja As Object, ByVal nein As Object) As Object
        If bed Then
            Return ja
        Else
            Return nein
        End If
    End Function

    Public Sub EnableControl(ByVal c As Control, ByVal ja As Boolean)
        If c.GetType.Name = "TextBox" Then
            With CType(c, System.Windows.Forms.TextBox)
                '.Tag = IIf(Not ja, menge.MengePlus(.Tag.ToString, "DISABLE"), menge.MengeMinus(.Tag.ToString, "DISABLE"))
                .Enabled = ja

                .TabStop = (.Enabled And Not .ReadOnly)
                '.BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "ComboBox" Then
            With CType(c, ComboBox)
                .Tag = IIf(Not ja, menge.MengePlus(Obj2Str(.Tag), "DISABLE"), menge.MengeMinus(Obj2Str(.Tag), "DISABLE"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "CheckedListBox" Then
            With CType(c, CheckedListBox)
                .Tag = IIf(Not ja, menge.MengePlus(Obj2Str(.Tag), "DISABLE"), menge.MengeMinus(Obj2Str(.Tag), "DISABLE"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "NumericUpDown" Then
            With CType(c, System.Windows.Forms.NumericUpDown)
                .Tag = IIf(Not ja, menge.MengePlus(Obj2Str(.Tag), "DISABLE"), menge.MengeMinus(Obj2Str(.Tag), "DISABLE"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "Button" Then
            With CType(c, System.Windows.Forms.Button)
                .Tag = IIf(Not ja, menge.MengePlus(Obj2Str(.Tag), "DISABLE"), menge.MengeMinus(Obj2Str(.Tag), "DISABLE"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
            End With

        ElseIf c.GetType.Name = "CheckBox" Then
            With CType(c, System.Windows.Forms.CheckBox)
                .Tag = IIf(Not ja, menge.MengePlus(Obj2Str(.Tag), "DISABLE"), menge.MengeMinus(Obj2Str(.Tag), "DISABLE"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
            End With

        ElseIf c.GetType.Name = "ListView" Then
            With CType(c, System.Windows.Forms.ListView)
                .Tag = IIf(Not ja, menge.MengePlus(Obj2Str(.Tag), "DISABLE"), menge.MengeMinus(Obj2Str(.Tag), "DISABLE"))
                .Enabled = ja

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "RadioButton" Then
            With CType(c, System.Windows.Forms.RadioButton)
                .Tag = IIf(Not ja, menge.MengePlus(Obj2Str(.Tag), "DISABLE"), menge.MengeMinus(Obj2Str(.Tag), "DISABLE"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
            End With

        ElseIf c.GetType.Name = "GroupBox" Then
            With CType(c, System.Windows.Forms.GroupBox)
                .Tag = IIf(Not ja, menge.MengePlus(Obj2Str(.Tag), "DISABLE"), menge.MengeMinus(Obj2Str(.Tag), "DISABLE"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
            End With

        ElseIf c.GetType.Name = "DateTimePicker" Then
            With CType(c, System.Windows.Forms.DateTimePicker)
                .Tag = IIf(Not ja, menge.MengePlus(Obj2Str(.Tag), "DISABLE"), menge.MengeMinus(Obj2Str(.Tag), "DISABLE"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With
            'ElseIf c.GetType.Name = "ctlDate" Then
            '    With CType(c, ctlDate)
            '        .ctlTag = IIf(ja, menge.MengePlus(Obj2Str(.ctlTag), "DISABLE"), menge.MengeMinus(Obj2Str(.ctlTag), "DISABLE"))
            '        .ctlEnabled = ja

            '        .TabStop = CBool(.ctlTag = "")
            '        .BackColor = CType(IIf(CBool(.ctlTag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            '    End With
            'ElseIf c.GetType.Name = "ctlPartnerBox" Then
            '    With CType(c, ctlPartnerBox)
            '        .ctlTag = IIf(ja, menge.MengePlus(Obj2Str(.ctlTag), "DISABLE"), menge.MengeMinus(Obj2Str(.ctlTag), "DISABLE"))
            '        .ctlEnabled = ja

            '        .TabStop = CBool(.ctlTag = "")
            '        .BackColor = CType(IIf(CBool(.ctlTag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            '    End With
        End If
    End Sub

    Public Sub LockControl(ByVal c As Control, ByVal yes As Boolean)
        If c.GetType.Name = "TextBox" Then
            With CType(c, System.Windows.Forms.TextBox)
                '.Tag = IIf(ja, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))
                .ReadOnly = yes

                .TabStop = (.Enabled And Not .ReadOnly)
                '.BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "ComboBox" Then
            With CType(c, ComboBox)
                .Tag = IIf(yes, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "CheckedListBox" Then
            With CType(c, CheckedListBox)
                .Tag = IIf(yes, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "NumericUpDown" Then
            With CType(c, System.Windows.Forms.NumericUpDown)
                .Tag = IIf(yes, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "Button" Then
            With CType(c, System.Windows.Forms.Button)
                .Tag = IIf(yes, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
            End With

        ElseIf c.GetType.Name = "CheckBox" Then
            With CType(c, System.Windows.Forms.CheckBox)
                .Tag = IIf(yes, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
            End With

        ElseIf c.GetType.Name = "ListView" Then
            With CType(c, System.Windows.Forms.ListView)
                .Tag = IIf(yes, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With

        ElseIf c.GetType.Name = "RadioButton" Then
            With CType(c, System.Windows.Forms.RadioButton)
                .Tag = IIf(yes, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
            End With

        ElseIf c.GetType.Name = "GroupBox" Then
            With CType(c, System.Windows.Forms.GroupBox)
                .Tag = IIf(yes, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
            End With

        ElseIf c.GetType.Name = "DateTimePicker" Then
            With CType(c, System.Windows.Forms.DateTimePicker)
                .Tag = IIf(yes, menge.MengePlus(Obj2Str(.Tag), "LOCK"), menge.MengeMinus(Obj2Str(.Tag), "LOCK"))
                .Enabled = CBool(.Tag = "")

                .TabStop = CBool(.Tag = "")
                .BackColor = CType(IIf(CBool(.Tag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            End With
            'ElseIf c.GetType.Name = "ctlDate" Then
            '    With CType(c, ctlDate)
            '        .ctlTag = IIf(ja, menge.MengePlus(Obj2Str(.ctlTag), "LOCK"), menge.MengeMinus(Obj2Str(.ctlTag), "LOCK"))
            '        .ctlReadOnly = ja

            '        .TabStop = CBool(.ctlTag = "")
            '        .BackColor = CType(IIf(CBool(.ctlTag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            '    End With
            'ElseIf c.GetType.Name = "ctlPartnerBox" Then
            '    With CType(c, ctlPartnerBox)
            '        .ctlTag = IIf(ja, menge.MengePlus(Obj2Str(.ctlTag), "LOCK"), menge.MengeMinus(Obj2Str(.ctlTag), "LOCK"))
            '        .ctlLocked = ja

            '        .TabStop = CBool(.ctlTag = "")
            '        .BackColor = CType(IIf(CBool(.ctlTag = ""), SystemColors.Window, SystemColors.Control), System.Drawing.Color)
            '    End With
        End If
    End Sub

    Public Function ControlIsLocked(ByVal c As Control) As Boolean
        If c.GetType.Name = "TextBox" Then
            Return CType(c, TextBox).ReadOnly

        ElseIf c.GetType.Name = "ComboBox" Then
            Return Obj2Str(CType(c, ComboBox).Tag).Contains("LOCK")

        ElseIf c.GetType.Name = "CheckedListBox" Then
            Return Obj2Str(CType(c, CheckedListBox).Tag).Contains("LOCK")

        ElseIf c.GetType.Name = "ListView" Then
            Return Obj2Str(CType(c, ListView).Tag).Contains("LOCK")

            'ElseIf c.GetType.Name = "ctlDate" Then
            '    Return Obj2Str(CType(c, ctlDate).Tag).Contains("LOCK")

        ElseIf c.GetType.Name = "Button" Then
            Return Obj2Str(CType(c, Button).Tag).Contains("LOCK")

            'ElseIf c.GetType.Name = "ctlDate" Then
            '    Return CType(c, ctlDate).ctlReadOnly

            'ElseIf c.GetType.Name = "ctlPartnerBox" Then
            '    Return CType(c, ctlPartnerBox).ctlLocked

        End If
        Return False
    End Function

    Public Function GetPcLogin() As String
        Dim mUser As System.Security.Principal.WindowsIdentity
        mUser = System.Security.Principal.WindowsIdentity.GetCurrent()
        Dim Login As String = mUser.Name
        Login = Login.Substring(Login.LastIndexOf("\") + 1)
        Return Login
    End Function

    Public Function ModifyFileName(ByVal s As String) As String
        ' aus ".\lpass.mdb" wird "c:\programme\lpass\lpass.mdb"
        If Left$(s, 2) = ".\" Then
            Return session.appDir & Mid$(s, 2)
        End If
        Return s
    End Function

    Public Function ModifyDirName(ByVal s As String) As String
        ' aus ".\vorlagen" wird "c:\programme\lpass\vorlagen"
        If s = "." Then
            Return session.appDir
        ElseIf Left$(s, 2) = ".\" Then
            Return IO.Path.Combine(session.appDir, Mid$(s, 3))
        End If
        Return s
    End Function

    Public Function GetWSPTempPath() As String
        Dim s As String = IO.Path.Combine(IO.Path.GetTempPath, "WSP")
        If Not clsDirectory.Exists(s) Then clsDirectory.Make(s)
        Return s
    End Function

    Public Function NullenRaus(ByVal txt As String) As String
        Dim i As Long
        Dim ret As String
        ret = txt
        For i = 1 To txt.Length
            If ret.Substring(0, 1) = "0" Then
                ret = Right(ret, ret.Length - 1)
            Else
                Exit For
            End If
        Next
        Return ret
    End Function

    Public Function umlaute_raus(ByVal s As String) As String
        Dim ret As String : ret = s
        If s.IndexOf("ä") >= 0 Then ret = ret.Replace("ä", "ae")
        If s.IndexOf("Ä") >= 0 Then ret = ret.Replace("Ä", "Ae")
        If s.IndexOf("ö") >= 0 Then ret = ret.Replace("ö", "oe")
        If s.IndexOf("Ö") >= 0 Then ret = ret.Replace("Ö", "Oe")
        If s.IndexOf("ü") >= 0 Then ret = ret.Replace("ü", "ue")
        If s.IndexOf("Ü") >= 0 Then ret = ret.Replace("Ü", "Ue")
        If s.IndexOf("ß") >= 0 Then ret = ret.Replace("ß", "ss")
        umlaute_raus = ret
    End Function

    Public Sub Wait(ByVal wl As Double)
        Dim i As Date = Now.AddSeconds(wl)
        While i > Now
            Application.DoEvents()
        End While
    End Sub

    Public Sub Sleep(ByVal milliseconds As Integer)
        Threading.Thread.Sleep(milliseconds)
    End Sub

    Public Function FileIsOpen(ByVal filePfad As String) As Boolean
        Try
            File.OpenWrite(filePfad).Close()
        Catch ex As Exception
            Return True
        End Try
        Return False

    End Function

    Function FileExist(ByVal f As String) As Boolean
        Try
            Return Not (Dir(f) = "")
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function FileCopy(ByVal sour As String, ByVal dest As String) As Boolean
        Try
            File.Copy(sour, dest)
        Catch ex As Exception
            clsShow.ErrorMsg(PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Function FileCopy(ByVal sour As String, ByVal dest As String, ByVal pOverwrite As Boolean) As Boolean
        Try
            File.Copy(sour, dest, pOverwrite)
        Catch ex As Exception
            clsShow.ErrorMsg(PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Function FileDelete(ByVal filePfad As String) As Boolean
        Try
            File.Delete(filePfad)
        Catch ex As Exception
            clsShow.ErrorMsg(PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Sub KommaCat(ByRef s As String, ByVal wort As String)
        If s.Length <> 0 Then s += ", "
        s += wort
    End Sub

    Public Function StringContains(ByVal s1 As String, ByVal s2 As String) As Integer
        Dim ret As Integer = 0
        Dim lastPos As Integer = 1

        Do
            lastPos = InStr(lastPos, s1, s2)
            If lastPos <> 0 Then
                ret += 1
                lastPos += 1
            End If

        Loop While lastPos <> 0

        Return ret
    End Function

    '2008-06-20 GB
    'Funktion zur Entfernung und Ersetzung von Leerzeichen (auch mehrere aufeinanderfolgende) in Zeichenfolgen
    Public Function strTrimReplaceBlanks(ByVal pStr As String, Optional ByVal pTrenner As String = ",") As String
        If pStr.Trim = "" Then Return ""

        Dim s() As String = Split(pStr, " ")
        Dim tmp As String
        Dim ret As String = ""

        For i As Integer = 0 To UBound(s)
            tmp = s(i).Trim
            If tmp.Length > 0 Then tex.Cat(ret, tmp, pTrenner)
        Next

        Return ret
    End Function

    Public Function TrennzeichenStringToIntList(ByVal Zeichen As Char, ByVal value As String) As Generic.List(Of Integer)
        Dim liste As New Generic.List(Of Integer)

        Dim parts() As String = value.Split(CChar(Zeichen))
        For i As Integer = 0 To parts.Length - 1
            If Trim(parts(i)) = "" Then Continue For

            liste.Add(CInt(Trim(parts(i))))
        Next
        Return liste
    End Function

    Function strpartLeft(ByVal s As String, ByVal p As Integer, ByVal t As String) As String
        'strpartLeft("A,B,C,D", 3, ",") -> "A,B,C"
        If s = "" Then Return ""

        s = Replace(s, t, ControlChars.NullChar)
        Dim parts() As String = s.Split(ControlChars.NullChar)

        If p < 1 Then Return ""
        If p > parts.Length Then p = parts.Length

        s = String.Join(t, parts, 0, p)

        Return s
    End Function

    Public Function Obj2Str(ByVal o As Object) As String
        Dim ret As String = ""
        If o Is Nothing Then Return ret
        Dim s As String
        Dim b As Boolean
        Dim d As DateTime
        Dim f As Double
        Dim x As Short
        Dim i As Integer
        Dim l As Long

        If o.GetType() Is Type.GetType("System.String") Then
            s = CType(o, String)
            Return s
        ElseIf o.GetType() Is Type.GetType("System.Boolean") Then
            b = CType(o, Boolean)
            Return CStr(IIf(b, "True", "False"))
        ElseIf o.GetType() Is Type.GetType("System.DateTime") Then
            d = CType(o, DateTime)
            If dat.IsNull(d) Then Return ""
            Return dat.printDIN(d)
        ElseIf o.GetType() Is Type.GetType("System.Double") Then
            f = CType(o, Double)
            Return zahl.printGerDecimal(f)
        ElseIf o.GetType() Is Type.GetType("System.Int16") Then
            x = CType(o, Short)
            Return x.ToString()
        ElseIf o.GetType() Is Type.GetType("System.Integer") Then
            i = CType(o, Integer)
            Return i.ToString()
        ElseIf o.GetType() Is Type.GetType("System.Int32") Then
            i = CType(o, Integer)
            Return i.ToString()
        ElseIf o.GetType() Is Type.GetType("System.Int64") Then
            'ElseIf o.GetType() Is Type.GetType("System.Long") Then
            l = CType(o, Long)
            Return l.ToString()
        ElseIf o.GetType Is Type.GetType("System.Byte[]") Then
            o = clsComDecompress.Compress(CType(o, Byte()))
            Return Convert.ToBase64String(CType(o, Byte()))
        Else
            Try
                ret = CType(o, String)
            Catch ex As Exception
                'nix
            End Try
        End If
        Return ret
    End Function

    Public Function PrintException(ByVal ex As Exception) As String
        Dim msg As String = ex.Message & vbCrLf & ex.StackTrace
        If ex.InnerException IsNot Nothing Then
            msg = msg & vbCrLf & vbCrLf & My.Resources.resGlobal.TextCausedByFollowingInnerException & vbCrLf & vbCrLf & PrintException(ex.InnerException)
        End If
        Return msg
    End Function

    Public Class clsComDecompress

        Public Shared Function Compress(ByVal input As Byte()) As Byte()
            Dim output As New MemoryStream
            Dim Zip As New GZipStream(output, CompressionMode.Compress)
            Zip.Write(input, 0, input.Length)
            Zip.Close()
            Return output.ToArray
        End Function

        Public Shared Function Decompress(ByVal input As Byte()) As Byte()
            Dim decomStream As New MemoryStream(input)
            Dim ZIP As New GZipStream(decomStream, CompressionMode.Decompress, True)

            Dim stepp As Byte()
            ReDim stepp(1023) 'Instead of 16 can put any 2^x     
            Dim outStream As New MemoryStream()
            Dim readCount As Integer

            Do
                readCount = ZIP.Read(stepp, 0, stepp.Length)
                outStream.Write(stepp, 0, readCount)
            Loop While (readCount > 0)

            ZIP.Close()
            Return outStream.ToArray()
        End Function

    End Class

    Public Sub ControlSetDoubleBuffered(ByVal ctl As Control, ByVal doubleBufferedOn As Boolean)
        Dim typ As Type = ctl.GetType()
        Dim pi As System.Reflection.PropertyInfo = typ.GetProperty("DoubleBuffered", Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
        pi.SetValue(ctl, doubleBufferedOn, Nothing)
    End Sub

End Module
