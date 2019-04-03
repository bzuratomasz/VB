Module modMain
    Public session As New clsSession

    Public Function Main() As Integer
        Application.EnableVisualStyles()
        Application.DoEvents()

        session = New clsSession

        Try
            If Not session.Init() Then Return -1
            If Not session.ReadIniFile() Then Return -1
        Catch ex As Exception
            If ex.ToString.IndexOf("SecurityException") = -1 Then
                clsShow.ErrorMsg(PrintException(ex))
            Else
                clsShow.ErrorMsg(My.Resources.resGlobal.MsgProgramCannotBeStartedBecauseOfMissingPermissions4RunningDotNetProgramsFromSecurityZoneLocalIntranet)
            End If
            Return -1
        End Try

        If Not session.OpenDB() Then Return -1

        ' Updating database
        If Not UpdateDatabase() Then
            session.CloseDB()
            Return -1
        End If

        ' Checking executable version
        If Not session.CheckVersion() Then
            session.CloseDB()
            Windows.Forms.Cursor.Current = Cursors.Default : Application.DoEvents()
            Return -1
        End If

        Dim frm As New frmMain
        Application.DoEvents()
        Application.Run(frm)
        frm.Dispose()

        session.CloseDB()

        Return 0
    End Function

    Public Function GetDirectories(ByVal Pfad As String) As List(Of String)
        If Not clsDirectory.Exists(Pfad) Then Return New List(Of String)
        Return System.IO.Directory.GetDirectories(Pfad).ToList()
    End Function

    Public Function GetFiles(ByVal Pfad As String) As List(Of String)
        If Not clsDirectory.Exists(Pfad) Then Return New List(Of String)
        Return System.IO.Directory.GetFiles(Pfad).ToList()
    End Function
End Module
