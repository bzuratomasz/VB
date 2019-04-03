Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient

'==================================================================================================
Public Class clsXWords
    Private XWords As New Generic.List(Of String)

    Public Sub New(ByVal str As String)
        If XWords Is Nothing Then XWords = New Generic.List(Of String)
        Dim strLst As List(Of String) = str.ToListOfString(";"c)
        For Each s As String In strLst
            If s.ToUpper.StartsWith("USER") Then Add(tex.Part(s, 2, "="))
            If s.ToUpper.StartsWith("PASS") Then Add(tex.Part(s, 2, "="))
        Next
    End Sub

    Public Sub Reset()
        XWords = New Generic.List(Of String)
    End Sub

    Public Sub Add(ByVal wort As String)
        XWords.Add(wort)
    End Sub

    Public Function Replace(ByVal s As String) As String
        Dim i As Integer
        For i = 0 To XWords.Count - 1
            Dim h As String = XWords.Item(i)
            s = Microsoft.VisualBasic.Replace(s, h, "***") 'New String("X"c, Len(h)))
        Next i

        Return s
    End Function
End Class

'==================================================================================================
Public Class clsConnection
    Public Connection As System.Data.Common.DbConnection 'OleDB.OleDbConnection / SqlClient.SqlConnection
    Public InUseOfDataReader As Boolean
End Class

'==================================================================================================
Public Class clsConnectionPool
    Private db As clsDatabase
    Private alConnections As Generic.List(Of clsConnection)

    Public Sub New(ByVal db As clsDatabase)
        Me.db = db
        alConnections = New Generic.List(Of clsConnection)
    End Sub

    Public Function OpenCon(Optional ByVal forDataReader As Boolean = True) As clsConnection
        Dim retConnection As clsConnection

        Dim StackTrace As New StackTrace(True)
        Dim StackFrame As New StackFrame(True)

        For Each Frame As StackFrame In StackTrace.GetFrames
            If Frame.GetFileName Is Nothing Then Continue For
            If Frame.GetFileName.ToUpper.Contains("clsDatabase".ToUpper) Then Continue For
            If Frame.GetFileName.ToUpper.Contains("modMain".ToUpper) Then Continue For
            StackFrame = Frame
            Exit For
        Next

        'Falls noch keine Verbindung besteht oder die Standard-Verbindung zur Zeit von einem DataReader genutzt wird, 
        'zuerst (neue) Standard-Verbindung auf Index 0 anlegen
        If alConnections.Count = 0 OrElse (TypeOf db Is clsDatabaseSQLServer And alConnections(0).InUseOfDataReader) Then
            retConnection = New clsConnection

            If TypeOf db Is clsDatabaseAccess Then
                retConnection.Connection = New OleDbConnection(db.connstr)
            ElseIf TypeOf db Is clsDatabaseSQLServer Then
                retConnection.Connection = New SqlConnection(db.connstr)
            ElseIf TypeOf db Is clsDatabaseOracle Then
                retConnection.Connection = New OleDbConnection(db.connstr)
            End If

            alConnections.Insert(0, retConnection)
            If alConnections.Count > CInt(IIf(session.IniTransmar, 4, 10)) Then clsShow.InternalError(My.Resources.resMain.MsgNumberOfConnectionsIncreased.Replace("{0}", alConnections.Count.ToString)) 'vermutlich fehlt irgendwo dr.Close()

            'Debug.Print(Now.ToString & " | Connection opened by " & StackFrame.GetMethod.DeclaringType.Name & "." & StackFrame.GetMethod.Name & " in " & IO.Path.GetFileName(StackFrame.GetFileName) & " at line " & StackFrame.GetFileLineNumber.ToString & " | {0} of {1} open connection(s)".Replace("{0}", Count(ConnectionState.Open)).Replace("{1}", Count()))
        End If

        If forDataReader Then alConnections(0).InUseOfDataReader = True
        retConnection = alConnections(0)

        If retConnection.Connection.State <> ConnectionState.Open Then
            If Not OpenCon_withLoop(retConnection.Connection) Then CloseCon(retConnection)
        End If

        Debug.Print(Now.ToString & " | Connection occupied by " & StackFrame.GetMethod.DeclaringType.Name & "." & StackFrame.GetMethod.Name & " in " & IO.Path.GetFileName(StackFrame.GetFileName) & " at line " & StackFrame.GetFileLineNumber.ToString & " | {0} of {1} open connection(s)".Replace("{0}", Count(ConnectionState.Open).ToString()).Replace("{1}", Count().ToString()))

        Return retConnection
    End Function

    Public Function OpenCon_withLoop(ByRef cn As Data.Common.DbConnection) As Boolean
        Dim wiederholen As Boolean = False
        Dim anzWiederhol As Integer = 5
        Dim anzWiederholSum As Integer = anzWiederhol

        Dim err_no As Integer = 0
        Dim err_txt As String = ""

        Dim StackTrace As New StackTrace(True)
        Dim StackFrame As New StackFrame(True)

        For Each Frame As StackFrame In StackTrace.GetFrames
            If Frame.GetFileName Is Nothing Then Continue For
            If Frame.GetFileName.ToUpper.Contains("clsDatabase".ToUpper) Then Continue For
            If Frame.GetFileName.ToUpper.Contains("modMain".ToUpper) Then Continue For
            StackFrame = Frame
            Exit For
        Next

        Do
            'If cn.State = ConnectionState.Open Then
            '    cn.Close()
            '    Debug.Print(Now.ToString & " | Connection closed by " & StackFrame.GetMethod.DeclaringType.Name & "." & StackFrame.GetMethod.Name & " in " & IO.Path.GetFileName(StackFrame.GetFileName) & " at line " & StackFrame.GetFileLineNumber.ToString & " | {0} of {1} open connection(s)".Replace("{0}", Count(ConnectionState.Open).ToString()).Replace("{1}", Count().ToString()))
            'End If
            If cn.State <> ConnectionState.Closed Then
                cn.Close()
                Debug.Print(Now.ToString & " | Connection closed by " & StackFrame.GetMethod.DeclaringType.Name & "." & StackFrame.GetMethod.Name & " in " & IO.Path.GetFileName(StackFrame.GetFileName) & " at line " & StackFrame.GetFileLineNumber.ToString & " | {0} of {1} open connection(s)".Replace("{0}", Count(ConnectionState.Open).ToString()).Replace("{1}", Count().ToString()))
            End If

            Try
                Debug.Print(Now.ToString & " | Attempt " & anzWiederholSum - anzWiederhol + 1 & " to open connection by " & StackFrame.GetMethod.DeclaringType.Name & "." & StackFrame.GetMethod.Name & " in " & IO.Path.GetFileName(StackFrame.GetFileName) & " at line " & StackFrame.GetFileLineNumber.ToString & " | {0} of {1} open connection(s)".Replace("{0}", Count(ConnectionState.Open).ToString()).Replace("{1}", Count().ToString()))

                cn.ConnectionString = db.connstr 'für Wiederholungsfall hier neu setzen, da Passwort bei cn.Open verloren geht
                cn.Open()
                err_no = 0
                err_txt = ""

                Debug.Print(Now.ToString & " | Connection opened by " & StackFrame.GetMethod.DeclaringType.Name & "." & StackFrame.GetMethod.Name & " in " & IO.Path.GetFileName(StackFrame.GetFileName) & " at line " & StackFrame.GetFileLineNumber.ToString & " | {0} of {1} open connection(s)".Replace("{0}", Count(ConnectionState.Open).ToString()).Replace("{1}", Count().ToString()))
            Catch SqlExp As SqlClient.SqlException
                '-2146232060 = Netzwerkbezogener oder instanzspezifischer Fehler beim Herstellen einer Verbindung mit SQL Server. Der Server wurde nicht gefunden, oder auf ihn kann nicht zugegriffen werden. Überprüfen Sie, ob der Instanzname richtig ist und ob SQL Server Remoteverbindungen zulässt. (provider: Shared Memory-Provider, error: 40 - Verbindung mit SQL Server konnte nicht geöffnet werden)
                err_no = SqlExp.ErrorCode
                err_txt = SqlExp.Message
            Catch OleExp As OleDb.OleDbException
                '-2147467259 = SQL Server existiert nicht oder Zugriff verweigert.
                '-2147467259 = Unrecognized database format '...\MTool.mdb'
                '-2147467259 = Could not find file '...\MTool.mdb'
                '-2147467259 = Could not use ''; file already in use.
                '-2147217843 = Cannot start your application. The workgroup information file is missing or opened exclusively by another user.
                err_no = OleExp.ErrorCode
                err_txt = OleExp.Message
            Catch Exp As Exception
                'The 'Microsoft.Jet.OLEDB.4.0' provider is not registered on the local machine.
                err_no = 99999
                err_txt = Exp.Message
            End Try

            wiederholen = False
            If err_no <> 0 Then
                Select Case err_no
                    Case -2146232060, -2147467259
                        If anzWiederhol > 0 Then
                            wiederholen = True
                            anzWiederhol -= 1
                            Sleep(5)
                        Else
                            Dim f As String
                            f = My.Resources.resMain.MsgReconnectingToDatabaseFailed
                            f += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                            f += vbCrLf & My.Resources.resMain.MsgPleaseWaitSomeSecondsMinutesToRepeatExecution
                            'If MsgBox(f, MsgBoxStyle.RetryCancel, GetSprachBez("Fehler", "Error")) = MsgBoxResult.Retry Then wiederholen = True
                            Throw New Exception(err_txt)
                            'If ShowRetry(f) Then wiederholen = True
                        End If
                    Case Else
                        Dim m As String
                        m = My.Resources.resMain.MsgReconnectingToDatabaseFailed
                        m += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                        clsShow.ErrorMsg(m)
                End Select

                If Not wiederholen Then
                    Dim h As String
                    h = My.Resources.resMain.MsgErrorWhileReopeningDatabase
                    h += vbCrLf & My.Resources.resMain.TextConnectionString.Replace("{0}", db.XWords.Replace(db.connstr))
                    h += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                    'h += vbCrLf & "Time to failure: " & Format$(p.TimeElapsed, "0.0") & " ms"

                    db.LastError = h
                    'TODO: clsLog.LogLine(h)
                End If
            End If
        Loop While wiederholen

        Return (err_no = 0)
    End Function

    Public Sub CloseCon(ByVal closeCon As clsConnection)
        If closeCon Is Nothing Then Return

        Dim StackTrace As New StackTrace(True)
        Dim StackFrame As New StackFrame(True)

        For Each Frame As StackFrame In StackTrace.GetFrames
            If Frame.GetFileName Is Nothing Then Continue For
            If Frame.GetFileName.ToUpper.Contains("clsDatabase".ToUpper) Then Continue For
            If Frame.GetFileName.ToUpper.Contains("modMain".ToUpper) Then Continue For
            StackFrame = Frame
            Exit For
        Next

        'If Not closeCon.InUseOfDataReader Then Return 'Standard-Verbindung soll geöffnet bleiben

        Dim con As clsConnection
        Dim i As Integer
        For i = alConnections.Count - 1 To 0 Step -1
            con = alConnections(i)
            If con Is closeCon Then
                If i = 0 Then 'falls Standard-Verbindung zwischenzeitlich für DataReader genutzt wurde, nur Flag umsetzen
                    alConnections(i).InUseOfDataReader = False

                    Debug.Print(Now.ToString & " | Connection released by " & StackFrame.GetMethod.DeclaringType.Name & "." & StackFrame.GetMethod.Name & " in " & IO.Path.GetFileName(StackFrame.GetFileName) & " at line " & StackFrame.GetFileLineNumber.ToString & " | {0} of {1} open connection(s)".Replace("{0}", Count(ConnectionState.Open).ToString()).Replace("{1}", Count().ToString()))
                Else          'sonst schließen und aus Pool entfernen
                    con.Connection.Close()
                    alConnections.RemoveAt(i)

                    Debug.Print(Now.ToString & " | Connection closed by " & StackFrame.GetMethod.DeclaringType.Name & "." & StackFrame.GetMethod.Name & " in " & IO.Path.GetFileName(StackFrame.GetFileName) & " at line " & StackFrame.GetFileLineNumber.ToString & " | {0} of {1} open connection(s)".Replace("{0}", Count(ConnectionState.Open).ToString()).Replace("{1}", Count().ToString()))
                End If

                Exit For
            End If
        Next
    End Sub

    Public Sub CloseAll()
        Dim con As clsConnection

        Dim StackTrace As New StackTrace(True)
        Dim StackFrame As New StackFrame(True)

        For Each Frame As StackFrame In StackTrace.GetFrames
            If Frame.GetFileName Is Nothing Then Continue For
            If Frame.GetFileName.ToUpper.Contains("clsDatabase".ToUpper) Then Continue For
            If Frame.GetFileName.ToUpper.Contains("modMain".ToUpper) Then Continue For
            StackFrame = Frame
            Exit For
        Next

        Dim i As Integer
        For i = alConnections.Count - 1 To 0 Step -1
            con = alConnections(i)
            con.Connection.Close()
            alConnections.RemoveAt(i)

            Debug.Print(Now.ToString & " | Connection closed by " & StackFrame.GetMethod.DeclaringType.Name & "." & StackFrame.GetMethod.Name & " in " & IO.Path.GetFileName(StackFrame.GetFileName) & " at line " & StackFrame.GetFileLineNumber.ToString & " | {0} of {1} open connection(s)".Replace("{0}", Count(ConnectionState.Open).ToString()).Replace("{1}", Count().ToString()))
        Next
    End Sub

    Public ReadOnly Property Count(Optional ByVal State As ConnectionState = Nothing) As Integer
        Get
            Dim ret As Integer = 0

            For Each Connection As clsConnection In alConnections
                If State <> Nothing Then If Not Connection.Connection.State = State Then Continue For
                ret += 1
            Next

            Return ret
        End Get
    End Property

End Class

'==================================================================================================
Public Class clsDatabaseAccess
    Inherits clsDatabase

    Private mIdentitySupported As Boolean

    Public Sub New(ByVal connstr As String, ByVal Identity_Supported As Boolean)
        MyBase.New(connstr)

        Me.mIdentitySupported = Identity_Supported
    End Sub

    Public Overrides ReadOnly Property IdentitySupported() As Boolean
        Get
            Return mIdentitySupported
        End Get
    End Property

    Public Overrides Function sqlBln(ByVal b As Boolean) As String
        Return IIf(b, "-1", "0").ToString()
    End Function

    Public Overrides Function sqlDate(ByVal dat As Date) As String
        Return "#" & dat.ToString("yyyy\-MM\-dd") & "#"
    End Function

    Public Overrides Function sqlDateTime(ByVal dat As Date) As String
        Return "#" & dat.ToString("yyyy\-MM\-dd HH\:mm\:ss") & "#"
    End Function

    Public Overrides Function sqlTime(ByVal dat As Date) As String
        Return "#" & dat.ToString("HH\:mm\:ss") & "#"
    End Function

    Public Overrides Function sqlGetDateTime() As Date
        Return sqlGetDat("SELECT NOW")
    End Function

    Public Overrides Function sqlGetUTCDateTime() As Date
        Return sqlGetDat("SELECT NOW").ToUniversalTime
    End Function

    Public Overrides Function sqlNNLng(ByVal fieldname As String) As String
        Return "IIF(" & fieldname & " IS NULL,0," & fieldname & ")"
    End Function

    Public Overrides Function sqlNNStr(ByVal fieldname As String) As String
        Return "IIF(" & fieldname & " IS NULL,''," & fieldname & ")"
    End Function

    Public Overrides Function sqlISNULL(ByVal fieldname1 As String, ByVal fieldname2 As String, Optional ByVal fieldname3 As String = "") As String
        If fieldname3 = "" Then
            Return "IIF(" & fieldname1 & " IS NULL," & fieldname2 & "," & fieldname1 & ")"
        Else
            Return "IIF(" & fieldname1 & " IS NULL,IIF(" & fieldname2 & " IS NULL," & fieldname3 & "," & fieldname2 & ")," & fieldname1 & ")"
        End If
    End Function

    Public Overrides Function sqlBOOL(ByVal bed As String) As String
        Return "IIF(" & bed & ",1,0)"
    End Function

    Public Overrides Function sqlTRIM(ByVal fieldname As String) As String
        Return "TRIM(" & fieldname & ")"
    End Function

    Public Overrides Function sqlTOP(ByVal anz As Integer) As String
        Return "TOP " & anz
    End Function

    Public Overrides Function sqlMemo2Text(ByVal memo As String) As String
        Return memo
    End Function

    Public Overrides Function sqlText2Float(ByVal fieldname As String) As String
        Return "VAL(" & fieldname & ")"
    End Function
End Class

'==================================================================================================
Public Class clsDatabaseSQLServer
    Inherits clsDatabase

    Public Const ErrorDuplicateKey = 2601

    Public Sub New(ByVal connstr As String)
        MyBase.New(connstr)
    End Sub

    Public Overrides ReadOnly Property IdentitySupported() As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property WithRowLock() As String
        Get
            Return " WITH (ROWLOCK)"
        End Get
    End Property

    Public Overrides ReadOnly Property WithNoLock() As String
        Get
            Return " WITH (NOLOCK)"
        End Get
    End Property

    Public Overrides Function sqlStr(ByVal s As String) As String
        Return "N'" & s.Replace("'", "''") & "'"
    End Function

    Public Overrides Function sqlBln(ByVal b As Boolean) As String
        Return IIf(b, "1", "0").ToString()
    End Function

    Public Overrides Function sqlDate(ByVal dat As Date) As String
        Return "{d '" & dat.ToString("yyyy\-MM\-dd") & "'}"
    End Function

    Public Overrides Function sqlDateTime(ByVal dat As Date) As String
        Return "{ts '" & dat.ToString("yyyy\-MM\-dd HH\:mm\:ss") & "'}"
    End Function

    Public Overrides Function sqlTime(ByVal dat As Date) As String
        'Return "{t '" & dat.ToString("HH\:mm\:ss") & "'}"
        Return "{ts '1900-01-01 " & dat.ToString("HH\:mm\:ss") & "'}"
    End Function

    Public Overrides Function sqlGetDateTime() As Date
        Return sqlGetDat("SELECT GETDATE()")
    End Function

    Public Overrides Function sqlGetUTCDateTime() As Date
        Return sqlGetDat("SELECT GETUTCDATE()")
    End Function

    Public Overrides Function isSQLServer() As Boolean
        Return True
    End Function

    Public Overrides Function CatalogName() As String
        Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
        If tmpCon.Connection.State <> ConnectionState.Open Then ConnectionPool.CloseCon(tmpCon) : Return ""

        Dim ret As String = CType(tmpCon.Connection, SqlConnection).Database

        ConnectionPool.CloseCon(tmpCon)

        Return ret
    End Function

    Public Overrides Function sqlNNLng(ByVal fieldname As String) As String
        Return "ISNULL(" & fieldname & ",0)"
    End Function

    Public Overrides Function sqlNNStr(ByVal fieldname As String) As String
        Return "ISNULL(" & fieldname & ",'')"
    End Function

    Public Overrides Function sqlISNULL(ByVal fieldname1 As String, ByVal fieldname2 As String, Optional ByVal fieldname3 As String = "") As String
        If fieldname3 = "" Then
            Return "ISNULL(" & fieldname1 & "," & fieldname2 & ")"
        Else
            Return "ISNULL(" & fieldname1 & ",ISNULL(" & fieldname2 & "," & fieldname3 & "))"
        End If
    End Function

    Public Overrides Function sqlBOOL(ByVal bed As String) As String
        Return "(CASE WHEN " & bed & " THEN 1 ELSE 0 END)"
    End Function

    Public Overrides Function sqlTRIM(ByVal fieldname As String) As String
        Return "LTRIM(RTRIM(" & fieldname & "))"
    End Function

    Public Overrides Function sqlTOP(ByVal anz As Integer) As String
        Return "TOP " & anz
    End Function

    Public Overrides Function sqlMemo2Text(ByVal memo As String) As String
        Return "CONVERT(NVARCHAR(4000)," & memo & ")"
    End Function

    Public Overrides Function sqlText2Float(ByVal fieldname As String) As String
        Return "CONVERT(FLOAT," & fieldname & ")"
    End Function

    Public ReadOnly Property SQL2005() As Boolean
        Get
            Static ver As Integer
            If ver = 0 Then
                Dim s As String = "SELECT SERVERPROPERTY('ProductVersion')"
                Dim h As String = Me.sqlGetStr(s)
                ver = zahl.getInt(tex.Part(h, 1, "."))
            End If
            Return (ver >= 9)
        End Get
    End Property
End Class

'==================================================================================================
Public Class clsDatabaseOracle
    Inherits clsDatabase

    Public Sub New(ByVal connstr As String)
        MyBase.New(connstr)
    End Sub

    Public Overrides ReadOnly Property IdentitySupported() As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides Function sqlBln(ByVal b As Boolean) As String
        Return IIf(b, "1", "0").ToString()
    End Function

    Public Overrides Function sqlStr(ByVal s As String) As String
        Dim h As String = s
        h = h.Replace("'", "''")
        h = h.Replace("`", "''")
        h = h.Replace("´", "''")
        Return "'" & h & "'"
    End Function

    Public Overrides Function sqlDate(ByVal dat As Date) As String
        Return "#" & dat.ToString("yyyy\-MM\-dd") & "#"
    End Function

    Public Overrides Function sqlDateTime(ByVal dat As Date) As String
        Return "#" & dat.ToString("yyyy\-MM\-dd HH\:mm\:ss") & "#"
    End Function

    Public Overrides Function sqlTime(ByVal dat As Date) As String
        Return "#" & dat.ToString("HH\:mm\:ss") & "#"
    End Function

    Public Overrides Function sqlGetDateTime() As Date
        Return sqlGetDat("SELECT NOW")
    End Function

    Public Overrides Function sqlGetUTCDateTime() As Date
        Return sqlGetDat("SELECT GETUTCDATE()")
    End Function

    Public Overrides Function sqlNNLng(ByVal fieldname As String) As String
        Return "IIF(" & fieldname & " IS NULL,0," & fieldname & ")"
    End Function

    Public Overrides Function sqlNNStr(ByVal fieldname As String) As String
        Return "IIF(" & fieldname & " IS NULL,''," & fieldname & ")"
    End Function

    Public Overrides Function sqlISNULL(ByVal fieldname1 As String, ByVal fieldname2 As String, Optional ByVal fieldname3 As String = "") As String
        If fieldname3 = "" Then
            Return "IIF(" & fieldname1 & " IS NULL," & fieldname2 & "," & fieldname1 & ")"
        Else
            Return "IIF(" & fieldname1 & " IS NULL,IIF(" & fieldname2 & " IS NULL," & fieldname3 & "," & fieldname2 & ")," & fieldname1 & ")"
        End If
    End Function

    Public Overrides Function sqlBOOL(ByVal bed As String) As String
        Return "IIF(" & bed & ",1,0)"
    End Function

    Public Overrides Function sqlTRIM(ByVal fieldname As String) As String
        Return "TRIM(" & fieldname & ")"
    End Function

    Public Overrides Function sqlTOP(ByVal anz As Integer) As String
        Return "TOP " & anz
    End Function

    Public Overrides Function sqlMemo2Text(ByVal memo As String) As String
        Return memo
    End Function

    Public Overrides Function sqlText2Float(ByVal fieldname As String) As String
        Return fieldname
    End Function
End Class

'==================================================================================================
Public MustInherit Class clsDatabase
    Friend connstr As String
    Friend ConnectionPool As clsConnectionPool
    Friend XWords As clsXWords
    Public LastError As String

    Public DBFileDownloadDriveMinSpaceLeft As Integer = 10485760 '10 MB
    'Public DBPacketSize4Varbinary As Integer = 16777216 'max. 16 MB je VarBinary-Datenpaket -> öfter eine MemoryOverflowException
    Public DBPacketSize4Varbinary As Integer = 8388608 'max. 8 MB je VarBinary-Datenpaket
    'Public DBPacketSize4Varbinary As Integer = 4194304 'max. 4 MB je VarBinary-Datenpaket

    Public Sub New(ByVal connstr As String)
        Me.connstr = connstr
        Me.ConnectionPool = New clsConnectionPool(Me)
        Me.XWords = New clsXWords(connstr)
        Me.LastError = ""
    End Sub

    Public Function Open() As Boolean
        'Prüfung der Verbindung bzw. Öffnen der "Hauptverbindung"
        Dim tmpCon As clsConnection = ConnectionPool.OpenCon(False)

        If tmpCon.Connection.State <> ConnectionState.Open Then ConnectionPool.CloseCon(tmpCon) : Return False

        'ConnectionPool.CloseCon(tmpCon) 'Hauptverbindung wird erst beim kompletten Schließen zurückgegeben

        Return True
    End Function

    Public Sub Close()
        ConnectionPool.CloseAll()
    End Sub

    Public Overridable ReadOnly Property WithRowLock() As String
        Get
            Return ""
        End Get
    End Property
    Public Overridable ReadOnly Property WithNoLock() As String
        Get
            Return ""
        End Get
    End Property

    Public MustOverride ReadOnly Property IdentitySupported() As Boolean

    Public Overridable Function sqlStr(ByVal s As String) As String
        Return "'" & s.Replace("'", "''") & "'"
    End Function

    Public Overridable Function sqlBln(ByVal b As Boolean) As String
        Return IIf(b, "TRUE", "FALSE").ToString() 'overridden für SQL-Server (1/0) und Access (-1/0)
    End Function

    Public MustOverride Function sqlDate(ByVal dat As Date) As String
    Public MustOverride Function sqlDateTime(ByVal dat As Date) As String
    Public MustOverride Function sqlTime(ByVal dat As Date) As String
    Public MustOverride Function sqlGetDateTime() As Date
    Public MustOverride Function sqlGetUTCDateTime() As Date

    Public Overridable Function sqlDbl(ByVal f As Double) As String
        Return f.ToString().Replace(",", ".")
    End Function

    Public Overridable Function sqlLike(ByVal s As String) As String
        Return sqlStr("%" & s & "%")

        'Access verwendet *, per OleDB jedoch % anzugeben
        'If TypeOf Me Is clsDatabaseAccess Then
        '    Return sqlStr("*" & s & "*")
        'Else
        '    Return sqlStr("%" & s & "%")
        'End If
    End Function

    Public Overridable Function isSQLServer() As Boolean
        Return False
    End Function

    Public Overridable Function CatalogName() As String
        Return "-"
    End Function

    Public MustOverride Function sqlNNLng(ByVal fieldname As String) As String
    Public MustOverride Function sqlNNStr(ByVal fieldname As String) As String

    Public MustOverride Function sqlISNULL(ByVal fieldname1 As String, ByVal fieldname2 As String, Optional ByVal fieldname3 As String = "") As String
    Public MustOverride Function sqlBOOL(ByVal bed As String) As String
    Public MustOverride Function sqlTRIM(ByVal fieldname As String) As String

    Public MustOverride Function sqlTOP(ByVal anz As Integer) As String

    Public MustOverride Function sqlMemo2Text(ByVal memo As String) As String
    Public MustOverride Function sqlText2Float(ByVal fieldname As String) As String

    Public Function sqlValue(ByVal o As Object) As String
        'Umwandlung des Objektes zur Nutzung in SQL-Statements
        If o Is Nothing OrElse IsDBNull(o) Then Return "NULL"

        Dim s As String
        Dim b As Boolean
        Dim d As DateTime
        Dim f As Double
        Dim x As Short
        Dim i As Integer
        Dim l As Long
        If o.GetType() Is Type.GetType("System.String") Then
            s = CType(o, String)
            If s = "" Then Return "NULL"
            Return sqlStr(s)
        ElseIf o.GetType() Is Type.GetType("System.Boolean") Then
            b = CType(o, Boolean)
            Return sqlBln(b)
        ElseIf o.GetType() Is Type.GetType("System.DateTime") Then
            d = CType(o, DateTime)
            If dat.IsNull(d) Then Return "NULL"
            If dat.IsTime(d) Then Return sqlTime(d)
            Return sqlDateTime(d)
        ElseIf o.GetType() Is Type.GetType("System.Double") Then
            f = CType(o, Double)
            Return sqlDbl(f)
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
            'ElseIf o.GetType() Is Type.GetType("System.Byte[]") Then 'upsize_ts
            '    Return "NULL" 'dürfte nicht passieren
        Else
            Throw New System.Exception(My.Resources.resMain.MsgUnknownDataType.Replace("{0}", o.GetType().ToString()))
        End If
    End Function

    Public Function sqlValues(ByVal Values As List(Of String)) As String
        Dim SQL As String = ""

        For Each Value As String In Values
            tex.Cat(SQL, sqlValue(Value), ",")
        Next

        Return SQL
    End Function

    Friend Overridable Function SqlCondEqual(ByVal rsf As clsRecordsetfield) As String
        Dim h As String = sqlValue(rsf.FieldValue)

        If h = "NULL" Then Return " " & rsf.FieldName & " IS NULL"

        If TypeOf Me Is clsDatabaseSQLServer Then
            If rsf.FieldValue.GetType() Is Type.GetType("System.String") Then
                Return " CONVERT(NVARCHAR(" & Len(h) & ")," & rsf.FieldName & ")=" & h 'Len(h) reicht wegen Hochkommata
            End If
        End If

        Return " " & rsf.FieldName & "=" & h
    End Function

    Public Function sqlExecuteBatch(ByVal sql As String) As Boolean
        Dim ret As Boolean = True

        If TypeOf Me Is clsDatabaseAccess Then
            Dim i As Integer
            For i = 1 To tex.PartCount(sql, vbCrLf)
                Dim s As String = tex.Part(sql, i, vbCrLf)
                If sqlExecute(s) = -1 Then ret = False
            Next
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            sqlExecute(sql)
        End If

        Return ret
    End Function

    Public Function sqlExecute(ByVal sql As String) As Integer
        Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
        If tmpCon.Connection.State <> ConnectionState.Open Then ConnectionPool.CloseCon(tmpCon) : Return -1

        Dim cmd As System.Data.Common.DbCommand = Nothing
        If TypeOf Me Is clsDatabaseAccess Then
            cmd = New OleDbCommand(sql, CType(tmpCon.Connection, OleDbConnection))
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            cmd = New SqlCommand(sql, CType(tmpCon.Connection, SqlConnection))
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            cmd = New OleDbCommand(sql, CType(tmpCon.Connection, OleDbConnection))
        End If

        cmd.CommandTimeout = 90 'Sekunden (Standard=30)

        If sql.Length > DBPacketSize4Varbinary Then 'String wird doppelt so groß wie Paketgröße
            cmd.CommandTimeout = 600 ' 10 Min
        End If

        Dim wiederholen As Boolean
        Dim anzWiederhol As Integer = 5

        Dim err_no As Integer = 0
        Dim err_txt As String = ""
        Dim innerExp As Exception = Nothing

        Dim ret As Integer = -1

        Do
            'ggf. versuchen, die DB erneut zu öffnen
            Select Case err_no
                Case -2147467259, -2146232060, 3709, -2147217865, -2146232060
                    If Not ConnectionPool.OpenCon_withLoop(tmpCon.Connection) Then ConnectionPool.CloseCon(tmpCon) : Return -1
            End Select

            Try
                ret = cmd.ExecuteNonQuery()
                err_no = 0
                err_txt = ""
            Catch SqlExp As SqlClient.SqlException
                err_no = SqlExp.ErrorCode
                err_txt = SqlExp.Message
                innerExp = SqlExp
            Catch OleExp As OleDb.OleDbException
                '-2147467259 = Fehler beim Verbinden
                '-2147467259 = Allgemeiner Netzwerkfehler
                '-2147467259 = SQL Server existiert nicht oder Zugriff verweigert.
                '-2147467259 = Protokollfehler im TDS-Datenstrom
                '-2147467259 = Zu viele Felder definiert (evtl. nur bei Access)
                '-2147217871 = Timeout abgelaufen
                '-2147217904 = Fehler im SQL (evtl. nur bei Access)
                '3709        = The connection cannot be used to perform this operation. It is either closed or invalid in this context.
                '-2147217865 = The Microsoft Jet database engine cannot find the input table or query '...'.  Make sure it exists and that its name is spelled correctly.
                err_no = OleExp.ErrorCode
                err_txt = OleExp.Message
                innerExp = OleExp
            Catch Exp As Exception
                err_no = 99999
                err_txt = Exp.Message
                innerExp = Exp
            End Try

            wiederholen = False
            If err_no <> 0 Then
                Select Case err_no
                    Case -2147467259, -2147217871, 3709, -2147217865, -2146232060
                        If anzWiederhol > 0 Then
                            wiederholen = True
                            anzWiederhol = anzWiederhol - 1
                            Sleep(5)
                        Else
                            Dim f As String
                            f = My.Resources.resMain.MsgExecutionOfSQLCommandFailed
                            f += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                            f += vbCrLf & My.Resources.resMain.MsgPleaseWaitSomeSecondsMinutesToRepeatExecution
                            'f += vbCrLf & GetSprachBez("(Das Abbrechen kann zu Datenverlust führen.)", "(Cancelling may result in data loss.)")
                            'If MsgBox(f, MsgBoxStyle.RetryCancel, GetSprachBez("Fehler", "Error")) = MsgBoxResult.Retry Then wiederholen = True
                            Throw New Exception(err_txt, innerExp)
                            'If ShowRetry(f) Then wiederholen = True
                        End If
                    Case Else
                        Dim m As String
                        m = My.Resources.resMain.MsgExecutionOfSQLCommandFailed
                        m += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                        'm += vbCrLf & GetSprachBez("Bitte überprüfen Sie Ihre Daten, da es unter Umständen zu einem Datenverlust gekommen ist.", "Please check your data since the error could result in data loss.")
                        clsShow.ErrorMsg(m)
                End Select

                If Not wiederholen Then
                    Dim h As String
                    h = My.Resources.resMain.MsgErrorWhileExecutingSQLCommand
                    h += vbCrLf & Me.XWords.Replace(sql)
                    h += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                    'h += vbCrLf & "Time to failure: " & Format$(p.TimeElapsed, "0.0") & " ms"

                    Me.LastError = h
                    'TODO: clsLog.LogLine(h)
                    'AutoEMail("db-error@r-c-i.de", "", "DB-Error (clsDatabase) - " & My.Application.Info.Title, GetBenutzerName() & vbCrLf & vbCrLf & h, "", "")
                    'session.AutoEMailDbError(h, False, False)
                End If
            End If
        Loop While wiederholen

        ConnectionPool.CloseCon(tmpCon)

        Return ret
    End Function

    Public Function sqlExecute(ByVal sql As String, ByVal SqlParList As List(Of clsSqlParameter)) As Integer
        Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
        If tmpCon.Connection.State <> ConnectionState.Open Then ConnectionPool.CloseCon(tmpCon) : Return -1

        Dim cmd As System.Data.Common.DbCommand = Nothing
        If TypeOf Me Is clsDatabaseAccess Then
            cmd = New OleDbCommand(sql, CType(tmpCon.Connection, OleDbConnection))
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            cmd = New SqlCommand(sql, CType(tmpCon.Connection, SqlConnection))
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            cmd = New OleDbCommand(sql, CType(tmpCon.Connection, OleDbConnection))
        End If

        cmd.CommandTimeout = 90 'Sekunden (Standard=30)

        Dim SizeInBytes As Integer = 0
        For Each Parameter As clsSqlParameter In SqlParList
            SizeInBytes += Parameter.ParValue.Length
        Next
        If SizeInBytes > 1024 * 1024 Then cmd.CommandTimeout = 600 ' Ab 1 MB Timeout auf 10 Min

        '-----
        cmd.Parameters.Clear()
        Dim Par As clsSqlParameter
        For Each Par In SqlParList
            If TypeOf Me Is clsDatabaseAccess Then
                Dim p As OleDbParameter
                p = New OleDbParameter(Par.ParName, OleDbType.VarBinary, Par.ParValue.Length)
                p.Value = Par.ParValue
                cmd.Parameters.Add(p)
            ElseIf TypeOf Me Is clsDatabaseSQLServer Then
                Dim p As SqlParameter
                p = New SqlParameter(Par.ParName, SqlDbType.VarBinary, Par.ParValue.Length)
                p.Value = Par.ParValue
                cmd.Parameters.Add(p)
            ElseIf TypeOf Me Is clsDatabaseOracle Then
                Dim p As SqlParameter
                p = New SqlParameter(Par.ParName, SqlDbType.VarBinary, Par.ParValue.Length)
                p.Value = Par.ParValue
                cmd.Parameters.Add(p)
            End If
        Next
        '-----
        Dim wiederholen As Boolean = False
        Dim anzWiederhol As Integer = 5

        Dim err_no As Integer = 0
        Dim err_txt As String = ""
        Dim innerExp As Exception = Nothing

        Dim ret As Integer = -1

        Do
            'ggf. versuchen, die DB erneut zu öffnen
            Select Case err_no
                Case -2147467259, -2146232060, 3709, -2147217865, -2146232060
                    If Not ConnectionPool.OpenCon_withLoop(tmpCon.Connection) Then ConnectionPool.CloseCon(tmpCon) : Return -1
            End Select

            Try
                ret = cmd.ExecuteNonQuery()
                err_no = 0
                err_txt = ""
            Catch SqlExp As SqlClient.SqlException
                err_no = SqlExp.ErrorCode
                err_txt = SqlExp.Message
                innerExp = SqlExp
            Catch OleExp As OleDb.OleDbException
                '-2147467259 = Fehler beim Verbinden
                '-2147467259 = Allgemeiner Netzwerkfehler
                '-2147467259 = SQL Server existiert nicht oder Zugriff verweigert.
                '-2147467259 = Protokollfehler im TDS-Datenstrom
                '-2147467259 = Zu viele Felder definiert (evtl. nur bei Access)
                '-2147217871 = Timeout abgelaufen
                '-2147217904 = Fehler im SQL (evtl. nur bei Access)
                '3709        = The connection cannot be used to perform this operation. It is either closed or invalid in this context.
                '-2147217865 = The Microsoft Jet database engine cannot find the input table or query '...'.  Make sure it exists and that its name is spelled correctly.
                err_no = OleExp.ErrorCode
                err_txt = OleExp.Message
                innerExp = OleExp
            Catch Exp As Exception
                err_no = 99999
                err_txt = Exp.Message
                innerExp = Exp
            End Try

            wiederholen = False
            If err_no <> 0 Then
                Select Case err_no
                    Case -2147467259, -2147217871, 3709, -2147217865, -2146232060
                        If anzWiederhol > 0 Then
                            wiederholen = True
                            anzWiederhol = anzWiederhol - 1
                            Sleep(5)
                        Else
                            Dim f As String
                            f = My.Resources.resMain.MsgExecutionOfSQLCommandFailed
                            f += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                            f += vbCrLf & My.Resources.resMain.MsgPleaseWaitSomeSecondsMinutesToRepeatExecution
                            'f += vbCrLf & GetSprachBez("(Das Abbrechen kann zu Datenverlust führen.)", "(Cancelling may result in data loss.)")
                            'If MsgBox(f, MsgBoxStyle.RetryCancel, GetSprachBez("Fehler", "Error")) = MsgBoxResult.Retry Then wiederholen = True
                            Throw New Exception(err_txt, innerExp)
                            'If ShowRetry(f) Then wiederholen = True
                        End If
                    Case Else
                        Dim m As String
                        m = My.Resources.resMain.MsgExecutionOfSQLCommandFailed
                        m += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                        'm += vbCrLf & GetSprachBez("Bitte überprüfen Sie Ihre Daten, da es unter Umständen zu einem Datenverlust gekommen ist.", "Please check your data since the error could result in data loss.")
                        clsShow.ErrorMsg(m)
                End Select

                If Not wiederholen Then
                    Dim h As String
                    h = My.Resources.resMain.MsgErrorWhileExecutingSQLCommand
                    h += vbCrLf & Me.XWords.Replace(sql)
                    h += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                    h += vbCrLf & My.Resources.resMain.TextSizeBKBMB.Replace("{0}", SizeInBytes).Replace("{1}", SizeInBytes / 1024).Replace("{2}", SizeInBytes / 1024 / 1024)
                    'h += vbCrLf & "Time to failure: " & Format$(p.TimeElapsed, "0.0") & " ms"

                    Me.LastError = h
                    'TODO: clsLog.LogLine(h)
                    'AutoEMail("db-error@r-c-i.de", "", "DB-Error (clsDatabase) - " & My.Application.Info.Title, GetBenutzerName() & vbCrLf & vbCrLf & h, "", "")
                    'session.AutoEMailDbError(h, False, False)
                End If
            End If
        Loop While wiederholen

        ConnectionPool.CloseCon(tmpCon)

        Return ret
    End Function

    Public Function sqlGetObject(ByVal sql As String) As Object
        Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
        If tmpCon.Connection.State <> ConnectionState.Open Then ConnectionPool.CloseCon(tmpCon) : Return Nothing

        Dim cmd As System.Data.Common.DbCommand = Nothing
        If TypeOf Me Is clsDatabaseAccess Then
            cmd = New OleDbCommand(sql, CType(tmpCon.Connection, OleDbConnection))
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            cmd = New SqlCommand(sql, CType(tmpCon.Connection, SqlConnection))
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            cmd = New OleDbCommand(sql, CType(tmpCon.Connection, OleDbConnection))
        End If

        cmd.CommandTimeout = 120 'Sekunden (Standard=30)

        Dim wiederholen As Boolean
        Dim anzWiederhol As Integer = 5

        Dim err_no As Integer = 0
        Dim err_txt As String = ""

        Dim ret As Object = Nothing

        Do
            'ggf. versuchen, die DB erneut zu öffnen
            Select Case err_no
                Case -2147467259, -2146232060, 3709, -2147217865, -2146232060
                    If Not ConnectionPool.OpenCon_withLoop(tmpCon.Connection) Then ConnectionPool.CloseCon(tmpCon) : Return Nothing
            End Select

            Try
                ret = cmd.ExecuteScalar()
                err_no = 0
                err_txt = ""
            Catch SqlExp As SqlClient.SqlException
                err_no = SqlExp.ErrorCode
                err_txt = SqlExp.Message
            Catch OleExp As OleDb.OleDbException
                err_no = OleExp.ErrorCode
                err_txt = OleExp.Message
            Catch Exp As Exception
                err_no = 99999
                err_txt = Exp.Message
            End Try

            wiederholen = False
            If err_no <> 0 Then
                Select Case err_no
                    Case -2147467259, -2147217871, 3709, -2147217865, -2146232060
                        If anzWiederhol > 0 Then
                            wiederholen = True
                            anzWiederhol = anzWiederhol - 1
                            Sleep(5)
                        Else
                            Dim f As String
                            f = My.Resources.resMain.MsgExecutionOfSQLQueryFailed
                            f += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                            f += vbCrLf & My.Resources.resMain.MsgPleaseWaitSomeSecondsMinutesToRepeatExecution
                            'f += vbCrLf & GetSprachBez("(Das Abbrechen kann zu Datenverlust führen.)", "(Cancelling may result in data loss.)")
                            'If MsgBox(f, vbRetryCancel, GetSprachBez("Fehler", "Error")) = vbRetry Then wiederholen = True
                            Throw New Exception(err_txt)
                            'If ShowRetry(f) Then wiederholen = True
                        End If
                    Case Else
                        Dim m As String
                        m = My.Resources.resMain.MsgExecutionOfSQLQueryFailed
                        m += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                        'm += vbCrLf & GetSprachBez("Bitte überprüfen Sie Ihre Daten, da es unter Umständen zu einem Datenverlust gekommen ist.", "Please check your data since the error could result in data loss.")
                        clsShow.ErrorMsg(m)
                End Select

                If Not wiederholen Then
                    Dim h As String
                    h = My.Resources.resMain.MsgErrorWhileOpeningRecordSet
                    h += vbCrLf & Me.XWords.Replace(sql)
                    h += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                    'h += vbCrLf & "Time to failure: " & Format$(p.TimeElapsed, "0.0") & " ms"

                    Me.LastError = h
                    'TODO: clsLog.LogLine(h)
                    'AutoEMail("db-error@r-c-i.de", "", "DB-Error (clsDatabase) - " & My.Application.Info.Title, GetBenutzerName() & vbCrLf & vbCrLf & h, "", "")
                    'session.AutoEMailDbError(h, False, False)
                End If
            End If
        Loop While wiederholen

        ConnectionPool.CloseCon(tmpCon)

        Return ret
    End Function

    Public Function sqlGetByte(ByVal sql As String) As Byte()
        Dim o As Object = sqlGetObject(sql)
        If o Is Nothing OrElse IsDBNull(o) Then Return Nothing
        Return CType(o, Byte())
    End Function

    Public Function sqlGetStr(ByVal sql As String) As String
        Dim o As Object = sqlGetObject(sql)
        If o Is Nothing OrElse IsDBNull(o) Then Return ""
        Return CType(o, String)
    End Function
    Public Function sqlGetLng(ByVal sql As String) As Integer
        Dim o As Object = sqlGetObject(sql)
        If o Is Nothing OrElse IsDBNull(o) Then Return 0
        Return CType(o, Integer)
    End Function
    Public Function sqlGetDbl(ByVal sql As String) As Double
        Dim o As Object = sqlGetObject(sql)
        If o Is Nothing OrElse IsDBNull(o) Then Return 0
        Return CType(o, Double)
    End Function
    Public Function sqlGetDat(ByVal sql As String) As Date
        Dim o As Object = sqlGetObject(sql)
        If o Is Nothing OrElse IsDBNull(o) Then Return dat.NullDate()
        Return CType(o, Date)
    End Function

    Public Function sqlGetLngList(ByVal sql As String) As String
        Dim ret As String = ""

        Using dr As New clsDataReader
            If dr.OpenReadonly(Me, sql) Then
                While dr.Read()
                    tex.Cat(ret, dr.getLng(0).ToString(), ",")
                End While
            End If
        End Using

        Return ret
    End Function

    Public Function sqlGetStrList(ByVal sql As String) As String
        Dim ret As String = ""

        Using dr As New clsDataReader
            If dr.OpenReadonly(Me, sql) Then
                While dr.Read()
                    tex.Cat(ret, dr.getStr(0).ToString(), ",")
                End While
            End If
        End Using

        Return ret
    End Function

    Public Function sqlGetStrArray(ByVal sql As String) As String()
        Dim ret As New List(Of String)

        Using dr As New clsDataReader
            If dr.OpenReadonly(Me, sql) Then
                While dr.Read()
                    ret.Add(dr.getStr(0).ToString())
                End While
            End If
        End Using

        Return ret.ToArray
    End Function

    Public Function TableExist(ByVal tabname As String) As Boolean
        Dim ret As Boolean = False

        If TypeOf Me Is clsDatabaseAccess Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, tabname, "TABLE"})
            'Array: TABLE_CATALOG (SQLServer: Database), TABLE_SCHEMA (SQLServer: Tableowner), TABLE_NAME, TABLE_TYPE
            ret = (schemaTable.Rows.Count > 0)
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, SqlConnection).GetSchema(SqlClientMetaDataCollectionNames.Tables, New String() {Nothing, Nothing, tabname, Nothing})
            'Array: Database/Catalog Name, Tableowner/Schema Name (z.B. "dbo"), Table Name, Table Type (z.B. "BASE TABLE")
            ret = (schemaTable.Rows.Count > 0)
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung

            'If CType(Me, clsDatabaseSQLServer).SQL2005 Then
            '    Dim s As String
            '    s = "SELECT COUNT(*)"
            '    s += " FROM sys.tables" & Me.WithNoLock
            '    s += " WHERE sys.tables.name=" & Me.sqlStr(tabname) 
            '    ret = (sqlGetLng(s) > 0)
            'Else
            '    Dim s As String
            '    s = "SELECT COUNT(*) FROM dbo.sysobjects" & Me.WithNoLock & " WHERE id = object_id('" & tabname & "')"
            '    ret = (sqlGetLng(s) > 0)
            'End If
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            ret = False
        End If

        Return ret
    End Function

    Public Function TableNames() As String()
        Dim l As New List(Of String)

        If TypeOf Me Is clsDatabaseAccess Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon Is Nothing Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
            'Array: TABLE_CATALOG (SQLServer: Database), TABLE_SCHEMA (SQLServer: Tableowner), TABLE_NAME, TABLE_TYPE
            Dim i As Integer
            For i = 0 To schemaTable.Rows.Count - 1
                l.Add(Obj2Str(schemaTable.Rows(i).Item(2)))
            Next
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, SqlConnection).GetSchema(SqlClientMetaDataCollectionNames.Tables, New String() {Nothing, Nothing, Nothing, "BASE TABLE"})
            'Array: Database/Catalog Name, Tableowner/Schema Name (z.B. "dbo"), Table Name, Table Type (z.B. "BASE TABLE")
            Dim i As Integer
            For i = 0 To schemaTable.Rows.Count - 1
                l.Add(Obj2Str(schemaTable.Rows(i).Item(2)))
            Next
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung

            'Dim s As String
            's = "SELECT * "
            's += "FROM INFORMATION_SCHEMA.TABLES" & session.db.WithNoLock
            's += " WHERE TABLE_TYPE='BASE TABLE'"
            's += "ORDER BY TABLE_NAME"
            'Using dr As New clsDataReader : dr.OpenReadonly(Me, s)
            'While dr.Read
            '    l.Add(dr.getStr("TABLE_NAME"))
            'End While
            'End Using
        End If

        Return l.ToArray()
    End Function

    Public Function FieldExist(ByVal tabname As String, ByVal fldname As String) As Boolean
        Dim ret As Boolean = False

        If TypeOf Me Is clsDatabaseAccess Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, tabname, fldname})
            'Array: TABLE_CATALOG (SQLServer: Database/Catalog Name), TABLE_SCHEMA (SQLServer: Tableowner/Schema Name z.B. "dbo"), TABLE_NAME, COLUMN_NAME
            ret = (schemaTable.Rows.Count > 0)
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, SqlConnection).GetSchema(SqlClientMetaDataCollectionNames.Columns, New String() {Nothing, Nothing, tabname, fldname})
            'Array: Database/Catalog Name, Tableowner/Schema Name (z.B. "dbo"), Table Name, Field Name
            ret = (schemaTable.Rows.Count > 0)
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung

            'If CType(Me, clsDatabaseSQLServer).SQL2005 Then
            '    Dim s As String
            '    s = "SELECT COUNT(*)"
            '    s += " FROM sys.tables" & Me.WithNoLock
            '    s += " INNER JOIN sys.columns" & Me.WithNoLock & " ON sys.tables.object_id=sys.columns.object_id"
            '    s += " WHERE sys.tables.name=" & Me.sqlStr(tabname) 
            '    s += "   AND sys.columns.name=" & Me.sqlStr(fldname)
            '    ret = (sqlGetLng(s) > 0)
            'Else
            '    Dim s As String
            '    s = "SELECT COUNT(*) FROM dbo.syscolumns" & Me.WithNoLock & " WHERE id=object_id('" & tabname & "') AND name=" & Me.sqlStr(fldname ) 
            '    ret = (sqlGetLng(s) > 0)
            'End If
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            ret = False
        End If

        Return ret
    End Function

    Public Function FieldLength(ByVal tabname As String, ByVal fldname As String) As Integer
        Dim ret As Integer = 0

        If TypeOf Me Is clsDatabaseAccess Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon Is Nothing Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, tabname, fldname})
            'Array: TABLE_CATALOG (SQLServer: Database), TABLE_SCHEMA (SQLServer: Tableowner), TABLE_NAME, COLUMN_NAME
            ret = CInt(schemaTable.Rows(0).Item("prec"))
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
            tmpCon = Nothing
            schemaTable.Dispose()
            schemaTable = Nothing
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            Dim s As String
            s = "SELECT prec FROM dbo.syscolumns" & Me.WithNoLock & " WHERE id=object_id('" & tabname & "') AND name='" & fldname & "'"
            ret = sqlGetLng(s)
        End If

        Return ret
    End Function

    Public Sub PrintTablesAndFields()
        Dim ret As String = ""

        If TypeOf Me Is clsDatabaseAccess Then
            '...
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then

            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, SqlConnection).GetSchema(SqlClientMetaDataCollectionNames.Columns, New String() {Nothing, Nothing, Nothing, Nothing})
            'Array: Database/Catalog Name, Tableowner/Schema Name (z.B. "dbo"), Table Name, Field Name
            Dim i As Integer
            For i = 0 To schemaTable.Rows.Count - 1
                Dim h As String = ""
                tex.Cat(h, Obj2Str(schemaTable.Rows(i).Item(2)), vbTab)
                tex.Cat(h, Obj2Str(schemaTable.Rows(i).Item(3)), vbTab)
                tex.Cat(h, Obj2Str(schemaTable.Rows(i).Item(7)) & IIf(Obj2Str(schemaTable.Rows(i).Item(7)) = "nvarchar", "(" & Obj2Lng(schemaTable.Rows(i).Item(8)) & ")", "").ToString(), vbTab)

                tex.Cat(ret, h, vbCrLf)
            Next
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung

            'Dim s As String
            's = "SELECT sys.tables.name AS tabName, sys.columns.name AS fieldName, sys.types.name AS typeName, sys.columns.max_length"
            's += " FROM sys.tables" & Me.WithNoLock
            's += " INNER JOIN sys.columns" & Me.WithNoLock & " ON sys.tables.object_id=sys.columns.object_id"
            's += " INNER JOIN sys.types" & Me.WithNoLock & " ON sys.columns.user_type_id=sys.types.user_type_id"
            's += " WHERE is_ms_shipped=0"
            's += " ORDER BY sys.tables.name, sys.columns.column_id"
            'Using dr As New clsDataReader : dr.OpenReadonly(session.db, s)
            'While dr.Read
            '    Dim h As String = ""
            '    tex.Cat(h, dr.getStr("tabName"), vbTab)
            '    tex.Cat(h, dr.getStr("fieldName"), vbTab)
            '    tex.Cat(h, dr.getStr("typeName") & IIf(dr.getStr("typeName") = "nvarchar", "(" & dr.getLng("max_length") & ")", ""), vbTab)

            '    tex.Cat(ret, h, vbCrLf)
            'End While
            'End Using
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            '...
        End If

        Clipboard.Clear()
        Clipboard.SetText(ret)
    End Sub

    Public Function FieldCount(ByVal tabname As String) As Integer
        Dim ret As Integer = 0

        If TypeOf Me Is clsDatabaseAccess Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, tabname, Nothing})
            'Array: TABLE_CATALOG (SQLServer: Database/Catalog Name), TABLE_SCHEMA (SQLServer: Tableowner/Schema Name z.B. "dbo"), TABLE_NAME, COLUMN_NAME
            ret = schemaTable.Rows.Count
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, SqlConnection).GetSchema(SqlClientMetaDataCollectionNames.Columns, New String() {Nothing, Nothing, tabname, Nothing})
            'Array: Database/Catalog Name, Tableowner/Schema Name (z.B. "dbo"), Table Name, Field Name
            ret = schemaTable.Rows.Count
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            ret = -1
        End If

        Return ret
    End Function

    Private Sub PrintDataTable(ByVal dt As DataTable)
        Dim ret As String = ""

        Dim s As String = ""
        Dim j As Integer
        For j = 0 To dt.Columns.Count - 1
            tex.Cat(s, dt.Columns(j).Caption, vbTab)
        Next
        tex.Cat(ret, s, vbCrLf)

        Dim i As Integer
        For i = 0 To dt.Rows.Count - 1
            s = ""
            For j = 0 To dt.Columns.Count - 1
                If j > 0 Then s += vbTab
                s = s & dt.Rows(i).Item(j).ToString()
            Next
            tex.Cat(ret, s, vbCrLf)
        Next

        'TODO: OpenInNotepad(ret)
    End Sub

    Public Function QueryCount() As Integer
        Dim ret As Integer : ret = 0

        If TypeOf Me Is clsDatabaseAccess Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon Is Nothing Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Views, New Object() {Nothing, Nothing, Nothing})
            ret = schemaTable.Rows.Count
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            '
        End If

        QueryCount = ret
    End Function

    Public Function QueryNames() As String()
        Dim l As New Generic.List(Of String)

        If TypeOf Me Is clsDatabaseAccess Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon Is Nothing Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Views, New Object() {Nothing, Nothing, Nothing})
            Dim i As Integer
            For i = 0 To schemaTable.Rows.Count - 1
                l.Add(Obj2Str(schemaTable.Rows(i).Item(2)))
            Next
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon Is Nothing Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, SqlConnection).GetSchema(SqlClientMetaDataCollectionNames.Views, New String() {Nothing, Nothing, Nothing})
            Dim i As Integer
            For i = 0 To schemaTable.Rows.Count - 1
                l.Add(Obj2Str(schemaTable.Rows(i).Item(2)))
            Next
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung

            'Dim s As String
            's = "SELECT name"
            's += " FROM sys.views" & session.db.WithNoLock
            'Using dr As New clsDataReader : dr.OpenReadonly(Me, s)
            'While dr.Read
            '    l.Add(dr.getStr("name"))
            'End While
            'End Using
        End If

        Return l.ToArray()
    End Function

    Public Sub AddField(ByVal tabname As String, ByVal colnam As String, ByVal coltyp As String)
        AddField(tabname, colnam, coltyp, False)
    End Sub

    Public Sub AddField(ByVal tabname As String, ByVal colnam As String, ByVal coltyp As String, ByVal primkey As Boolean)
        Dim s As String

        If Not TableExist(tabname.Replace("[", "").Replace("]", "")) Then
            If TypeOf Me Is clsDatabaseSQLServer Then
                s = "CREATE TABLE " & tabname
                s += " (" & colnam
                If coltyp = "AUTOINCREMENT" Then coltyp = "INT IDENTITY (1, 1)"
                s += " " & coltyp
                If primkey Then s += " NOT NULL" Else s += " NULL"
                If primkey Then s += "," & vbCrLf & "CONSTRAINT PK_" & tabname.Replace("[", "").Replace("]", "") & " PRIMARY KEY  NONCLUSTERED (" & colnam & ")"
                s += ")"
                s += vbCrLf & "EXEC sp_changeobjectowner '" & tabname & "', 'dbo'"
            Else
                s = "CREATE TABLE " & tabname
                s += " (" & colnam
                s += " " & coltyp
                If primkey Then s += " PRIMARY KEY"
                s += ")"
            End If
        Else
            s = "ALTER TABLE " & tabname
            If TypeOf Me Is clsDatabaseSQLServer Then
                If coltyp = "AUTOINCREMENT" Then
                    coltyp = "INT IDENTITY (1, 1) NOT NULL"
                Else
                    If Left$(coltyp, 4) = "TEXT" Then
                        coltyp = coltyp.Replace("TEXT", "NVARCHAR") & " COLLATE Latin1_General_CI_AS"
                    ElseIf coltyp = "MEMO" Then
                        coltyp = "NTEXT COLLATE Latin1_General_CI_AS"
                    ElseIf coltyp = "INTEGER" Then
                        coltyp = "INT"
                    End If
                    coltyp += " NULL"
                End If
            End If

            s += " ADD " & colnam & " " & coltyp

            If primkey Then
                s += vbCrLf & "ALTER TABLE " & tabname & " ADD CONSTRAINT PK_" & tabname.Replace("[", "").Replace("]", "") & " PRIMARY KEY  NONCLUSTERED (" & colnam & ")"
            End If
        End If

        sqlExecute(s)
    End Sub

    Public Sub FieldRename(ByVal tabname As String, ByVal Oldfldname As String, ByVal Newfldname As String)
        Dim s As String

        s = "SP_RENAME '" & tabname & "." & Oldfldname & "','" & Newfldname & "'"
        sqlExecute(s)
    End Sub

    Public Sub TableRename(ByVal OldTabName As String, ByVal NewTabName As String)
        Dim s As String

        s = "SP_RENAME '" & OldTabName & "','" & NewTabName & "'"
        sqlExecute(s)
    End Sub

    Public Function TableContainsFieldtype(ByVal tabname As String, ByVal fieldtype As String) As Boolean
        Dim TableInfo As SortedList(Of String, String) = GetTableInformation(tabname)

        For i = 0 To TableInfo.Count - 1
            If TableInfo.Values(i).ToUpper.Contains(fieldtype.ToUpper) Then Return True
        Next

        Return False
    End Function

    Public Sub ChangeFieldType(ByVal tabname As String, ByVal colnam As String, ByVal coltyp As String)
        Dim s As String

        s = "ALTER TABLE " & tabname
        s += " ALTER COLUMN " & colnam
        If TypeOf Me Is clsDatabaseSQLServer Then
            If Left$(coltyp, 4) = "TEXT" Then
                coltyp = coltyp.Replace("TEXT", "NVARCHAR") & " COLLATE Latin1_General_CI_AS NULL"
            ElseIf coltyp = "MEMO" Then
                coltyp = "NTEXT COLLATE Latin1_General_CI_AS NULL"
            ElseIf coltyp = "INTEGER" Then
                coltyp = "INT"
            End If
        End If
        s += " " & coltyp

        sqlExecute(s)
    End Sub

    Public Sub DelField(ByVal tabname As String, ByVal fldname As String)
        Dim s As String

        s = "ALTER TABLE " & tabname
        s += " DROP COLUMN " & fldname
        sqlExecute(s)

        Dim TabelleLoeschen As Boolean = (Me.FieldCount(tabname) = 0)

        If TabelleLoeschen Then
            s = "DROP TABLE " & tabname
            sqlExecute(s)
        End If
    End Sub

    Public Sub DelTable(ByVal tabname As String)
        Dim s As String

        s = "DROP TABLE " & tabname
        sqlExecute(s)
    End Sub

    Public Function IndexNames(Optional ByVal tabName As String = Nothing) As String()
        Dim l As New List(Of String)

        If TypeOf Me Is clsDatabaseAccess Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, OleDbConnection).GetOleDbSchemaTable(OleDbSchemaGuid.Indexes, New Object() {Nothing, Nothing, Nothing, Nothing, tabName})
            'Array: TABLE_CATALOG (SQLServer: Database), TABLE_SCHEMA (SQLServer: Tableowner), INDEX_NAME, TYPE, TABLE_NAME
            Dim i As Integer
            For i = 0 To schemaTable.Rows.Count - 1
                l.Add(Obj2Str(schemaTable.Rows(i).Item(5)) & "." & Obj2Str(schemaTable.Rows(i).Item(2)))
            Next
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, SqlConnection).GetSchema(SqlClientMetaDataCollectionNames.Indexes, New String() {Nothing, Nothing, tabName, Nothing})
            'Array: Database/Catalog Name, Tableowner/Schema Name (z.B. "dbo"), Table Name, Index Name
            Dim i As Integer
            For i = 0 To schemaTable.Rows.Count - 1
                l.Add(Obj2Str(schemaTable.Rows(i).Item(5)) & "." & Obj2Str(schemaTable.Rows(i).Item(2)))
            Next
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            '
        End If

        Return l.ToArray()
    End Function

    Public Function ForeignKeyNames(Optional ByVal tabName As String = Nothing) As String()
        Dim l As New List(Of String)

        If TypeOf Me Is clsDatabaseAccess Then
            '
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            Dim tmpCon As clsConnection = ConnectionPool.OpenCon()
            'If tmpCon.Connection.State <> ConnectionState.Open Then ...
            Dim schemaTable As DataTable = CType(tmpCon.Connection, SqlConnection).GetSchema(SqlClientMetaDataCollectionNames.ForeignKeys, New String() {Nothing, Nothing, tabName, Nothing})
            'Array: Database/Catalog Name, Tableowner/Schema Name (z.B. "dbo"), Table Name, Index Name
            Dim i As Integer
            For i = 0 To schemaTable.Rows.Count - 1
                l.Add(Obj2Str(schemaTable.Rows(i).Item(5)) & "." & Obj2Str(schemaTable.Rows(i).Item(2)))
            Next
            ConnectionPool.CloseCon(tmpCon) 'hier ohne Wirkung
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            '
        End If

        Return l.ToArray()
    End Function

    Public Sub CreateForeignKey(ByVal tb_master As String, ByVal fd_master As String, ByVal tb_detail As String, ByVal fd_detail As String)
        CreateForeignKey(tb_master, fd_master, tb_detail, fd_detail, False, False)
    End Sub

    Public Sub CreateForeignKey(ByVal tb_master As String, ByVal fd_master As String, ByVal tb_detail As String, ByVal fd_detail As String, ByVal delCasc As Boolean, ByVal updCasc As Boolean)
        Dim s As String
        If TypeOf Me Is clsDatabaseSQLServer Then
            s = "CREATE INDEX IDX_" & fd_detail & " ON " & tb_detail & " (" & fd_detail & ")"
            sqlExecute(s)
        End If
        s = "ALTER TABLE " & tb_detail
        s += " ADD CONSTRAINT FK_" & tb_detail.Replace("[", "").Replace("]", "") & "_" & fd_detail
        s += " FOREIGN KEY (" & fd_detail & ")"
        s += " REFERENCES " & tb_master & " (" & fd_master & ")"
        If delCasc Then s += " ON DELETE CASCADE"
        If updCasc Then s += " ON UPDATE CASCADE"
        sqlExecute(s)
    End Sub

    Public Sub DropForeignKey(ByVal tb_detail As String, ByVal fd_detail As String)
        DropConstraint(tb_detail, "FK_" & tb_detail & "_" & fd_detail)
    End Sub

    Public Sub DelForeignKey(ByVal tb_detail As String, ByVal fd_detail As String)
        Dim s As String
        If TypeOf Me Is clsDatabaseSQLServer Then
            s = "DROP INDEX IDX_" & fd_detail & " ON " & tb_detail
            sqlExecute(s)
        End If
        s = "ALTER TABLE " & tb_detail
        s += " DROP CONSTRAINT FK_" & tb_detail & "_" & fd_detail
        sqlExecute(s)
    End Sub

    Public Sub CreateIndex(ByVal tb As String, ByVal fd As String, ByVal isUnique As Boolean)
        Dim s As String = "CREATE"
        If isUnique Then s += " UNIQUE"
        If TypeOf Me Is clsDatabaseSQLServer Then s += " NONCLUSTERED"
        s += " INDEX IDX_" & fd.Replace(",", "_") & " ON " & tb & " (" & fd & ")"
        sqlExecute(s)
    End Sub

    Public Sub DropIndex(ByVal tb As String, ByVal idx As String)
        Dim s As String = ""
        If TypeOf Me Is clsDatabaseAccess Then
            s = "DROP INDEX " & idx & " ON " & tb
        ElseIf TypeOf Me Is clsDatabaseSQLServer Then
            s = "DROP INDEX " & tb & "." & idx
        ElseIf TypeOf Me Is clsDatabaseOracle Then
            '
        End If
        If s <> "" Then sqlExecute(s)
    End Sub

    Public Sub DropConstraint(ByVal tb As String, ByVal constr As String)
        Dim s As String
        s = "ALTER TABLE " & tb & " DROP CONSTRAINT " & constr
        sqlExecute(s)
    End Sub

    Public Sub CreateNonClusteredIndex(ByVal tabname As String, ByVal fldname As String)
        Dim s As String
        s = "CREATE NONCLUSTERED INDEX IDX_" & fldname & " ON " & tabname
        s += " (" & fldname & " Asc) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
        sqlExecute(s)
    End Sub

    Public Function IndexExist(ByVal tabname As String, ByVal fldname As String) As Boolean
        Dim s As String
        s = "SELECT COUNT(*) FROM sys.indexes WHERE object_id = OBJECT_ID(N'" & tabname & "') AND name = N'" & fldname & "'"
        Return (sqlGetLng(s) > 0)
    End Function

    Public Sub CreatePrimaryKey(ByVal tb As String, ByVal fd As String)
        Dim s As String
        s = "ALTER TABLE " & tb & " ADD CONSTRAINT PK_" & tb & " PRIMARY KEY  NONCLUSTERED (" & fd & ")"
        sqlExecute(s)
    End Sub

    Public Function StatisticExist(ByVal tabname As String, ByVal fldname As String) As Boolean
        Return (StatisticGetName(tabname, fldname) <> "")
    End Function

    Public Function StatisticGetName(ByVal tabname As String, ByVal fldname As String) As String
        Dim s As String
        s = "SELECT stat.name AS StatName"
        s += " FROM sys.stats AS stat"
        s += " INNER JOIN sys.stats_columns AS scol ON stat.object_id = scol.object_id AND stat.stats_id = scol.stats_id"
        s += " INNER JOIN sys.columns AS col ON scol.object_id = col.object_id AND col.column_id = scol.column_id"
        s += " WHERE stat.object_id = OBJECT_ID('" & tabname & "')"
        s += " AND col.name='" & fldname & "'"
        Return sqlGetStr(s)
    End Function

    Public Sub CreateStatistic(ByVal tabname As String, ByVal fldname As String, ByVal recompute As Boolean, Optional ByVal samplepercent As Integer = -1)
        Dim s As String = ""
        Dim w As String = ""
        s = "CREATE STATISTICS STAT_" & fldname & " ON " & tabname & " (" & fldname & ")"
        If samplepercent > -1 Then
            tex.Cat(w, " WITH SAMPLE " & samplepercent & " PERCENT", ",")
        Else
            tex.Cat(w, " WITH FULLSCAN", ",")
        End If
        If Not recompute Then tex.Cat(w, " NORECOMPUTE", ",")

        s = s + w
        sqlExecute(s)
    End Sub

    Public Sub DeleteStatistic(ByVal tabname As String, ByVal fldname As String)
        Dim statname As String = StatisticGetName(tabname, fldname)
        If statname = "" Then Exit Sub

        Dim s As String
        s = "DROP STATISTICS " & tabname & "." & statname
        sqlExecute(s)
    End Sub

    Public Function SqlResult(ByVal batchSQL As String, ByVal mitUmbruch255 As Boolean) As String
        Dim txt As New System.Text.StringBuilder

        Using dr As New clsDataReader : dr.OpenReadonly(session.db, batchSQL)

            Do
                Dim h As New System.Text.StringBuilder
                Dim i As Integer
                For i = 0 To dr.FieldCount - 1
                    Dim t As String = IIf(mitUmbruch255 AndAlso (i Mod 256) = 0, vbCrLf, vbTab).ToString()
                    If i > 0 Then h.Append(t)
                    h.Append(dr.getFieldName(i))
                Next
                If txt.Length > 0 Then txt.Append(vbCrLf)
                txt.Append(h.ToString)

                Dim anz As Integer = 0
                While dr.Read()
                    h = New System.Text.StringBuilder
                    For i = 0 To dr.FieldCount - 1
                        Dim w As String = ""

                        If dr.getFieldType(i) Is Type.GetType("System.String") Then
                            w = dr.getStr(i)
                        ElseIf dr.getFieldType(i) Is Type.GetType("System.Boolean") Then
                            w = IIf(dr.getInt(i).ToBoolean() = True, "Wahr", "Falsch").ToString()
                        ElseIf dr.getFieldType(i) Is Type.GetType("System.DateTime") Then
                            w = dat.printDINTime(dr.getDat(i))
                            w = w.Replace(" 00:00:00", "")
                        ElseIf dr.getFieldType(i) Is Type.GetType("System.Double") Then
                            w = zahl.printGerDecimal(dr.getDbl(i))
                        ElseIf dr.getFieldType(i) Is Type.GetType("System.Int16") Then
                            w = dr.getInt(i).ToString()
                        ElseIf dr.getFieldType(i) Is Type.GetType("System.Integer") Then
                            w = dr.getInt(i).ToString()
                        ElseIf dr.getFieldType(i) Is Type.GetType("System.Int32") Then
                            w = dr.getLng(i).ToString()
                        ElseIf dr.getFieldType(i) Is Type.GetType("System.Int64") Then
                            'ElseIf dr.getFieldType(i) Is Type.GetType("System.Long") Then
                            w = "LONG"
                        ElseIf dr.getFieldType(i) Is Type.GetType("System.Byte[]") Then 'upsize_ts
                            w = "BYTE[]"
                        Else
                            Throw New System.Exception(My.Resources.resMain.MsgUnknownDataType.Replace("{0}", dr.getFieldType(i).ToString()))
                        End If

                        Dim t As String = IIf(mitUmbruch255 AndAlso (i Mod 256) = 0, vbCrLf, vbTab).ToString()
                        If i > 0 Then h.Append(t)
                        h.Append(w)
                    Next
                    txt.Append(vbCrLf)
                    txt.Append(h.ToString)
                    anz += 1
                End While

                If anz = 0 Then txt.Append(vbCrLf) : txt.Append(My.Resources.resMain.MsgCannotFindDataset)

                txt.Append(vbCrLf) 'Leerzeile
            Loop While dr.NextResult()

        End Using

        Return txt.ToString
    End Function

    Public Function PrintTableCreateScript(tabname As String) As String
        Dim ret As String = ""

        Dim t As New clsDbTable(Me, tabname)

        Dim i As Integer
        For i = 0 To t.DbColumns.Count - 1
            Dim c As clsDbColumn = t.DbColumns.ItemByIdx(i)
            tex.Cat(ret, ".AddField(""" & t.Name & """, """ & c.Name & """, """ & c.Typ & IIf(c.Typ = "TEXT", "(" & c.Laenge & ")", "").ToString() & """" & IIf(c.Typ = "AUTOINCREMENT", ", True", "").ToString() & ")", vbCrLf)
        Next

        For i = 0 To t.DbForeignKeys.Count - 1
            Dim fk As clsDbForeignKey = t.DbForeignKeys.ItemByIdx(i)
            tex.Cat(ret, ".CreateForeignKey(""" & fk.tbMaster & """, """ & fk.fdMaster & """, """ & fk.tbDetail & """, """ & fk.fdDetail & """)", vbCrLf)
        Next

        Return ret
    End Function

    Public Function PrintTableInsertScript(tabName As String) As String
        Dim ret As String = ""

        Dim i As Integer
        Dim s As String

        Dim t As New clsDbTable(Me, tabName)

        Dim fieldnames As String = ""
        Dim nullvalues As String = ""


        For i = 0 To t.DbColumns.Count - 1
            Dim c As clsDbColumn = t.DbColumns.ItemByIdx(i)
            If c.Typ <> "AUTOINCREMENT" Then
                tex.Cat(fieldnames, c.Name, ", ")
                tex.Cat(nullvalues, "NULL", ", ")
            End If
        Next

        Dim primKey As String = t.DbColumns.ItemByIdx(0).Name
        Dim ID As Integer = 0
        Dim delIDs As String = ""

        s = "SELECT * FROM " & tabName & Me.WithNoLock & " ORDER BY " & primKey
        Using dr As New clsDataReader : dr.OpenReadonly(Me, s)
            While dr.Read
                ID += 1
                While dr.getLng(primKey) <> ID
                    s = "INSERT INTO " & tabName & " (" & fieldnames & ") VALUES (" & nullvalues & ")"
                    tex.Cat(ret, s, vbCrLf)
                    tex.Cat(delIDs, ID.ToString(), ",")
                    ID += 1
                End While

                Dim fieldvalues As String = ""

                For i = 0 To t.DbColumns.Count - 1
                    Dim c As clsDbColumn = t.DbColumns.ItemByIdx(i)
                    Select Case c.Typ
                        Case "AUTOINCREMENT"
                            'nix
                        Case Else
                            If dr.FieldNull(c.Name) Then
                                tex.Cat(fieldvalues, "NULL", ", ")
                            Else
                                Select Case c.Typ
                                    Case "INT"
                                        tex.Cat(fieldvalues, dr.getLng(c.Name).ToString(), ", ")
                                    Case "FLOAT"
                                        tex.Cat(fieldvalues, Me.sqlDbl(dr.getDbl(c.Name)), ", ")
                                    Case "DATETIME"
                                        Dim d As Date = dr.getDat(c.Name)
                                        If d = d.Date Then
                                            tex.Cat(fieldvalues, Me.sqlDate(d), ", ")
                                        Else
                                            tex.Cat(fieldvalues, Me.sqlDateTime(d), ", ")
                                        End If
                                    Case "TEXT", "MEMO"
                                        tex.Cat(fieldvalues, Me.sqlStr(dr.getStr(c.Name)), ", ")
                                    Case Else
                                        clsShow.InternalError(My.Resources.resMain.MsgErrorInPrintTableInsertScript & My.Resources.resMain.MsgUnknownDataType.Replace("{0}", c.Typ))
                                        tex.Cat(fieldvalues, "NULL", ", ")
                                End Select
                            End If
                    End Select
                Next

                s = "INSERT INTO " & tabName & " (" & fieldnames & ") VALUES (" & fieldvalues & ")"
                tex.Cat(ret, s, vbCrLf)
            End While
        End Using

        If delIDs <> "" Then
            s = "DELETE FROM " & tabName & " WHERE " & primKey & " IN (" & delIDs & ")"
            tex.Cat(ret, s, vbCrLf)
        End If

        Return ret
    End Function

    Public Function GetTableInformation(ByVal tabname As String) As SortedList(Of String, String)
        Dim ret As New SortedList(Of String, String) 'Fieldname, Datatype

        If TypeOf Me Is clsDatabaseAccess Then Return ret

        Dim s As String
        s = "SELECT sys.all_columns.name AS Fieldname, sys.types.name AS Datatype FROM sys.all_columns"
        s += " INNER JOIN sys.types" & Me.WithNoLock & " ON sys.all_columns.system_type_id = sys.types.system_type_id"
        s += " INNER JOIN sys.tables" & Me.WithNoLock & " ON sys.tables.object_id = sys.all_columns.object_id"
        s += " WHERE sys.types.name NOT LIKE 'sysname'"
        s += " AND sys.tables.name LIKE " & sqlValue(tabname)

        Using dr As New clsDataReader : dr.OpenReadonly(Me, s)
            While dr.Read()
                ret.Add(dr.getStr("Fieldname"), dr.getStr("Datatype"))
            End While
        End Using

        Return ret
    End Function

    Public Function FileDownload(ByVal Tablename As String, ByVal PrimaryKeyName As String, ByVal PrimaryKeyID As Integer, ByVal FieldName As String, ByVal Filepath As String) As Boolean

        If Not clsDirectory.Exists(IO.Path.GetDirectoryName(Filepath)) Then If Not clsDirectory.Make(IO.Path.GetDirectoryName(Filepath)) Then Return False
        Dim path As String = IO.Path.GetPathRoot(Filepath)

        Dim s As String = "SELECT DATALENGTH(" & FieldName & ") FROM " & Tablename & Me.WithNoLock
        s += " WHERE " & PrimaryKeyName & " = " & PrimaryKeyID

        Dim datalength As Long = Obj2Lng(Me.sqlGetDbl(s))
        If datalength < 1 Then Return False

        Dim drives() As IO.DriveInfo = IO.DriveInfo.GetDrives()

        For Each drive As IO.DriveInfo In drives
            If drive IsNot Nothing AndAlso (drive.DriveType = IO.DriveType.Fixed Or drive.DriveType = IO.DriveType.Network) And drive.Name.ToUpper = IO.Path.GetPathRoot(Filepath).ToUpper Then
                If drive.AvailableFreeSpace <= (datalength + DBFileDownloadDriveMinSpaceLeft) Then drives = Nothing : clsShow.ErrorMsg("Not enough space left on drive '" & drive.Name.ToUpper & "'.") : Return False
                Exit For
            End If
        Next
        drives = Nothing


        Dim ret As Boolean = False
        Dim fs As New IO.FileStream(Filepath, IO.FileMode.OpenOrCreate, IO.FileAccess.Write)
        Try
            Dim bw As IO.BinaryWriter = New IO.BinaryWriter(fs)

            Dim index As Long = 0
            Dim h As String = ""
            Dim dat As Byte()

            s = s.Replace("SELECT DATALENGTH(" & FieldName & ")", "SELECT " & FieldName)
            datalength = datalength + 2 'VarBinary als String beginnend mit 0x in der Datenbank abgelegt -> DATALENGTH() zählt diese nicht, SUBSTRING muss + 2 Byte lesen
            'Inhalt blockweise aus DB lesen
            While (datalength - index) > DBPacketSize4Varbinary
                h = s.Replace(FieldName, "SUBSTRING(" & FieldName & "," & index & "," & DBPacketSize4Varbinary & ")")
                dat = Me.sqlGetByte(h)
                If index = 0 Then
                    bw.Write(dat)
                Else
                    bw.Write(dat, 0, DBPacketSize4Varbinary)
                End If
                index += DBPacketSize4Varbinary
                Application.DoEvents()
            End While

            'Restinhalt lesen
            h = s.Replace(FieldName, "SUBSTRING(" & FieldName & "," & index & "," & datalength - index + 1 & ")")
            dat = Me.sqlGetByte(h)
            If dat IsNot Nothing Then
                If index = 0 Then
                    bw.Write(dat)
                Else
                    bw.Write(dat, 0, dat.Length)
                End If
            End If
            dat = Nothing

            bw.Flush()
            bw.Close()
            bw = Nothing

            ret = True
        Catch ex As Exception
            clsShow.Message(ex.Message.ToString)
        Finally
            fs.Close()
            fs = Nothing
        End Try

        Return ret
    End Function

    Public Function FileUpload(ByVal Tablename As String, ByVal PrimaryKeyName As String, ByVal PrimaryKeyID As Integer, ByVal FieldName As String, ByVal Filepath As String) As Boolean
        If Not datei.Exists(Filepath) Then Return False

        Dim fs As New IO.FileStream(Filepath, IO.FileMode.Open, IO.FileAccess.Read)
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(1252) 'muß der clsCryptlight entsprechen

        Dim br As New IO.BinaryReader(fs)
        Dim data As Byte()
        Dim index As Integer = 0
        Dim OK As Boolean = True
        Dim anzahl As Integer = DBPacketSize4Varbinary
        If fs.Length < anzahl Then anzahl = CInt(fs.Length)

        'Datei blockweise in DB schreiben
        Using dw As New clsDataWriter : dw.OpenEdit(Me, Tablename, PrimaryKeyName, PrimaryKeyID)
            While (fs.Length - index) > anzahl
                data = br.ReadBytes(anzahl)
                dw.SetFieldValue(FieldName, data, False, index > 0)
                If Not dw.Update() Then OK = False : Exit While
                index += anzahl
                Application.DoEvents()
            End While

            'Rest schreiben
            data = br.ReadBytes(CInt(fs.Length) - index)
            dw.SetFieldValue(FieldName, data, False, index > 0)
            If OK And Not dw.Update() Then OK = False
        End Using
        br.Close() : br = Nothing
        fs.Close() : fs = Nothing

        data = Nothing

        Return OK
    End Function

End Class

Public Class clsDbTable
    Public Name As String
    Public DbColumns As New clsDbColumns
    Public DbForeignKeys As New clsDbForeignKeys
    Public DbIndexes As New clsDbIndexes

    Public ReadOnly Property SafeName As String
        Get
            Dim tabName As String = Me.Name
            If tabName.Contains(" ") Then
                Dim r As New System.Text.RegularExpressions.Regex("^\[[a-z0-9 ]+\]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                If Not r.IsMatch(tabName) Then tabName = "[" & tabName & "]"
            End If
            Return tabName
        End Get
    End Property

    Public Sub New(ByVal db As clsDatabase, ByVal tabName As String)
        Me.Name = tabName

        If TypeOf db Is clsDatabaseAccess Then
            '...
        ElseIf TypeOf db Is clsDatabaseSQLServer Then

            Dim s As String = " SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH" & _
                              " FROM INFORMATION_SCHEMA.COLUMNS" & db.WithNoLock & _
                              " WHERE TABLE_NAME LIKE " & db.sqlValue(tabName) & _
                              " ORDER BY ORDINAL_POSITION"

            Using dr As New clsDataReader
                dr.OpenReadonly(db, s)

                While dr.Read
                    Dim c As New clsDbColumn

                    c.Name = dr.getStr("COLUMN_NAME")
                    c.Typ = dr.getStr("DATA_TYPE").ToUpper()
                    Select Case c.Typ
                        Case "NVARCHAR" : c.Typ = "TEXT"
                        Case "NTEXT" : c.Typ = "MEMO"
                        Case "INT" : c.Typ = "INT"
                    End Select

                    c.Laenge = dr.getInt("CHARACTER_MAXIMUM_LENGTH")

                    Me.DbColumns.Add(c)
                End While
            End Using

        ElseIf TypeOf db Is clsDatabaseOracle Then
            '...
        End If

        '--- DbForeignKeys

        If TypeOf db Is clsDatabaseAccess Then
            'evtl. mit GetOleDbSchemaTable
        ElseIf TypeOf db Is clsDatabaseSQLServer Then
            Dim s As String
            s = "SELECT FK.name AS constraint_name, UT.name AS tbMaster, UC.name AS fdMaster, T.name AS tbDetail, C.name AS fdDetail"
            s += " FROM sys.tables AS T"
            s += " INNER JOIN sys.foreign_keys AS FK ON T.object_id = FK.parent_object_id"
            s += " INNER JOIN sys.foreign_key_columns AS FKC ON FK.object_id = FKC.constraint_object_id"
            s += " INNER JOIN sys.columns AS C ON (FKC.parent_object_id = C.object_id AND FKC.parent_column_id = C.column_id)"
            s += " INNER JOIN sys.columns AS UC ON (FKC.referenced_object_id = UC.object_id AND FKC.referenced_column_id = UC.column_id)"
            s += " INNER JOIN sys.tables AS UT ON FKC.referenced_object_id = UT.object_id"
            s += " WHERE  T.name=" & session.db.sqlStr(tabName)
            s += " ORDER BY tbDetail, fdDetail"
            Using dr As New clsDataReader : dr.OpenReadonly(db, s)
                While dr.Read
                    Dim fk As New clsDbForeignKey
                    fk.Name = dr.getStr("constraint_name")
                    fk.tbMaster = dr.getStr("tbMaster")
                    fk.fdMaster = dr.getStr("fdMaster")
                    fk.tbDetail = dr.getStr("tbDetail")
                    fk.fdDetail = dr.getStr("fdDetail")
                    Me.DbForeignKeys.Add(fk)
                End While
            End Using
        ElseIf TypeOf db Is clsDatabaseOracle Then
            '
        End If

        '--- DbIndexes (für PrimaryKeys und sonstige)

        If TypeOf db Is clsDatabaseAccess Then
            'evtl. mit GetOleDbSchemaTable
        ElseIf TypeOf db Is clsDatabaseSQLServer Then
            Dim s As String
            s = "SELECT T.name AS table_name, IX.name AS index_name, C.name AS column_name, IX.is_primary_key"
            s += " FROM sys.tables AS T"
            s += " INNER JOIN sys.indexes AS IX ON T.object_id = IX.object_id"
            s += " INNER JOIN sys.index_columns AS IC ON (IX.object_id = IC.object_id AND IX.index_id = IC.index_id)"
            s += " INNER JOIN sys.columns AS C ON (IC.column_id = C.column_id AND IC.object_id = C.OBJECT_ID)"
            s += " WHERE T.name=" & session.db.sqlStr(tabName)
            s += " ORDER BY table_name, index_name"
            Using dr As New clsDataReader : dr.OpenReadonly(db, s)
                While dr.Read
                    Dim ix As New clsDbIndex
                    ix.Name = dr.getStr("index_name")
                    ix.tbName = dr.getStr("table_name")
                    ix.fdName = dr.getStr("column_name")
                    ix.isPrimaryKey = (dr.getInt("is_primary_key").ToBoolean() = True)
                    Me.DbIndexes.Add(ix)
                End While
            End Using
        ElseIf TypeOf db Is clsDatabaseOracle Then
            '
        End If
    End Sub
End Class

Public Class clsDbColumns
    Private l As New Generic.List(Of clsDbColumn)

    Public ReadOnly Property Count() As Integer
        Get
            Return l.Count
        End Get
    End Property

    Public Sub Add(itm As clsDbColumn)
        l.Add(itm)
    End Sub

    Public Function ItemByIdx(idx As Integer) As clsDbColumn
        Return l.Item(idx)
    End Function
End Class

Public Class clsDbColumn
    Public Name As String = ""
    Public Typ As String = ""
    Public Laenge As Integer = 0 'nur für TEXT()
End Class

Public Class clsDbForeignKeys
    Private l As New Generic.List(Of clsDbForeignKey)

    Public ReadOnly Property Count() As Integer
        Get
            Return l.Count
        End Get
    End Property

    Public Sub Add(itm As clsDbForeignKey)
        l.Add(itm)
    End Sub

    Public Function ItemByIdx(idx As Integer) As clsDbForeignKey
        Return l.Item(idx)
    End Function
End Class

Public Class clsDbForeignKey
    Public Name As String

    Public tbMaster As String
    Public fdMaster As String
    Public tbDetail As String
    Public fdDetail As String
End Class

Public Class clsDbIndexes
    Private l As New Generic.List(Of clsDbIndex)

    Public ReadOnly Property Count() As Integer
        Get
            Return l.Count
        End Get
    End Property

    Public Sub Add(itm As clsDbIndex)
        l.Add(itm)
    End Sub

    Public Function ItemByIdx(idx As Integer) As clsDbIndex
        Return l.Item(idx)
    End Function
End Class

Public Class clsDbIndex
    Public Name As String

    Public tbName As String
    Public fdName As String
    Public isPrimaryKey As Boolean
End Class

'==================================================================================================
Public Class clsDataReader
    Implements IDisposable

    Private dr As System.Data.Common.DbDataReader 'OleDb.OleDbDataReader / SqlClient.SqlDataReader
    Private drCon As clsConnection
    Private db As clsDatabase

    Private multiSql As String 'mehrere SQL-Abfragen mit vbCrLf getrennt
    Private multiIdx As Integer = 0

    Public Function OpenReadonlyMulti(ByVal data_base As clsDatabase, ByVal sql As String) As Boolean
        If TypeOf data_base Is clsDatabaseAccess Then
            Me.multiSql = sql
            Me.multiIdx = 1
            Dim s As String = tex.Part(Me.multiSql, Me.multiIdx, vbCrLf)
            Return Me.OpenReadonly(data_base, s)
        ElseIf TypeOf data_base Is clsDatabaseSQLServer Then
            Return Me.OpenReadonly(data_base, sql)
        End If

        Return False
    End Function

    Public Function OpenReadonly(ByVal data_base As clsDatabase, ByVal sql As String, Optional ByVal CommandType As System.Data.CommandType = CommandType.Text) As Boolean
        db = data_base
        drCon = db.ConnectionPool.OpenCon()
        If drCon.Connection.State <> ConnectionState.Open Then db.ConnectionPool.CloseCon(drCon) : drCon = Nothing : Return False

        Dim cmd As System.Data.Common.DbCommand = Nothing
        If TypeOf db Is clsDatabaseAccess Then
            cmd = New OleDbCommand(sql, CType(drCon.Connection, OleDbConnection))
        ElseIf TypeOf db Is clsDatabaseSQLServer Then
            cmd = New SqlCommand(sql, CType(drCon.Connection, SqlConnection))
        ElseIf TypeOf db Is clsDatabaseOracle Then
            cmd = New OleDbCommand(sql, CType(drCon.Connection, OleDbConnection))
        End If

        cmd.CommandTimeout = 90 'Sekunden (Standard=30)
        cmd.CommandType = CommandType

        Dim wiederholen As Boolean
        Dim anzWiederhol As Integer = 5

        Dim err_no As Integer = 0
        Dim err_txt As String = ""
        Dim innerExp As Exception = Nothing

        Do
            'ggf. versuchen, die DB erneut zu öffnen
            Select Case err_no
                Case -2147467259, -2146232060, 3709, -2147217865, -2146232060
                    If Not db.ConnectionPool.OpenCon_withLoop(drCon.Connection) Then db.ConnectionPool.CloseCon(drCon) : drCon = Nothing : dr = Nothing : Return False
            End Select

            Try
                dr = cmd.ExecuteReader()
                err_no = 0
                err_txt = ""
            Catch sqlExp As Data.SqlClient.SqlException
                err_no = sqlExp.ErrorCode
                err_txt = sqlExp.Message
                innerExp = sqlExp
            Catch oleExp As Data.OleDb.OleDbException
                err_no = oleExp.ErrorCode
                err_txt = oleExp.Message
                innerExp = oleExp
            Catch Exp As Exception
                err_no = 99999
                err_txt = Exp.Message
                innerExp = Exp
            End Try

            wiederholen = False
            If err_no <> 0 Then
                Select Case err_no
                    Case -2147467259, -2147217871, 3709, -2146232060, -2147217865, -2146232060
                        If anzWiederhol > 0 Then
                            wiederholen = True
                            anzWiederhol = anzWiederhol - 1
                            Sleep(5)
                        Else
                            Dim f As String
                            f = My.Resources.resMain.MsgExecutionOfSQLQueryFailed
                            f += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                            f += vbCrLf & My.Resources.resMain.MsgPleaseWaitSomeSecondsMinutesToRepeatExecution
                            'f += vbCrLf & GetSprachBez("(Das Abbrechen kann zu Datenverlust führen.)", "(Cancelling may result in data loss.)")
                            'If MsgBox(f, vbRetryCancel, GetSprachBez("Fehler", "Error")) = vbRetry Then wiederholen = True
                            Throw New Exception(err_txt, innerExp)
                            'If ShowRetry(f) Then wiederholen = True
                        End If
                    Case Else
                        Dim m As String
                        m = My.Resources.resMain.MsgExecutionOfSQLQueryFailed
                        m += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                        'm += vbCrLf & GetSprachBez("Bitte überprüfen Sie Ihre Daten, da es unter Umständen zu einem Datenverlust gekommen ist.", "Please check your data since the error could result in data loss.")
                        clsShow.ErrorMsg(m)
                End Select

                If Not wiederholen Then
                    Dim h As String
                    h = My.Resources.resMain.MsgErrorWhileExecutingSQLQuery
                    h += vbCrLf & db.XWords.Replace(sql)
                    h += vbCrLf & My.Resources.resMain.TextErrorNoText.Replace("{0}", err_no).Replace("{1}", err_txt)
                    'h += vbCrLf & "Time to failure: " & Format$(p.TimeElapsed, "0.0") & " ms"

                    db.LastError = h
                    'TODO: clsLog.LogLine(h)
                    'AutoEMail("db-error@r-c-i.de", "", "DB-Error (clsDatabase) - " & My.Application.Info.Title, GetBenutzerName() & vbCrLf & vbCrLf & h, "", "")
                    'session.AutoEMailDbError(h, False, False)
                End If
            End If
        Loop While wiederholen

        If err_no <> 0 Then db.ConnectionPool.CloseCon(drCon) : drCon = Nothing : dr = Nothing 'CloseCon ggf. mit Wirkung, da für DataReader angefordert

        Return (err_no = 0)
    End Function

    Public Function OpenQuery(ByVal data_base As clsDatabase, ByVal queryName As String) As Boolean
        If data_base.isSQLServer Then
            Dim s As String
            s = "SELECT * FROM " & queryName & data_base.WithNoLock
            Return Me.OpenReadonly(data_base, s)
        Else
            Return Me.OpenReadonly(data_base, queryName, CommandType.StoredProcedure)
        End If
    End Function

    Public Function Read() As Boolean
        If dr Is Nothing Then Return False
        Return dr.Read()
    End Function

    Public Function NextResult() As Boolean
        If TypeOf db Is clsDatabaseAccess Then
            Me.Close()
            Me.multiIdx += 1
            Dim s As String = tex.Part(Me.multiSql, Me.multiIdx, vbCrLf)
            Return Me.OpenReadonly(db, s)
        ElseIf TypeOf db Is clsDatabaseSQLServer Then
            Return dr.NextResult()
        End If
        Return True
    End Function

    Private Function Item(ByVal index As Integer) As Object
        Dim o As Object = System.DBNull.Value
        If dr.HasRows Then
            Try
                o = dr.Item(index)
            Catch ex As Exception
                clsShow.InternalError(My.Resources.resMain.MsgClsDatareaderFieldNotFound.Replace("{0}", index))
            End Try
        End If
        Return o
    End Function

    Private Function Item(ByVal name As String) As Object
        Dim o As Object = System.DBNull.Value
        If dr.HasRows Then
            Try
                o = dr.Item(name)
            Catch ex As Exception
                clsShow.InternalError(My.Resources.resMain.MsgClsDatareaderFieldNotFound.Replace("{0}", name))
            End Try
        End If
        Return o
    End Function

    Public Function FieldExist(ByVal name As String) As Boolean
        Dim o As Object = System.DBNull.Value
        Try
            o = dr.Item(name)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Function FieldCount() As Integer
        Return dr.FieldCount
    End Function

    Public Function getFieldName(ByVal index As Integer) As String
        Return dr.GetName(index)
    End Function

    Public Function getFieldType(ByVal index As Integer) As System.Type
        Return dr.GetFieldType(index)
    End Function

    Public Function getIndexForFieldName(ByVal fieldname As String) As Integer
        Dim i As Integer
        For i = 0 To FieldCount() - 1
            If getFieldName(i) = fieldname Then Return i
        Next
        Return -1
    End Function

    Public Function FieldNull(ByVal fieldname As String) As Boolean
        Return IsDBNull(Item(fieldname))
    End Function

    Public Function getObject(ByVal fieldname As String) As Object
        Return Item(fieldname)
    End Function

    Public Function getObject(ByVal index As Integer) As Object
        Return Item(index)
    End Function

    Public Function getByte(ByVal fieldname As String) As Byte()
        If IsDBNull(Item(fieldname)) Then Return Nothing
        Return CType(Item(fieldname), Byte())
    End Function

    Public Function getByte(ByVal index As Integer) As Byte()
        If IsDBNull(Item(index)) Then Return Nothing
        Return CType(Item(index), Byte())
    End Function

    Public Function getStr(ByVal fieldname As String) As String
        If IsDBNull(Item(fieldname)) Then Return ""
        Return CType(Item(fieldname), String)
    End Function

    Public Function getStr(ByVal index As Integer) As String
        If IsDBNull(Item(index)) Then Return ""
        Return CType(Item(index), String)
    End Function

    Public Function getLng(ByVal fieldname As String) As Long
        If IsDBNull(Item(fieldname)) Then Return 0
        Return CType(Item(fieldname), Long)
    End Function

    Public Function getLng(ByVal index As Integer) As Long
        If IsDBNull(Item(index)) Then Return 0
        Return CType(Item(index), Long)
    End Function

    Public Function getInt(ByVal fieldname As String) As Integer
        If IsDBNull(Item(fieldname)) Then Return 0
        Return CType(Item(fieldname), Integer)
    End Function

    Public Function getInt(ByVal index As Integer) As Integer
        If IsDBNull(Item(index)) Then Return 0
        Return CType(Item(index), Integer)
    End Function

    Public Function getDbl(ByVal fieldname As String) As Double
        If IsDBNull(Item(fieldname)) Then Return 0
        Return CType(Item(fieldname), Double)
    End Function

    Public Function getDbl(ByVal index As Integer) As Double
        If IsDBNull(Item(index)) Then Return 0
        Return CType(Item(index), Double)
    End Function

    Public Function getDat(ByVal fieldname As String) As Date
        If IsDBNull(Item(fieldname)) Then Return dat.NullDate()
        'Return CType(Item(fieldname), Date)
        Dim d As Date = CType(Item(fieldname), Date)
        If d.Year = 1900 And d.Month = 1 And d.Day = 1 Then d = TimeSerial(d.Hour, d.Minute, d.Second)
        Return d
    End Function

    Public Function getDat(ByVal index As Integer) As Date
        If IsDBNull(Item(index)) Then Return dat.NullDate()
        'Return CType(Item(index), Date)
        Dim d As Date = CType(Item(index), Date)
        If d.Year = 1900 And d.Month = 1 And d.Day = 1 Then d = TimeSerial(d.Hour, d.Minute, d.Second)
        Return d
    End Function

    Private Sub Close()
        If dr IsNot Nothing AndAlso Not dr.IsClosed() Then dr.Close() 'für den Fall, dass dem Close() kein OpenReadonly() voran ging
        If drCon IsNot Nothing Then db.ConnectionPool.CloseCon(drCon) 'ggf. mit Wirkung, da für DataReader angefordert
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' So ermitteln Sie überflüssige Aufrufe

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: Verwalteten Zustand löschen (verwaltete Objekte).
            End If

            ' TODO: Nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalize() unten überschreiben.
            ' TODO: Große Felder auf NULL festlegen.

            Close()

            dr = Nothing
            drCon = Nothing
            db = Nothing
        End If

        Me.disposedValue = True
    End Sub

    ' TODO: Finalize() nur überschreiben, wenn Dispose(ByVal disposing As Boolean) oben über Code zum Freigeben von nicht verwalteten Ressourcen verfügt.
    Protected Overrides Sub Finalize()
        ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class

'==================================================================================================
Public Class clsDataWriter
    Implements IDisposable

    Private Enum EnumRecordsetMode
        rmAddNew
        rmEdit
    End Enum

    Private db As clsDatabase
    Private TableName As String = ""
    Private PrimaryKeyName As String = ""
    Private PrimaryKeyValue As Integer = 0
    Private PrimaryKeyValueVar As String = ""
    Private Fields As Collections.ArrayList
    Private Mode As EnumRecordsetMode

    Private TableInfo As SortedList(Of String, String)
    Private VarBinaryFieldsToUpdate As Generic.Dictionary(Of String, String) 'Feldname, temporärer Dateiname

    Private PrimaryKeys_Fields As New Dictionary(Of Integer, ArrayList) ' PrimaryKeyValue, Fields
    Private PrimaryKeys_VarBinaryFieldsToUpdate As New Dictionary(Of Integer, Dictionary(Of String, String)) ' PrimaryKeyValue, (Feldname, temporärer Dateiname)

    'max. Anzahl Spalte für UPDATE/INSERT
    Private maxAnzUpd As Integer
    Private maxAnzIns As Integer

    Public Function GetPrimaryKeyValue() As Integer
        Return PrimaryKeyValue
    End Function

    Private Sub Init()
        maxAnzUpd = CInt(IIf(db.isSQLServer, 1024, 127))
        maxAnzIns = CInt(IIf(db.isSQLServer, 1024, 255)) '= max. Anzahl Tabellenspalten
    End Sub

    Public Sub OpenEdit(ByVal data_base As clsDatabase, ByVal table_Name As String, ByVal primaryKey_Name As String, ByVal primaryKey_Value As Integer)
        PrimaryKeys_Fields.Clear()
        PrimaryKeys_VarBinaryFieldsToUpdate.Clear()

        db = data_base
        TableName = table_Name
        PrimaryKeyName = primaryKey_Name
        PrimaryKeyValue = primaryKey_Value
        PrimaryKeyValueVar = ""
        Fields = New ArrayList
        TableInfo = data_base.GetTableInformation(TableName)
        VarBinaryFieldsToUpdate = New Generic.Dictionary(Of String, String)

        Mode = EnumRecordsetMode.rmEdit
        Init()
    End Sub

    'Public Sub OpenEdit(ByVal data_base As clsDatabase, ByVal table_Name As String, ByVal primaryKey_Name As String, ByVal primaryKey_ValueVar As String)
    '    PrimaryKeys_Fields.Clear()
    '    PrimaryKeys_VarBinaryFieldsToUpdate.Clear()

    '    db = data_base
    '    TableName = table_Name
    '    PrimaryKeyName = primaryKey_Name
    '    PrimaryKeyValue = 0
    '    PrimaryKeyValueVar = primaryKey_ValueVar
    '    Fields = New ArrayList
    '    TableInfo = data_base.GetTableInformation(TableName)
    '    VarBinaryFieldsToUpdate = New Generic.Dictionary(Of String, String)

    '    Mode = EnumRecordsetMode.rmEdit
    '    Init()
    'End Sub

    Public Sub OpenAddNew(ByVal data_base As clsDatabase, ByVal table_Name As String, ByVal primaryKey_Name As String)
        PrimaryKeys_Fields.Clear()
        PrimaryKeys_VarBinaryFieldsToUpdate.Clear()

        db = data_base
        TableName = table_Name
        PrimaryKeyName = primaryKey_Name
        PrimaryKeyValue = 0
        PrimaryKeyValueVar = ""
        Fields = New ArrayList
        TableInfo = data_base.GetTableInformation(TableName)
        VarBinaryFieldsToUpdate = New Generic.Dictionary(Of String, String)

        Mode = EnumRecordsetMode.rmAddNew
        Init()
    End Sub

    'TODO: 
    'Public Function OpenEditSpecial(ByVal data_base As clsDatabase, ByVal table_Name As String, ByVal primaryKey_Name As String, ByVal primaryKey_Value As Integer) As Boolean
    '    PrimaryKeys_Fields.Clear()
    '    PrimaryKeys_VarBinaryFieldsToUpdate.Clear()

    '    db = data_base
    '    TableName = table_Name
    '    PrimaryKeyName = primaryKey_Name
    '    PrimaryKeyValue = primaryKey_Value
    '    PrimaryKeyValueVar = ""
    '    Fields = New ArrayList
    '    TableInfo = data_base.GetTableInformation(TableName)
    '    VarBinaryFieldsToUpdate = New Generic.Dictionary(Of String, String)

    '    Mode = EnumRecordsetMode.rmEdit
    '    Init()

    '    'aktuelle Werte laden
    '    Dim i As Integer
    '    Dim s As String = ""
    '    Dim stdcol As String = primaryKey_Name

    '    For i = 0 To TableInfo.Count - 1
    '        If TableInfo.Keys(i) = primaryKey_Name Then Continue For
    '        If TableInfo.Values(i).ToUpper.Contains("VARBINARY") Then
    '            s = "SELECT DATALENGTH(" & TableInfo.Keys(i) & ") FROM " & TableName & db.WithNoLock & " WHERE " & PrimaryKeyName & "=" & PrimaryKeyValue
    '            If db.sqlGetLng(s) > db.DBPacketSize4Varbinary Then
    '                VarBinaryFieldsToUpdate.Add(TableInfo.Keys(i).ToUpper, "")
    '                Continue For
    '            End If

    '        End If

    '        tex.Cat(stdcol, TableInfo.Keys(i), ", ")
    '    Next

    '    'statt * nur die Nicht-Varbinary-Felder laden, sonst ggf. sehr umfangreicher Datareader
    '    s = "SELECT " & stdcol & " FROM " & TableName & db.WithNoLock & " WHERE " & PrimaryKeyName & "=" & PrimaryKeyValue

    '    Using dr As New clsDataReader
    '        If Not dr.OpenReadonly(db, s) Then Return False
    '        If Not dr.Read() Then Return False

    '        For i = 0 To dr.FieldCount - 1
    '            If dr.getFieldName(i) <> PrimaryKeyName Then
    '                SetFieldValue(dr.getFieldName(i), dr.getObject(i))
    '            End If
    '        Next i
    '    End Using

    '    For Each key As String In VarBinaryFieldsToUpdate.Keys.ToList
    '        Dim tmpDatei As String = GetPrismaTempFilenameLocal(table_Name, key, primaryKey_Value)
    '        If Not data_base.FileDownload(table_Name, primaryKey_Name, primaryKey_Value, key, tmpDatei) Then Datei.Delete(tmpDatei) : Return False
    '        VarBinaryFieldsToUpdate(key) = tmpDatei
    '    Next

    '    Return True
    'End Function

    'Public Function OpenEditSpecial4PrepopulationByTransactionsOnly(ByVal data_base As clsDatabase, ByVal table_Name As String, ByVal primaryKey_Name As String, ByVal primaryKey_Values As List(Of Integer), ByVal limitColumns As List(Of String)) As Boolean
    '    PrimaryKeys_Fields.Clear()
    '    PrimaryKeys_VarBinaryFieldsToUpdate.Clear()

    '    For Each primaryKey_Value As Integer In primaryKey_Values
    '        If Not PrimaryKeys_Fields.ContainsKey(primaryKey_Value) Then PrimaryKeys_Fields.Add(primaryKey_Value, New ArrayList)
    '        If Not PrimaryKeys_VarBinaryFieldsToUpdate.ContainsKey(primaryKey_Value) Then PrimaryKeys_VarBinaryFieldsToUpdate.Add(primaryKey_Value, New Dictionary(Of String, String))
    '    Next

    '    db = data_base
    '    TableName = table_Name
    '    PrimaryKeyName = primaryKey_Name
    '    TableInfo = data_base.GetTableInformation(TableName)

    '    Mode = EnumRecordsetMode.rmEdit
    '    Init()

    '    'aktuelle Werte laden
    '    Dim i As Integer
    '    Dim s As String = ""
    '    Dim stdcol As String = PrimaryKeyName

    '    For i = 0 To TableInfo.Count - 1
    '        If TableInfo.Keys(i) = primaryKey_Name Then Continue For

    '        If TableInfo.Values(i).ToUpper.Contains("VARBINARY") Then
    '            s = "SELECT " & PrimaryKeyName & ", DATALENGTH(" & TableInfo.Keys(i) & ") AS Size FROM " & TableName & db.WithNoLock & " WHERE " & PrimaryKeyName & " IN (" & primaryKey_Values.ToStringOfCommaSeparatedValues() & ")"

    '            Using dr As New clsDataReader() : dr.OpenReadonly(db, s)
    '                While dr.Read()
    '                    Dim PrimaryKeyValue As Integer = dr.getInt(PrimaryKeyName)

    '                    If dr.getInt("Size") > db.DBPacketSize4Varbinary Then
    '                        PrimaryKeys_VarBinaryFieldsToUpdate(PrimaryKeyValue).Add(TableInfo.Keys(i).ToUpper, "")
    '                    End If
    '                End While
    '            End Using
    '        End If

    '        If (limitColumns Is Nothing) OrElse (limitColumns.Contains(TableInfo.Keys(i))) Then tex.Cat(stdcol, TableInfo.Keys(i), ", ")
    '    Next

    '    'statt * nur die Nicht-Varbinary-Felder laden, sonst ggf. sehr umfangreicher Datareader
    '    s = "SELECT " & stdcol & " FROM " & TableName & db.WithNoLock & " WHERE " & PrimaryKeyName & " IN (" & PrimaryKeys_Fields.Keys.ToList.ToStringOfCommaSeparatedValues() & ")"

    '    Using dr As New clsDataReader
    '        If Not dr.OpenReadonly(db, s) Then Return False

    '        While dr.Read()
    '            For i = 0 To dr.FieldCount - 1
    '                SetFieldValue4PrepopulationByTransactionsOnly(PrimaryKeyName, dr.getInt(PrimaryKeyName), dr.getFieldName(i), dr.getObject(i))
    '            Next i
    '        End While
    '    End Using

    '    For Each PrimaryKey As Integer In PrimaryKeys_VarBinaryFieldsToUpdate.Keys
    '        For Each key As String In PrimaryKeys_VarBinaryFieldsToUpdate(PrimaryKey).Keys.ToList
    '            Dim tmpDatei As String = GetPrismaTempFilenameLocal(TableName, key, PrimaryKey)
    '            If Not data_base.FileDownload(TableName, PrimaryKeyName, PrimaryKey, key, tmpDatei) Then Datei.Delete(tmpDatei) : Return False
    '            PrimaryKeys_VarBinaryFieldsToUpdate(PrimaryKey)(key) = tmpDatei
    '        Next
    '    Next

    '    Return True
    'End Function

    Public Sub PrepareForDuplicate()
        PrimaryKeys_Fields.Clear()
        PrimaryKeys_VarBinaryFieldsToUpdate.Clear()

        PrimaryKeyValue = 0
        PrimaryKeyValueVar = ""
        Mode = EnumRecordsetMode.rmAddNew
    End Sub

    Public Sub PrepareForPrePopulate(ByVal data_base As clsDatabase, ByVal primaryKey_Value As Integer)
        PrimaryKeys_Fields.Clear()
        PrimaryKeys_VarBinaryFieldsToUpdate.Clear()

        db = data_base
        PrimaryKeyValue = primaryKey_Value
        Mode = CType(IIf(PrimaryKeyValue > 0, EnumRecordsetMode.rmEdit, EnumRecordsetMode.rmAddNew), EnumRecordsetMode)

        Dim PrePFieldNames As String() = {"LOCKSTANDORT", "LOCKBENUTZERID", "LOCKCOMPUTER", "LOCKSESSIONID", "LOCKDATUMZEIT", "PREPOPULATION", "PREPOPULATIONRUNNING", "LOESCHKZ"}

        For i As Integer = 0 To Fields.Count - 1
            If PrePFieldNames.Contains(CType(Fields(i), clsRecordsetfield).FieldName.ToUpper) Then
                Fields.RemoveAt(i)
                i -= 1
            End If

            If i = Fields.Count - 1 Then Exit For
        Next
    End Sub

    Public Sub PrepareForPrePopulate4PrepopulationByTransactionsOnly()
        Dim PrePFieldNames As String() = {"LOCKSTANDORT", "LOCKBENUTZERID", "LOCKCOMPUTER", "LOCKSESSIONID", "LOCKDATUMZEIT", "PREPOPULATION", "PREPOPULATIONRUNNING", "LOESCHKZ"}
        Mode = EnumRecordsetMode.rmEdit

        For Each Fields As ArrayList In PrimaryKeys_Fields.Values
            For i As Integer = 0 To Fields.Count - 1
                If PrePFieldNames.Contains(CType(Fields(i), clsRecordsetfield).FieldName.ToUpper) Then
                    Fields.RemoveAt(i)
                    i -= 1
                End If

                If i = Fields.Count - 1 Then Exit For
            Next
        Next
    End Sub

    Public Sub SetRecord4MultiInsert(recordFields As ArrayList)
        PrimaryKeys_Fields.Add(PrimaryKeys_Fields.Count, recordFields)
    End Sub

    Public Sub SetRecord4MultiUpdate(PK As Integer, recordFields As ArrayList)
        PrimaryKeys_Fields.Add(PK, recordFields)
    End Sub

    Public Sub SetFieldValue(ByVal field_Name As String, ByVal field_Value As Object, Optional ByVal direct As Boolean = False, Optional ByVal AppendVarBinary As Boolean = False)
        PrimaryKeys_Fields.Clear()
        PrimaryKeys_VarBinaryFieldsToUpdate.Clear()

        Dim f As clsRecordsetfield

        '--- Liste durchsuchen
        For Each f In Fields
            If UCase$(f.FieldName) = UCase$(field_Name) Then
                f.FieldValue = field_Value
                f.Direct = direct
                f.FieldAppendVarBinary = AppendVarBinary
                If VarBinaryFieldsToUpdate.ContainsKey(UCase$(field_Name)) And IsDBNull(field_Value) Then VarBinaryFieldsToUpdate.Remove(UCase$(field_Name))
                Exit Sub
            End If
        Next
        '--- nicht gefunden, daher neu eintragen
        f = New clsRecordsetfield
        f.FieldName = field_Name
        f.FieldValue = field_Value
        f.Direct = direct
        f.FieldAppendVarBinary = AppendVarBinary
        Fields.Add(f)
        If VarBinaryFieldsToUpdate.ContainsKey(UCase$(field_Name)) And IsDBNull(field_Value) Then VarBinaryFieldsToUpdate.Remove(UCase$(field_Name))
    End Sub

    Public Sub SetFieldValue4PrepopulationByTransactionsOnly(ByVal primaryKey_Name As String, ByVal primaryKey_Value As Integer, ByVal field_Name As String, ByVal field_Value As Object, Optional ByVal direct As Boolean = False, Optional ByVal AppendVarBinary As Boolean = False)
        If UCase$(field_Name) = UCase$(primaryKey_Name) Then Exit Sub

        Dim recordsetField As clsRecordsetfield
        '--- Liste durchsuchen
        For Each recordsetField In PrimaryKeys_Fields(primaryKey_Value)
            If UCase$(recordsetField.FieldName) = UCase$(field_Name) Then
                recordsetField.FieldValue = field_Value
                recordsetField.Direct = direct
                recordsetField.FieldAppendVarBinary = AppendVarBinary
                If PrimaryKeys_VarBinaryFieldsToUpdate.ContainsKey(primaryKey_Value) AndAlso PrimaryKeys_VarBinaryFieldsToUpdate(primaryKey_Value).ContainsKey(UCase$(field_Name)) And IsDBNull(field_Value) Then PrimaryKeys_VarBinaryFieldsToUpdate(primaryKey_Value).Remove(UCase$(field_Name))
                Exit Sub
            End If
        Next

        '--- nicht gefunden, daher neu eintragen
        recordsetField = New clsRecordsetfield
        recordsetField.FieldName = field_Name
        recordsetField.FieldValue = field_Value
        recordsetField.Direct = direct
        recordsetField.FieldAppendVarBinary = AppendVarBinary
        PrimaryKeys_Fields(primaryKey_Value).Add(recordsetField)
        If PrimaryKeys_VarBinaryFieldsToUpdate.ContainsKey(primaryKey_Value) AndAlso PrimaryKeys_VarBinaryFieldsToUpdate(primaryKey_Value).ContainsKey(UCase$(field_Name)) And IsDBNull(field_Value) Then PrimaryKeys_VarBinaryFieldsToUpdate(primaryKey_Value).Remove(UCase$(field_Name))
    End Sub

    Public Function GetFieldValue(ByVal field_Name As String) As Object
        Dim f As clsRecordsetfield
        '--- Liste durchsuchen

        For Each f In Fields
            If UCase$(f.FieldName) = UCase$(field_Name) Then
                Return f.FieldValue
            End If
        Next
        '--- falls nicht gefunden, prüfen, ob PrimaryKey gesucht wurde
        If UCase$(field_Name) = UCase$(Me.PrimaryKeyName) Then
            Return Me.PrimaryKeyValue
        End If
        '--- nicht gefunden
        Return Nothing
    End Function

    'Public Function getLng(ByVal field_Name As String) As Integer
    '    Dim o As Object = GetFieldValue(field_Name)
    '    If IsDBNull(o) Then Return 0
    '    Return CType(o, Integer)
    'End Function

    'Public Function getDbl(ByVal field_Name As String) As Double
    '    Dim o As Object = GetFieldValue(field_Name)
    '    If IsDBNull(o) Then Return 0.0
    '    Return CType(o, Double)
    'End Function

    'Public Function getInt(ByVal field_Name As String) As Integer
    '    Dim o As Object = GetFieldValue(field_Name)
    '    If IsDBNull(o) Then Return 0
    '    Return CType(o, Integer)
    'End Function

    'Public Function getStr(ByVal field_Name As String) As String
    '    Dim o As Object = GetFieldValue(field_Name)
    '    If IsDBNull(o) Then Return ""
    '    Return CType(o, String)
    'End Function

    'Public Function getDat(ByVal field_Name As String) As Date
    '    Dim o As Object = GetFieldValue(field_Name)
    '    If IsDBNull(o) Then Return dat.NullDate()
    '    Return CType(o, Date)
    'End Function

    'Public Function fieldNull(ByVal field_Name As String) As Boolean
    '    Dim o As Object = GetFieldValue(field_Name)
    '    If IsDBNull(o) Then Return True
    '    Return False
    'End Function

    Public Function Update() As Boolean
        Dim sqls As New List(Of String)
        Dim ret As Boolean = PrivUpdate(False, sqls)
        If ret Then ret = UpdateVarBinaryFields()
        Return ret
    End Function

    Public Function UpdateSQLs() As List(Of String)
        Dim sqls As New List(Of String)
        If Not PrivUpdate(True, sqls) Then sqls.Clear()
        Return sqls
    End Function

    Private Function UpdateVarBinaryFields() As Boolean
        If VarBinaryFieldsToUpdate.Count = 0 Then Return True

        For Each kvp As KeyValuePair(Of String, String) In VarBinaryFieldsToUpdate
            If Not db.FileUpload(TableName, PrimaryKeyName, PrimaryKeyValue, kvp.Key, kvp.Value) Then Return False
        Next

        Return True
    End Function

    'Private Function PrivUpdate(ByVal onlySqls As Boolean, ByRef retSqls As String) As Boolean
    '    If Fields.Count = 0 Then Fields = Nothing : Return False

    '    Dim s As String
    '    Dim attNamen As String
    '    Dim attWerte As String
    '    Dim attNamWer As String
    '    Dim f As clsRecordsetfield

    '    Dim aktFeld As Integer
    '    Dim anzFelder As Integer = Fields.Count

    '    Dim SqlParList As New List(Of clsSqlParameter)
    '    Select Case Mode
    '        Case EnumRecordsetMode.rmAddNew
    '            Dim fieldvarBin As Boolean = False
    '            attNamen = ""
    '            attWerte = ""
    '            For aktFeld = 0 To anzFelder - 1
    '                f = CType(Fields(aktFeld), clsRecordsetfield)
    '                'Debug.Print f.fieldname

    '                If f.FieldValue IsNot Nothing AndAlso f.FieldValue.GetType Is Type.GetType("System.Byte[]") Then
    '                    'fieldvarBin = True

    '                    Dim p As New clsSqlParameter
    '                    p.ParName = "@" & f.FieldName
    '                    p.ParValue = CType(f.FieldValue, Byte())
    '                    tex.Cat(attNamen, "[" & f.FieldName & "]", ", ")
    '                    tex.Cat(attWerte, p.ParName, ", ")
    '                    SqlParList.Add(p)
    '                    p = Nothing
    '                Else
    '                    tex.Cat(attNamen, f.FieldName, ", ")
    '                    If f.Direct Then
    '                        tex.Cat(attWerte, f.FieldValue.ToString(), ", ")
    '                    Else
    '                        tex.Cat(attWerte, db.sqlValue(f.FieldValue), ", ")
    '                    End If
    '                End If

    '            Next aktFeld
    '            s = "INSERT INTO " & TableName & db.WithRowLock & " (" & attNamen & ") VALUES (" & attWerte & ")"
    '            If onlySqls Then
    '                tex.Cat(retSqls, s, vbCrLf)
    '            Else
    '                If db.IdentitySupported Then
    '                    If TypeOf db Is clsDatabaseSQLServer Then
    '                        s += vbCrLf & "SELECT @@IDENTITY"
    '                        Using dr As New clsDataReader
    '                            If Not dr.OpenReadonly(db, s) Then Fields = Nothing : Return False
    '                            'dr.NextResult()
    '                            If Not dr.Read() Then Fields = Nothing : Return False
    '                            PrimaryKeyValue = dr.getLng(0)
    '                        End Using
    '                    Else
    '                        If db.sqlExecute(s, SqlParList) = -1 Then Fields = Nothing : Return False
    '                        s = "SELECT @@IDENTITY"
    '                        PrimaryKeyValue = db.sqlGetLng(s)
    '                    End If
    '                Else 'Notlösung, nicht 100%ig
    '                    If db.sqlExecute(s, SqlParList) = -1 Then Fields = Nothing : Return False
    '                    s = "SELECT MAX(" & PrimaryKeyName & ") FROM " & TableName & db.WithNoLock
    '                    PrimaryKeyValue = db.sqlGetLng(s)
    '                End If
    '                If fieldvarBin Then Mode = EnumRecordsetMode.rmEdit : PrivUpdate(onlySqls, retSqls) 'Byteblöcke nur über Update einfügen, ansonsten keine Blobübertragung möglich (PrimaryKeyValue erst nach INSERT-Aufruf bekannt)
    '            End If
    '        Case EnumRecordsetMode.rmEdit
    '            Dim aktBlock As Integer = 0
    '            Dim anzBloecke As Integer = ((anzFelder - 1) \ maxAnzUpd) + 1

    '            For aktBlock = 0 To anzBloecke - 1
    '                attNamWer = ""
    '                For aktFeld = aktBlock * maxAnzUpd To Math.Min((aktBlock + 1) * maxAnzUpd - 1, anzFelder - 1)
    '                    f = CType(Fields(aktFeld), clsRecordsetfield)
    '                    If f.FieldValue IsNot Nothing AndAlso f.FieldValue.GetType Is Type.GetType("System.Byte[]") Then
    '                        'UpdateVarBinary(f, SqlParList, attNamWer)
    '                        Dim p As New clsSqlParameter
    '                        p.ParName = "@" & f.FieldName
    '                        p.ParValue = CType(f.FieldValue, Byte())
    '                        tex.Cat(attNamWer, "[" & f.FieldName & "]" & "=" & p.ParName, ", ")
    '                        SqlParList.Add(p)
    '                        p = Nothing
    '                    Else
    '                        If f.Direct Then
    '                            tex.Cat(attNamWer, f.FieldName & "=" & f.FieldValue.ToString(), ", ")
    '                        Else
    '                            tex.Cat(attNamWer, f.FieldName & "=" & db.sqlValue(f.FieldValue), ", ")
    '                        End If
    '                    End If
    '                Next aktFeld
    '                s = "UPDATE " & TableName & db.WithRowLock & " SET " & attNamWer & " WHERE " & PrimaryKeyName & "=" & IIf(PrimaryKeyValueVar = "", PrimaryKeyValue, PrimaryKeyValueVar).ToString()
    '                If onlySqls Then
    '                    tex.Cat(retSqls, s, vbCrLf)
    '                Else
    '                    'If attNamWer <> "" AndAlso db.sqlExecute(s, SqlParList) = -1 Then Fields = Nothing : Return False
    '                    If db.sqlExecute(s, SqlParList) = -1 Then Fields = Nothing : Return False
    '                End If
    '            Next aktBlock
    '    End Select
    '    'Fields = Nothing
    '    Return True
    'End Function

    Private Function PrivUpdate(ByVal onlySqls As Boolean, retSqls As List(Of String)) As Boolean
        If PrimaryKeys_Fields.Count = 0 Then
            PrimaryKeys_Fields.Add(PrimaryKeyValue, Fields)
        End If

        If PrimaryKeys_Fields.Count = 0 Then Return False

        For Each PrimaryKey As Integer In PrimaryKeys_Fields.Keys
            Dim s As String
            Dim attNamen As New List(Of String)
            Dim attWerte As New List(Of String)
            Dim attNamWer As String
            Dim f As clsRecordsetfield

            Dim aktFeld As Integer
            Dim anzFelder As Integer = PrimaryKeys_Fields(PrimaryKey).Count

            Dim SqlParList As New List(Of clsSqlParameter)
            Select Case Mode
                Case EnumRecordsetMode.rmAddNew
                    Dim fieldvarBin As Boolean = False

                    For aktFeld = 0 To anzFelder - 1
                        f = CType(PrimaryKeys_Fields(PrimaryKey)(aktFeld), clsRecordsetfield)
                        'Debug.Print f.fieldname

                        If f.FieldValue IsNot Nothing AndAlso f.FieldValue.GetType Is Type.GetType("System.Byte[]") Then
                            fieldvarBin = True

                            Dim p As New clsSqlParameter
                            p.ParName = "@" & f.FieldName
                            p.ParValue = CType(f.FieldValue, Byte())
                            attNamen.Add("[" & f.FieldName & "]")
                            attWerte.Add(p.ParName)
                            SqlParList.Add(p)
                            p = Nothing
                        Else
                            attNamen.Add(f.FieldName)

                            If f.Direct Then
                                attWerte.Add(If(f.FieldValue Is Nothing, "NULL", f.FieldValue.ToString()))
                            Else
                                attWerte.Add(db.sqlValue(f.FieldValue))
                            End If
                        End If
                    Next aktFeld

                    s = "INSERT INTO " & TableName & db.WithRowLock & " (" & String.Join(",", attNamen.ToArray) & ") VALUES (" & String.Join(",", attWerte.ToArray) & ")"

                    If onlySqls Then
                        retSqls.Add(s)
                    Else
                        If db.IdentitySupported Then
                            If TypeOf db Is clsDatabaseSQLServer AndAlso SqlParList.Count = 0 Then
                                s += vbCrLf & "SELECT @@IDENTITY"

                                Using dr As New clsDataReader
                                    If Not dr.OpenReadonly(db, s) Then PrimaryKeys_Fields.Clear() : Fields = Nothing : Return False
                                    'dr.NextResult()
                                    If Not dr.Read() Then PrimaryKeys_Fields.Clear() : Fields = Nothing : Return False

                                    PrimaryKeyValue = dr.getLng(0)
                                End Using
                            Else
                                If db.sqlExecute(s, SqlParList) = -1 Then PrimaryKeys_Fields.Clear() : Fields = Nothing : Return False

                                s = "SELECT @@IDENTITY"
                                PrimaryKeyValue = db.sqlGetLng(s)
                            End If
                        Else 'Notlösung, nicht 100%ig
                            If db.sqlExecute(s, SqlParList) = -1 Then PrimaryKeys_Fields.Clear() : Fields = Nothing : Return False

                            s = "SELECT MAX(" & PrimaryKeyName & ") FROM " & TableName & db.WithNoLock
                            PrimaryKeyValue = db.sqlGetLng(s)
                        End If

                        If fieldvarBin Then Mode = EnumRecordsetMode.rmEdit : PrivUpdate(onlySqls, retSqls) 'Byteblöcke nur über Update einfügen, ansonsten keine Blobübertragung möglich (PrimaryKeyValue erst nach INSERT-Aufruf bekannt)
                    End If

                Case EnumRecordsetMode.rmEdit
                    Dim aktBlock As Integer = 0
                    Dim anzBloecke As Integer = ((anzFelder - 1) \ maxAnzUpd) + 1

                    For aktBlock = 0 To anzBloecke - 1
                        attNamWer = ""

                        For aktFeld = aktBlock * maxAnzUpd To Math.Min((aktBlock + 1) * maxAnzUpd - 1, anzFelder - 1)
                            f = CType(PrimaryKeys_Fields(PrimaryKey)(aktFeld), clsRecordsetfield)

                            If f.FieldValue IsNot Nothing AndAlso f.FieldValue.GetType Is Type.GetType("System.Byte[]") Then
                                UpdateVarBinary(f, SqlParList, attNamWer)
                            Else
                                If f.Direct Then
                                    tex.Cat(attNamWer, f.FieldName & "=" & If(f.FieldValue Is Nothing, "NULL", f.FieldValue.ToString()), ", ")
                                Else
                                    tex.Cat(attNamWer, f.FieldName & "=" & db.sqlValue(f.FieldValue), ", ")
                                End If
                            End If
                        Next aktFeld

                        s = "UPDATE " & TableName & db.WithRowLock & " SET " & attNamWer & " WHERE " & PrimaryKeyName & "=" & PrimaryKey

                        If onlySqls Then
                            retSqls.Add(s)
                        Else
                            If attNamWer <> "" AndAlso db.sqlExecute(s, SqlParList) = -1 Then PrimaryKeys_Fields.Clear() : Fields = Nothing : Return False
                            'If db.sqlExecute(s, SqlParList) = -1 Then PrimaryKeys_Fields.Clear() : Fields = Nothing : Return False
                        End If
                    Next aktBlock

            End Select
        Next

        'Fields = Nothing
        Return True
    End Function

    Private Sub UpdateVarBinary(ByVal field As clsRecordsetfield, ByRef parList As List(Of clsSqlParameter), ByRef attNamWer As String)
        Dim s As String = ""
        Dim index As Integer = 0
        Dim p As New clsSqlParameter

        p.ParName = "@" & field.FieldName
        p.ParValue = CType(field.FieldValue, Byte())

        If Not field.FieldAppendVarBinary And p.ParValue.Length <= db.DBPacketSize4Varbinary Then
            'NULL oder Byteblock kleiner/gleich Maximalgröße
            If p.ParValue.Length = 0 Then
                tex.Cat(attNamWer, "[" & field.FieldName & "]" & "=" & db.sqlValue(DBNull.Value), ", ")
            Else
                tex.Cat(attNamWer, "[" & field.FieldName & "]" & "=" & p.ParName, ", ")
                parList.Add(p)
            End If
        Else
            'Byteblock größer Maximalgröße (komplette Datei in Speicher/Variable geladen) -> zerteilen
            'oder Inhalt anfügen (Beispiel siehe clsDatabase.FileUpload)
            index = 0
            Dim pp As New clsSqlParameter
            pp.ParName = p.ParName
            pp.ParValue = p.ParValue.Take(db.DBPacketSize4Varbinary).ToArray

            Dim pparl As New List(Of clsSqlParameter)
            pparl.Add(pp)

            s = "UPDATE " & TableName & db.WithRowLock & " SET [" & field.FieldName & "]" & "=" & pp.ParName & " WHERE " & PrimaryKeyName & "=" & PrimaryKeyValue
            If Not field.FieldAppendVarBinary Then db.sqlExecute(s, pparl) : index = db.DBPacketSize4Varbinary 'ersten Byteblock speichern

            pp = Nothing
            pparl = Nothing

            Dim arr As Byte()
            While (p.ParValue.Length - index) > db.DBPacketSize4Varbinary 'Byteblocks anfügen
                arr = p.ParValue.Skip(index).Take(db.DBPacketSize4Varbinary).ToArray
                Dim builder As New System.Text.StringBuilder("", 2 + db.DBPacketSize4Varbinary)
                builder.Append("0x" & BitConverter.ToString(arr, 0, db.DBPacketSize4Varbinary).Replace("-", String.Empty))
                s = "UPDATE " & TableName & db.WithRowLock & " SET [" & field.FieldName & "].WRITE(" & builder.ToString & ", NULL, 0) WHERE " & PrimaryKeyName & "=" & PrimaryKeyValue
                db.sqlExecute(s)
                builder = Nothing
                Array.Clear(arr, 0, db.DBPacketSize4Varbinary)
                index += db.DBPacketSize4Varbinary
                Application.DoEvents()
            End While

            arr = p.ParValue.Skip(index).Take(db.DBPacketSize4Varbinary).ToArray 'restliche Bytes anfügen
            If arr.Length > 0 Then
                Dim ppst As String = "0x" & BitConverter.ToString(arr, 0, p.ParValue.Length - index).Replace("-", String.Empty)
                s = "UPDATE " & TableName & db.WithRowLock & " SET [" & field.FieldName & "].WRITE(" & ppst & ", NULL, 0) WHERE " & PrimaryKeyName & "=" & PrimaryKeyValue
                db.sqlExecute(s)
            End If
            Array.Clear(arr, 0, arr.Length)
        End If

        p = Nothing
        GC.WaitForPendingFinalizers()
    End Sub

    Public Function LookForSameExcept(ByVal except As String) As Integer
        Dim s As String
        Dim h As String
        Dim i As Integer

        s = "SELECT " & PrimaryKeyName
        s += " FROM " & TableName & db.WithNoLock
        s += " WHERE " & PrimaryKeyName & "<>" & PrimaryKeyValue
        For i = 1 To Fields.Count
            Dim f As clsRecordsetfield
            f = CType(Fields.Item(i), clsRecordsetfield)
            If InStr("," & UCase$(except) & ",", "," & UCase$(f.FieldName) & ",") = 0 Then
                h = db.SqlCondEqual(f)
                tex.Cat(s, h, " AND")
            End If
        Next i

        LookForSameExcept = db.sqlGetLng(s)
    End Function

    Private Sub Close()
        For Each tmpDatei As String In VarBinaryFieldsToUpdate.Values
            If datei.Exists(tmpDatei) Then datei.Delete(tmpDatei)
        Next

        If Fields IsNot Nothing Then Fields.Clear()
        Fields = Nothing
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' So ermitteln Sie überflüssige Aufrufe

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: Verwalteten Zustand löschen (verwaltete Objekte).
            End If

            ' TODO: Nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalize() unten überschreiben.
            ' TODO: Große Felder auf NULL festlegen.

            If VarBinaryFieldsToUpdate IsNot Nothing Then Close() : VarBinaryFieldsToUpdate.Clear() : VarBinaryFieldsToUpdate = Nothing
            If Fields IsNot Nothing Then Fields.Clear() : Fields = Nothing
            If TableInfo IsNot Nothing Then TableInfo.Clear() : TableInfo = Nothing

            If PrimaryKeys_Fields IsNot Nothing Then PrimaryKeys_Fields.Clear() : PrimaryKeys_Fields = Nothing
            If PrimaryKeys_VarBinaryFieldsToUpdate IsNot Nothing Then PrimaryKeys_VarBinaryFieldsToUpdate.Clear() : PrimaryKeys_VarBinaryFieldsToUpdate = Nothing

            db = Nothing
        End If

        Me.disposedValue = True
    End Sub

    ' TODO: Finalize() nur überschreiben, wenn Dispose(ByVal disposing As Boolean) oben über Code zum Freigeben von nicht verwalteten Ressourcen verfügt.
    Protected Overrides Sub Finalize()
        ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

'==================================================================================================
Public Class clsRecordsetfield
    Public FieldName As String
    Public FieldValue As Object
    Public Direct As Boolean
    Public FieldAppendVarBinary As Boolean = False
End Class

'==================================================================================================
Public Class clsSqlParameter
    Public ParName As String
    Public ParValue As Byte()
End Class

Public Class clsSqlTransaction
    Implements IDisposable

    Public db As clsDatabase

    Private _errorMessage As String = ""
    Private _isCommited As Boolean = False

    Public trySqlStatements As New List(Of String)
    Public catchSqlStatements As New List(Of String)

    Public ReadOnly Property ErrorMessage As String
        Get
            Return _errorMessage
        End Get
    End Property

    Public ReadOnly Property IsCommited As Boolean
        Get
            Return _isCommited
        End Get
    End Property

    Public ReadOnly Property Transaction As String
        Get
            Return clsSqlBuilder.GetTransaction(trySqlStatements, catchSqlStatements)
        End Get
    End Property

    Public Sub New(ByVal db As clsDatabase)
        Me.db = db
    End Sub

    Public Sub New(ByVal db As clsDatabase, ByVal trySqlStatements As List(Of String), Optional ByVal catchSqlStatements As List(Of String) = Nothing)
        Me.db = db
        Me.trySqlStatements = trySqlStatements
        If catchSqlStatements IsNot Nothing Then Me.catchSqlStatements = catchSqlStatements
        Me.Commit()
    End Sub

    Public Function Commit() As Boolean
        _errorMessage = ""
        _isCommited = False

        If db Is Nothing Then Return False
        If trySqlStatements Is Nothing Then Return False
        If catchSqlStatements Is Nothing Then Return False
        If trySqlStatements.Count = 0 Then _isCommited = True : Return IsCommited

        Try
            _errorMessage = db.sqlGetStr(Transaction)

        Catch ex As System.OutOfMemoryException
            clsShow.ErrorMsg(My.Resources.resMain.MsgOutOfMemorySplittingTransactionOfSQLStatementsInto2Parts.Replace("{0}", trySqlStatements.Count))

            Dim sqlStatements As List(Of String) = Me.trySqlStatements
            Dim half As Integer = sqlStatements.Count \ 2

            Me.trySqlStatements = sqlStatements.Take(half).ToList
            If Commit() Then
                Me.trySqlStatements = sqlStatements.Skip(half).ToList
                Commit()
            End If
        End Try

        ' If _errorMessage = "" AndAlso db.LastError <> "" Then _errorMessage = db.LastError 'e.g. timeout
        If _errorMessage <> "" Then Return False

        _isCommited = True

        Return IsCommited
    End Function

    Public Shared Function Commit(ByVal db As clsDatabase, ByVal trySqlStatements As List(Of String), Optional ByVal catchSqlStatements As List(Of String) = Nothing) As Boolean
        Using transaction As New clsSqlTransaction(db, trySqlStatements, catchSqlStatements)
            Return transaction.IsCommited
        End Using
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' So ermitteln Sie überflüssige Aufrufe

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: Verwalteten Zustand löschen (verwaltete Objekte).
            End If

            ' TODO: Nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalize() unten überschreiben.
            ' TODO: Große Felder auf NULL festlegen.

            db = Nothing
            trySqlStatements = Nothing
            catchSqlStatements = Nothing
        End If

        Me.disposedValue = True
    End Sub

    ' TODO: Finalize() nur überschreiben, wenn Dispose(ByVal disposing As Boolean) oben über Code zum Freigeben von nicht verwalteten Ressourcen verfügt.
    Protected Overrides Sub Finalize()
        ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' Dieser Code wird von Visual Basic hinzugefügt, um das Dispose-Muster richtig zu implementieren.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class

Public Class clsSqlBuilder

    Public Const sqlCrLf As String = "\r\n"

    Private Shared Function GetDeclareErrorVars() As String
        Dim s As String = ""

        s += "DECLARE @error_message VARCHAR(MAX)" & vbCrLf
        s += "DECLARE @error_severity INT" & vbCrLf
        s += "DECLARE @error_state INT" & vbCrLf

        Return s
    End Function

    Public Shared Function GetIdentityInsertIfNotExists(ByVal tableName As String, ByVal primaryKeyName As String, ByVal primaryKeyValues As List(Of Integer), Optional ByVal catchSqlStatements As List(Of String) = Nothing) As String
        If catchSqlStatements Is Nothing Then catchSqlStatements = New List(Of String)

        Dim catchSql As New System.Text.StringBuilder(catchSqlStatements.ToStringOfCommaSeparatedValues(vbCrLf))
        catchSql.Replace(vbCrLf, "    " & vbCrLf)
        catchSql.Replace(sqlCrLf, vbCrLf)

        ' Stringbuilder mit der Kapazität von catchSql und 512 Zeichen für das TRY SET IDENTITY_INSERT CATCH Konstrukt und jeweils 512 Zeichen für ein IF Konstrukt
        Dim s As New System.Text.StringBuilder(catchSql.Capacity + 512 + (512 * primaryKeyValues.Count()))

        s.AppendLine("BEGIN TRY")
        s.AppendLine("    SET IDENTITY_INSERT " & tableName & " ON")

        For Each primaryKeyValue As Integer In primaryKeyValues
            s.AppendLine()
            s.AppendLine("    IF NOT EXISTS (SELECT 1 FROM " & tableName & session.db.WithNoLock & " WHERE " & primaryKeyName & " = " & primaryKeyValue & ")")
            s.AppendLine("    BEGIN")
            s.AppendLine("        INSERT INTO " & tableName & session.db.WithRowLock & " (" & primaryKeyName & ")")
            s.AppendLine("        VALUES (" & primaryKeyValue & ")")
            s.AppendLine("    END")
        Next

        s.AppendLine()
        s.AppendLine("    SET IDENTITY_INSERT " & tableName & " OFF")
        s.AppendLine("END TRY")
        s.AppendLine("BEGIN CATCH")
        s.AppendLine("    SET IDENTITY_INSERT " & tableName & " OFF")
        If catchSql.ToString() <> "" Then s.AppendLine()
        If catchSql.ToString() <> "" Then s.Append(catchSql) : catchSql = Nothing : GC.Collect() : s.AppendLine()
        s.AppendLine()
        s.AppendLine("    " & GetRaisError().Replace(vbCrLf, vbCrLf & "    "))
        s.AppendLine("END CATCH")

        Return s.ToString()
    End Function

    Private Shared Function GetRaisError() As String
        Dim s As String = ""

        s += "SET @error_message = ERROR_MESSAGE()" & vbCrLf
        s += "SET @error_severity = ERROR_SEVERITY()" & vbCrLf
        s += "SET @error_state = ERROR_STATE()" & vbCrLf
        s += vbCrLf
        s += "RAISERROR (@error_message, @error_severity, @error_state)" & vbCrLf

        Return s
    End Function

    Public Shared Function GetTransaction(ByVal trySqlStatements As List(Of String), Optional ByVal catchSqlStatements As List(Of String) = Nothing) As String
        If catchSqlStatements Is Nothing Then catchSqlStatements = New List(Of String)

        Dim trySql As New System.Text.StringBuilder(trySqlStatements.ToStringOfCommaSeparatedValues(vbCrLf))
        trySql.Replace(vbCrLf, vbCrLf & "        ")
        trySql.Replace(sqlCrLf, vbCrLf)

        Dim catchSql As New System.Text.StringBuilder(catchSqlStatements.ToStringOfCommaSeparatedValues(vbCrLf))
        catchSql.Replace(vbCrLf, vbCrLf & "        ")
        catchSql.Replace(sqlCrLf, vbCrLf)

        ' Stringbuilder mit der Kapazität von trySql und catchSql und 1024 Zeichen für das Transaction Konstrukt
        Dim s As New System.Text.StringBuilder(trySql.Capacity + catchSql.Capacity + 1024)

        s.AppendLine(GetDeclareErrorVars())
        s.AppendLine()
        s.AppendLine("BEGIN TRANSACTION")
        s.AppendLine("    BEGIN TRY")
        s.Append("        ") : s.Append(trySql) : trySql = Nothing : GC.Collect() : s.AppendLine()
        s.AppendLine()
        s.AppendLine("        COMMIT TRANSACTION")
        s.AppendLine("    END TRY")
        s.AppendLine("    BEGIN CATCH")
        s.AppendLine("        ROLLBACK TRANSACTION")
        s.AppendLine()

        If catchSql.ToString() <> "" Then
            s.Append("         ") : s.Append(catchSql) : catchSql = Nothing : GC.Collect() : s.AppendLine()
            s.AppendLine()
        End If

        s.AppendLine("        SELECT 'Msg ' + CONVERT(VARCHAR, ERROR_NUMBER()) + ', Level ' + CONVERT(VARCHAR, ERROR_SEVERITY()) + ', State ' + CONVERT(VARCHAR, ERROR_STATE()) + ', Line ' + CONVERT(VARCHAR, ERROR_LINE()) + CHAR(13) + CHAR(10) + ERROR_MESSAGE() + ' TRANSACTION WAS ROLLED BACK!'")
        s.AppendLine("    END CATCH")

        Return s.ToString()
    End Function

    Private Shared Sub Replace(ByVal sqlStatements As List(Of String), ByVal oldString As String, ByVal newString As String)
        For index As Integer = sqlStatements.Count - 1 To 0 Step -1
            sqlStatements(index) = sqlStatements(index).Replace(oldString, newString)
        Next
    End Sub

End Class