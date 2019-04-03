Option Strict Off
Option Explicit On

Friend Class clsSession
    Public SILENT As Boolean

    '- Databases
    Public db As clsDatabase
    Public dbconnstr As String
    Private dbconnstr2 As String 'inkl. User/Password
    Public dbTitle As String
    Public dbSource As String

    Public mtDB As clsDatabase
    Private mt_dbconnstr As String
    Private mt_dbconnstr2 As String

    '- Application
    Public appDir As String
    Public appExeFile As String
    Public appFileDateTime As Date
    Public appIni As String
    Public appName As String
    Public appVersionStr As String
    Public appDBVersion As Integer
    Public appState As String
    Public appCopyright As String
    Public appCompany As String
    Public appExePath As String
    Public appVersionInt As Integer 'from Project properties
    Private AppVersionDBParameter As Integer 'from DB

    '- Dirs
    Public ProjectFolder As String
    Public DestinationFolder As String
    Public IniTemplatesDir As String
    Public IniReportsDir As String

    '- Program
    Public IniTransmar As Boolean

    Public IniOutTypeLblConn As String
    Public IniOutTypeLblConnDest As String
    Public IniOutTypeLblDev As String
    Public IniOutTypeLblTerm As String
    Public IniOutTypeLblTerminals As String
    Public IniOutTypeCalcEffort As String
    Public IniOutTypeCalcWire As String
    Public IniPricePerHour As Double

    Public Function Init() As Boolean
        InitFilter()

        appName = Application.ProductName()
        appDir = Application.StartupPath()
        If appDir.EndsWith("\Debug") Or appDir.EndsWith("\Release") Then appDir = IO.Path.GetDirectoryName(appDir)
        If appDir.EndsWith("\bin") Then appDir = IO.Path.GetDirectoryName(appDir) 'appDir.Substring(0, appDir.Length - 4)

        Dim h As String = Application.ExecutablePath()
        appExeFile = IO.Path.GetFileName(h)

        Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(appDir & "\" & appExeFile)
        appFileDateTime = fi.LastWriteTime

        appIni = appDir & "\" & IO.Path.ChangeExtension(appExeFile, "ini")

        If Not IO.File.Exists(appIni) Then
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgIniFile_IsMissing.Replace("{0}", appIni))
            Return False
        End If

        Dim ass As Reflection.Assembly
        ass = Reflection.Assembly.GetExecutingAssembly()
        With Diagnostics.FileVersionInfo.GetVersionInfo(ass.Location)
            appVersionStr = .FileMajorPart() & "." & .FileMinorPart().ToString("00") & "." & .FileBuildPart.ToString("00")
            appState = .Comments()
            appVersionInt = ((.FileMajorPart * 1000) + .FileMinorPart) * 1000 + 0
            appCopyright = .LegalCopyright
            appCompany = .CompanyName
        End With

        Return True
    End Function

    Public Function ReadIniFile() As Boolean
        Dim inifile As New clsIniFile(appIni)
        Dim dbname As String = inifile.GetString("Files", "Database")
        dbTitle = "Access: " + dbname
        dbname = ModifyFileName(dbname)

        dbconnstr = inifile.GetString("Databases", "Self")

        Dim i As Integer
        Dim h As String

        If dbconnstr = "" Then 'nicht Netz-Version
            'Access-DB
            If dbname.ToLower.IndexOf("mdb") = 0 Then
                clsShow.ErrorMsg(My.Resources.resGlobal.MsgInvalidDatabaseFileName_InIniFile.Replace("{0}", dbname))
                Return False
            Else
                If Not IO.File.Exists(dbname) Then
                    clsShow.ErrorMsg(My.Resources.resGlobal.MsgDatabaseFile_IsMissing.Replace("{0}", dbname))
                    Return False
                End If
            End If
            dbconnstr = "Provider='Microsoft.Jet.OLEDB.4.0';Data Source='" + dbname + "';"
            dbconnstr2 = dbconnstr & "User ID='admin';Password='';"
            'dbTitle = dbname
            dbSource = ""
        Else 'SQL-Server
            If dbconnstr = "" Or Right$(dbconnstr, 1) <> ";" Then
                clsShow.ErrorMsg(My.Resources.resGlobal.MsgInvalidConnectionString_.Replace("{0}", dbconnstr))
                Return False
            End If
            dbconnstr2 = ""
            For i = 1 To tex.PartCount(dbconnstr, ";")
                h = tex.Part(dbconnstr, i, ";")
                If tex.Part(h, 1, "=") <> "Provider" And h <> "" Then 'Provider gibt es hier nicht
                    tex.Cat(dbconnstr2, h, ";")
                End If

                If tex.Part(h, 1, "=") = "Initial Catalog" Then
                    dbname = Replace(tex.Part(h, 2, "="), "'", "")
                End If
            Next
            tex.Cat(dbconnstr2, "Persist Security Info=False", ";") 'zusätzlich

            If dbconnstr2.ToUpper.Contains("(LOCAL") Or dbconnstr2.ToUpper.Contains("SQLEXPRESS") Then
                tex.Cat(dbconnstr2, "Integrated Security='SSPI'", ";")
            Else
                tex.Cat(dbconnstr2, "Us" & "er" & " ID='DB" & "lo" & "gin'", ";")
                tex.Cat(dbconnstr2, "Pa" & "ss" & "word='XYZ'", ";")
            End If

            dbconnstr2 += ";"
            dbTitle = "SQL Server DB '" & dbname & "'"
            dbSource = tex.Part(tex.Part(dbconnstr, 2, ";"), 2, "=")
        End If
        db = New clsDatabaseSQLServer(dbconnstr2)

        Return True
    End Function

    Public Sub InitFilter()

    End Sub

    Public Function OpenDB() As Boolean
        If Not db.Open() Then
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileOpeningDatabaseConnectionString_.Replace("{0}", dbconnstr))
            Return False
        End If

        'Daten aus Tabelle Parameter auslesen
        If db.FieldExist("Parameter", "AppVersion") Then
            Dim s As String = "SELECT * FROM Parameter" & session.db.WithNoLock
            Using dr As New clsDataReader()
                dr.OpenReadonly(session.db, s)
                dr.Read()
                AppVersionDBParameter = dr.getLng("AppVersion")
                If db.FieldExist("Parameter", "DBVersion") Then appDBVersion = dr.getLng("DBVersion")
            End Using
        End If

        Return True
    End Function

    Public Sub CloseDB()
        db.Close()
    End Sub

    Public Function CheckVersion() As Boolean
        Dim verSoll As Integer = AppVersionDBParameter
        Dim verIst As Integer = appVersionInt

        If verIst < verSoll Then
            clsShow.InternalError(My.Resources.resGlobal.MsgProgramFileYouHaveStartedIsNotUpToDateCurrentVersionNumberIs_.Replace("{0}", ((verSoll \ 1000) \ 1000).ToString). _
                                                                                                                      Replace("{1}", (verSoll \ 1000) Mod 1000 & verSoll Mod 1000))
            Return False
        End If

        Return True
    End Function

    Public Function CheckExeDate() As Boolean
        Dim fi As System.IO.FileInfo

        Try 'vielleicht Datei nicht mehr gefunden
            fi = My.Computer.FileSystem.GetFileInfo(session.appDir & "\" & session.appExeFile) 'aktuell eingespielten EXE
        Catch ex As Exception
            clsShow.Message(My.Resources.resGlobal.MsgBecauseNewVersionOfProgramIsRecordedProgramIsTerminatedAutomaticallyNowPleaseRestartProgram)
            Return False
        End Try

        If fi.LastWriteTime > session.appFileDateTime.Add(New TimeSpan(0, 0, 1)) Then
            clsShow.Message(My.Resources.resGlobal.MsgBecauseInMeantimeNewVersionOfProgramWasRecordedProgramIsTerminatedAutomaticallyNowPleaseRestart)
            Return False
        End If

        Return True
    End Function

    Public Function PrintAppDBVersion() As String
        Return appDBVersion.ToString.Substring(0, 2) & "." & appDBVersion.ToString.Substring(2, 2) & "." & appDBVersion.ToString.Substring(4, 3)
    End Function

End Class
