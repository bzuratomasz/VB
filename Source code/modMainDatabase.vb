Module modMainDatabase
    Public Function UpdateDatabase() As Boolean
        If session.appDBVersion < 1000000 AndAlso Not UpdateDatabaseV1000000() Then Return False
        If session.appDBVersion < 1000001 AndAlso Not UpdateDatabaseV1000001() Then Return False
        If session.appDBVersion < 1000002 AndAlso Not UpdateDatabaseV1000002() Then Return False

        Dim s As String
        s = "UPDATE Parameter SET AppVersion=" & session.appVersionInt
        s += " WHERE AppVersion IS NULL OR AppVersion<" & session.appVersionInt
        session.db.sqlExecute(s)

        Return True
    End Function

    Private Function UpdateDatabaseVX0XX000() As Boolean
        Dim tmpDBVersion As Integer = 0

        Dim s As String = ""
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim dr As clsDataReader = Nothing
        Dim dw As clsDataWriter = Nothing

        With session.db

        End With

        s = "UPDATE Parameter SET DBVersion=" & tmpDBVersion
        s += " WHERE DBVersion IS NULL OR DBVersion<" & tmpDBVersion
        session.db.sqlExecute(s)

        session.appDBVersion = tmpDBVersion

        Return True
    End Function

    Private Function UpdateDatabaseV1000000() As Boolean
        Dim tmpDBVersion As Integer = 1000000
        Dim s As String = ""

        With session.db
            If Not .TableExist("Parameter") Then
                .AddField("Parameter", "AppVersion", "INT")
                .AddField("Parameter", "DBVersion", "INT")
                .sqlExecute("INSERT INTO Parameter(AppVersion, DBVersion) VALUES(NULL, NULL)")
            End If
        End With

        s = "UPDATE Parameter" & session.db.WithRowLock & " SET DBVersion=" & tmpDBVersion
        s += " WHERE DBVersion IS NULL OR DBVersion<" & tmpDBVersion
        session.db.sqlExecute(s)
        session.appDBVersion = tmpDBVersion

        Return True
    End Function

    Private Function UpdateDatabaseV1000001() As Boolean
        Dim tmpDBVersion As Integer = 1000001
        Dim s As String = ""

        With session.db
            If Not .TableExist("Product") Then
                .AddField("Product", "ProductID", "AUTOINCREMENT", True)
                .AddField("Product", "ProductName", "TEXT(50)") : .CreateIndex("Product", "ProductName", True)
            End If

            If Not .TableExist("Wire") Then
                .AddField("Wire", "WireID", "AUTOINCREMENT", True)
                .AddField("Wire", "TypeName", "TEXT(50)")
                .AddField("Wire", "Color", "TEXT(10)")
                .AddField("Wire", "CrossSection", "FLOAT(53)")
                .CreateIndex("Wire", "TypeName,Color,CrossSection", True)
            End If

            If Not .TableExist("ConnLocation") Then
                .AddField("ConnLocation", "ConnLocationID", "AUTOINCREMENT", True)
                .AddField("ConnLocation", "ShortName", "TEXT(50)") : .CreateIndex("ConnLocation", "ShortName", True)
            End If

            If Not .TableExist("Project") Then
                .AddField("Project", "ProjectID", "AUTOINCREMENT", True)
                .AddField("Project", "ProductID", "INT") : .CreateForeignKey("Product", "ProductID", "Project", "ProductID")
                .AddField("Project", "OrderNo", "TEXT(50)") : .CreateIndex("Project", "OrderNo", True)
                .AddField("Project", "EntryDate", "DATE")
            End If
        End With

        s = "UPDATE Parameter" & session.db.WithRowLock & " SET DBVersion=" & tmpDBVersion
        s += " WHERE DBVersion IS NULL OR DBVersion<" & tmpDBVersion
        session.db.sqlExecute(s)
        session.appDBVersion = tmpDBVersion

        Return True
    End Function

    Private Function UpdateDatabaseV1000002() As Boolean
        Dim tmpDBVersion As Integer = 1000002
        Dim s As String = ""

        With session.db

            If Not .TableExist("ConnQuantity") Then
                .AddField("ConnQuantity", "ConnQuantityID", "AUTOINCREMENT", True)
                .AddField("ConnQuantity", "ProjectID", "INT") : .CreateForeignKey("Project", "ProjectID", "ConnQuantity", "ProjectID")
                .AddField("ConnQuantity", "ConnLocationID", "INT") : .CreateForeignKey("ConnLocation", "ConnLocationID", "ConnQuantity", "ConnLocationID")
                .AddField("ConnQuantity", "WireID", "INT") : .CreateForeignKey("Wire", "WireID", "ConnQuantity", "WireID")
                .AddField("ConnQuantity", "Quantity", "INT")
            End If

        End With

        s = "UPDATE Parameter" & session.db.WithRowLock & " SET DBVersion=" & tmpDBVersion
        s += " WHERE DBVersion IS NULL OR DBVersion<" & tmpDBVersion
        session.db.sqlExecute(s)
        session.appDBVersion = tmpDBVersion

        Return True
    End Function

End Module
