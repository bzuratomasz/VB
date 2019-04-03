Public MustInherit Class clsSQLRecord
    Public ID As Integer
    Public RecordSetMode As EnumRecordSetMode = EnumRecordSetMode.rmAddNew

    Protected TableName As String
    Protected PKName As String
    Protected Fields As New ArrayList

    Public Overridable Function GetFields() As ArrayList
        Return New ArrayList(Fields)
    End Function

    Public Overridable Sub SetFieldsFromDataReader(dr As clsDataReader)
        ID = dr.getInt(PKName)
        For Each field As clsRecordsetfield In Fields
            field.FieldValue = dr.getObject(field.FieldName)
            If field.FieldValue Is DBNull.Value Then field.FieldValue = Nothing
        Next
        RecordSetMode = EnumRecordSetMode.rmEdit
    End Sub

    Public Overridable Sub Load(ID As Integer)
        Me.ID = ID
        Using dr As New clsDataReader
            dr.OpenReadonly(session.db, "SELECT * FROM " & TableName & session.db.WithNoLock & " WHERE " & PKName & "=" & ID)
            If dr.Read Then
                For Each field As clsRecordsetfield In Fields
                    field.FieldValue = dr.getObject(field.FieldName)
                    If field.FieldValue Is DBNull.Value Then field.FieldValue = Nothing
                Next
                RecordSetMode = EnumRecordSetMode.rmEdit
            End If
        End Using
    End Sub

    Public Overridable Sub LoadBy(FieldName As String, FieldValue As Object)
        Dim val As String = ""
        If FieldValue Is Nothing Then
            val = " IS NULL "
        ElseIf FieldValue.GetType Is GetType(String) Then
            val = " LIKE "
        ElseIf FieldValue.GetType Is GetType(Boolean) Then
            val = " IS "
        Else
            val = "="
        End If
        val &= session.db.sqlValue(FieldValue)

        Using dr As New clsDataReader
            dr.OpenReadonly(session.db, "SELECT * FROM " & TableName & session.db.WithNoLock & " WHERE " & FieldName & val)
            If dr.Read Then
                ID = dr.getInt(PKName)
                For Each field As clsRecordsetfield In Fields
                    field.FieldValue = dr.getObject(field.FieldName)
                    If field.FieldValue Is DBNull.Value Then field.FieldValue = Nothing
                Next
                RecordSetMode = EnumRecordSetMode.rmEdit
            End If
        End Using
    End Sub

    Public Overridable Sub Update()
        If RecordSetMode = EnumRecordSetMode.rmEdit Then Exit Sub
        If RecordSetMode = EnumRecordSetMode.rmDelte Then Delete() : Exit Sub

        Using dw As New clsDataWriter
            If RecordSetMode = EnumRecordSetMode.rmAddNew Then
                dw.OpenAddNew(session.db, TableName, PKName)
            Else
                dw.OpenEdit(session.db, TableName, PKName, ID)
            End If
            For i = 0 To Fields.Count - 1
                dw.SetFieldValue(Fields(i).FieldName, Fields(i).FieldValue, Fields(i).Direct)
            Next

            dw.Update()
            If RecordSetMode = EnumRecordSetMode.rmAddNew Then ID = dw.GetPrimaryKeyValue
            RecordSetMode = EnumRecordSetMode.rmEdit
        End Using
    End Sub

    Protected Sub Delete()
        session.db.sqlExecute("DELETE FROM " & TableName & " WHERE " & PKName & "=" & ID)
    End Sub

    '-----
    Public Shared Function GetList(sql As String, sqlRecordType As Type) As List(Of clsSQLRecord)
        If Not sqlRecordType.IsSubclassOf(GetType(clsSQLRecord)) Then Throw New ArgumentException("Provided type must inherit from clsSqlRecord")

        Dim RecordsList As New List(Of clsSQLRecord)

        Using dr As New clsDataReader
            dr.OpenReadonly(session.db, sql)

            While dr.Read
                Dim Record As clsSQLRecord = Activator.CreateInstance(sqlRecordType)
                Record.SetFieldsFromDataReader(dr)
                RecordsList.Add(Record)
            End While
        End Using

        Return RecordsList
    End Function

    Public Shared Function SaveRecords(records As List(Of clsSQLRecord)) As Boolean
        If records Is Nothing Then Return False
        If records.Count = 0 Then Return True

        Dim insRecords As List(Of clsSQLRecord) = (From record As clsSQLRecord In records Select record Where record.RecordSetMode = EnumRecordSetMode.rmAddNew).ToList
        Dim updRecords As List(Of clsSQLRecord) = (From record As clsSQLRecord In records Select record Where record.RecordSetMode = EnumRecordSetMode.rmChanged).ToList
        Dim delRecords As List(Of clsSQLRecord) = (From record As clsSQLRecord In records Select record Where record.RecordSetMode = EnumRecordSetMode.rmDelte).ToList

        Dim sqlStatements As List(Of String) = PrepareSQLStatements4Insert(insRecords)
        sqlStatements.AddRange(PrepareSQLStatements4Update(updRecords))
        sqlStatements.AddRange(PrepareSQLStatements4Delete(delRecords))

        Dim i As Integer = 0
        Dim range As Integer = 100000
        While i < sqlStatements.Count
            Using transaction As New clsSqlTransaction(session.db, sqlStatements.Skip(i).Take(range).ToList)
                If Not transaction.IsCommited Then
                    clsShow.ErrorMsg("Transaction failed: " & transaction.ErrorMessage)
                    Return False
                End If
            End Using
            i += range
        End While

        Return True
    End Function

    Public Shared Function CopyRecords(records As List(Of clsSQLRecord)) As Boolean
        If records Is Nothing Then Return False
        If records.Count = 0 Then Return True

        Dim i As Integer = 0
        Dim sqlStatements As New List(Of String)
        
        For Each TableRecords In (From r In records Group r By r.TableName Into Group)
            sqlStatements.Add("SET IDENTITY_INSERT " & TableRecords.TableName & " ON")

            Dim PKName = TableRecords.Group.First.PKName

            Using dw As New clsDataWriter
                dw.OpenAddNew(session.db, TableRecords.TableName, PKName)

                For i = 0 To TableRecords.Group.Count - 1
                    Dim Fields = TableRecords.Group(i).GetFields
                    Fields.Add(New clsRecordsetfield With {.FieldName = PKName, .FieldValue = TableRecords.Group(i).ID, .Direct = True})
                    dw.SetRecord4MultiInsert(Fields)
                Next

                sqlStatements.AddRange(dw.UpdateSQLs)
            End Using

            sqlStatements.Add("SET IDENTITY_INSERT " & TableRecords.TableName & " OFF")
        Next

        i = 0
        Dim range As Integer = 100000
        While i < sqlStatements.Count
            Using transaction As New clsSqlTransaction(session.db, sqlStatements.Skip(i).Take(range).ToList)
                If Not transaction.IsCommited Then
                    clsShow.ErrorMsg("Transaction failed: " & transaction.ErrorMessage)
                    Return False
                End If
            End Using
            i += range
        End While

        Return True
    End Function

    Public Shared Sub DeleteByIDs(IDs As Integer(), Prototype As clsSQLRecord)
        Dim IDsList = If(IDs Is Nothing OrElse IDs.Length = 0, "0", IDs.ToStringOfCommaSeparatedValues())
        session.db.sqlExecute("DELETE FROM " & Prototype.TableName & " WHERE " & Prototype.PKName & " IN(" & IDsList & ")")
    End Sub

    Private Shared Function PrepareSQLStatements4Insert(records As List(Of clsSQLRecord)) As List(Of String)
        If records Is Nothing OrElse records.Count = 0 Then Return New List(Of String)

        Dim InsertStmts As New List(Of String)

        For Each TableRecords In (From r In records Group r By r.TableName Into Group)
            Using dw As New clsDataWriter
                dw.OpenAddNew(session.db, TableRecords.TableName, TableRecords.Group.First.PKName)
                For i = 0 To TableRecords.Group.Count - 1
                    dw.SetRecord4MultiInsert(TableRecords.Group(i).GetFields)
                Next
                InsertStmts.AddRange(dw.UpdateSQLs)
            End Using
        Next

        Return InsertStmts
    End Function

    Private Shared Function PrepareSQLStatements4Update(records As List(Of clsSQLRecord)) As List(Of String)
        If records Is Nothing OrElse records.Count = 0 Then Return New List(Of String)

        Dim UpdateStmts As New List(Of String)

        For Each TableRecords In (From r In records Group r By r.TableName Into Group)
            Using dw As New clsDataWriter
                dw.OpenEdit(session.db, TableRecords.TableName, TableRecords.Group.First.PKName, 0)
                For i = 0 To TableRecords.Group.Count - 1
                    dw.SetRecord4MultiUpdate(TableRecords.Group(i).ID, TableRecords.Group(i).GetFields)
                Next
                UpdateStmts.AddRange(dw.UpdateSQLs)
            End Using
        Next

        Return UpdateStmts
    End Function

    Private Shared Function PrepareSQLStatements4Delete(records As List(Of clsSQLRecord)) As List(Of String)
        If records Is Nothing OrElse records.Count = 0 Then Return New List(Of String)

        Dim statements As New List(Of String)

        For i = 0 To records.Count - 1
            statements.Add("DELETE FROM " & records(i).TableName & " WHERE " & records(i).PKName & "=" & records(i).ID)
        Next

        Return statements
    End Function
End Class

Public Class clsSQLProduct : Inherits clsSQLRecord
    Public Property ProductName As String
        Get
            Return Fields(0).FieldValue
        End Get
        Set(value As String)
            Fields(0).FieldValue = value
        End Set
    End Property

    Public Sub New()
        Fields.Add(New clsRecordsetfield With {.FieldName = "ProductName"})

        TableName = "Product"
        PKName = "ProductID"
    End Sub

    Public Sub New(ID As Integer)
        Me.New()
        Me.Load(ID)
    End Sub

    Public Overrides Function ToString() As String
        Return ProductName
    End Function

    '---
    Public Shared Function GetAll() As List(Of clsSQLProduct)
        Dim ProductsList As New List(Of clsSQLProduct)

        Using dr As New clsDataReader
            dr.OpenReadonly(session.db, "SELECT * FROM Product" & session.db.WithNoLock)

            While dr.Read
                Dim Product As New clsSQLProduct
                Product.SetFieldsFromDataReader(dr)
                ProductsList.Add(Product)
            End While
        End Using

        Return ProductsList
    End Function
End Class

Public Class clsSQLWire : Inherits clsSQLRecord
    Public Property TypeName As String
        Get
            Return Fields(0).FieldValue
        End Get
        Set(value As String)
            Fields(0).FieldValue = value
        End Set
    End Property
    Public Property Color As String
        Get
            Return Fields(1).FieldValue
        End Get
        Set(value As String)
            Fields(1).FieldValue = value
        End Set
    End Property
    Public Property CrossSection As Double
        Get
            Return Fields(2).FieldValue
        End Get
        Set(value As Double)
            Fields(2).FieldValue = value
        End Set
    End Property

    Public Sub New()
        Fields.Add(New clsRecordsetfield With {.FieldName = "TypeName"})
        Fields.Add(New clsRecordsetfield With {.FieldName = "Color"})
        Fields.Add(New clsRecordsetfield With {.FieldName = "CrossSection"})

        TableName = "Wire"
        PKName = "WireID"
    End Sub

    Public Sub New(ID As Integer)
        Me.New()
        Me.LoadBy(Me.PKName, ID)
    End Sub

    '---
    Public Shared Function GetAll() As List(Of clsSQLWire)
        Dim WireList As New List(Of clsSQLWire)

        Using dr As New clsDataReader
            dr.OpenReadonly(session.db, "SELECT * FROM Wire" & session.db.WithNoLock)

            While dr.Read
                Dim Wire As New clsSQLWire
                Wire.SetFieldsFromDataReader(dr)
                WireList.Add(Wire)
            End While
        End Using

        Return WireList
    End Function
End Class

Public Class clsSQLConnLocation : Inherits clsSQLRecord
    Public Property ShortName As String
        Get
            Return Fields(0).FieldValue
        End Get
        Set(value As String)
            Fields(0).FieldValue = value
        End Set
    End Property

    Public Sub New()
        Fields.Add(New clsRecordsetfield With {.FieldName = "ShortName"})

        TableName = "ConnLocation"
        PKName = "ConnLocationID"
    End Sub

    '---
    Public Shared Function GetAll() As List(Of clsSQLConnLocation)
        Dim ConnLocationList As New List(Of clsSQLConnLocation)

        Using dr As New clsDataReader
            dr.OpenReadonly(session.db, "SELECT * FROM ConnLocation" & session.db.WithNoLock)

            While dr.Read
                Dim ConnLocation As New clsSQLConnLocation
                ConnLocation.SetFieldsFromDataReader(dr)
                ConnLocationList.Add(ConnLocation)
            End While
        End Using

        Return ConnLocationList
    End Function
End Class

Public Class clsSQLProject : Inherits clsSQLRecord
    Public Property ProductID As Integer
        Get
            Return Fields(0).FieldValue
        End Get
        Set(value As Integer)
            Fields(0).FieldValue = value
        End Set
    End Property
    Public Property OrderNo As String
        Get
            Return Fields(1).FieldValue
        End Get
        Set(value As String)
            Fields(1).FieldValue = value
        End Set
    End Property
    Public Property EntryDate As Date
        Get
            Return If(Fields(2).FieldValue Is Nothing, dat.NullDate, Fields(2).FieldValue)
        End Get
        Set(value As Date)
            Fields(2).FieldValue = value
        End Set
    End Property

    Public Sub New()
        Fields.Add(New clsRecordsetfield With {.FieldName = "ProductID", .Direct = True})
        Fields.Add(New clsRecordsetfield With {.FieldName = "OrderNo"})
        Fields.Add(New clsRecordsetfield With {.FieldName = "EntryDate"})

        TableName = "Project"
        PKName = "ProjectID"
    End Sub

    Public Sub New(ID As Integer)
        Me.New()
        Me.Load(ID)
    End Sub

    '---
    Public Shared Function GetAll() As List(Of clsSQLProject)
        Dim Projects As New List(Of clsSQLProject)

        Using dr As New clsDataReader
            dr.OpenReadonly(session.db, "SELECT * FROM Project" & session.db.WithNoLock)

            While dr.Read
                Dim Project As New clsSQLProject
                Project.SetFieldsFromDataReader(dr)
                Projects.Add(Project)
            End While
        End Using

        Return Projects
    End Function
End Class

Public Class clsSQLConnQuantity : Inherits clsSQLRecord
    Public Property ProjectID As Integer
        Get
            Return Fields(0).FieldValue
        End Get
        Set(value As Integer)
            Fields(0).FieldValue = value
        End Set
    End Property
    Public Property ConnLocationID As Integer
        Get
            Return Fields(1).FieldValue
        End Get
        Set(value As Integer)
            Fields(1).FieldValue = value
        End Set
    End Property
    Public Property WireID As Integer
        Get
            Return Fields(2).FieldValue
        End Get
        Set(value As Integer)
            Fields(2).FieldValue = value
        End Set
    End Property
    Public Property Quantity As Integer
        Get
            Return Fields(3).FieldValue
        End Get
        Set(value As Integer)
            Fields(3).FieldValue = value
        End Set
    End Property
    Public Sub New()
        Fields.Add(New clsRecordsetfield With {.FieldName = "ProjectID", .Direct = True})
        Fields.Add(New clsRecordsetfield With {.FieldName = "ConnLocationID", .Direct = True})
        Fields.Add(New clsRecordsetfield With {.FieldName = "WireID", .Direct = True})
        Fields.Add(New clsRecordsetfield With {.FieldName = "Quantity"})

        TableName = "ConnQuantity"
        PKName = "ConnQuantityID"
    End Sub

    Public Sub New(ID As Integer)
        Me.New()
        Me.Load(ID)
    End Sub

    '---
    Public Shared Function GetAll() As List(Of clsSQLConnQuantity)
        Dim ConnQuantities As New List(Of clsSQLConnQuantity)

        Using dr As New clsDataReader
            dr.OpenReadonly(session.db, "SELECT * FROM ConnQuantity" & session.db.WithNoLock)

            While dr.Read
                Dim connQuantity As New clsSQLConnQuantity
                connQuantity.SetFieldsFromDataReader(dr)
                ConnQuantities.Add(connQuantity)
            End While
        End Using

        Return ConnQuantities
    End Function
End Class