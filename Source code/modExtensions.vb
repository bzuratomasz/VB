Imports System.Runtime.CompilerServices

Module modExtensions

    '***** Boolean *****

    ' Boolean.ToInteger -> 0 wenn Boolean False, 1 wenn Boolean True
    <Extension()>
    Public Function ToInteger(ByVal Value As Boolean) As Integer
        If Value Then Return 1
        Return 0
    End Function

    '***** Button *****

    ' Button.CenterHorizontalyOnParent -> Zentriert einen Button horizontal innerhalb seines Parents
    <Extension()>
    Public Sub CenterHorizontalyOnParent(ByVal Button As Button)
        Button.Location = New Point(CInt(Button.Parent.Width / 2 - Button.Width / 2), Button.Location.Y)
    End Sub

    '***** Form *****
    ' Form.BeginBusyLock -> Form wird deaktiviert und der Wartecursor wird angezeigt
    <Extension()>
    Public Sub BeginBusyLock(ByRef Value As Form)
        Value.Enabled = False
        Application.DoEvents()
        Value.Cursor = Cursors.WaitCursor
    End Sub

    ' Form.EndBusyLock -> Form wird aktiviert und der Standardcursor wird angezeigt
    <Extension()>
    Public Sub EndBusyLock(ByRef Value As Form)
        Value.Enabled = True
        Value.Cursor = Cursors.Default
    End Sub

    '***** Integer *****

    ' IntegerZahl.IsBetween(IntegerMin, IntegerMax) -> True wenn IntegerZahl innerhalb IntegerMin und IntegerMax, auch True wenn IntegerZahl gleich IntegerMin oder IntegerMax, ansonsten immer False
    ' IntegerZahl.IsBetween(IntegerMin, IntegerMax, True) -> True wenn Integer innerhalb IntegerMin und IntegerMax, False wenn IntegerZahl gleich IntegerMin oder IntegerMax, ansonsten immer False
    <Extension()>
    Public Function IsBetween(ByVal Value As Integer, ByVal Min As Integer, ByVal Max As Integer, Optional ByVal Exclude As Boolean = False) As Boolean
        If Exclude Then Return ((Value > Min) AndAlso (Value < Max))
        Return ((Value >= Min) AndAlso (Value <= Max))
    End Function

    ' Integer.ToBoolean -> false wenn Integer 0, True wenn Integer <> 0
    <Extension()>
    Public Function ToBoolean(ByVal Value As Integer) As Boolean
        If Value = 0 Then Return False
        Return True
    End Function

    '***** ISynchronizeInvoke *****

    ' Object.InvokeIfRequired(Sub()|Function()|Delegate|Address of, Object) -> Befindet sich der ActionDelegate in einem anderen Thread, wird ein Invoke ausgelöst und auf die Fertigstellung des ActionDelegates gewartet, andernfalls nur der ActionDelegate ohne Invoke ausgeführt. Dem ActionDelegate kann ein Objekt als Parameter übergeben werden welcher selbst ein Objekt zurückliefern kann
    <Extension()>
    Public Function InvokeIfRequired(ByVal obj As System.ComponentModel.ISynchronizeInvoke, ByVal ActionDelegate As System.Delegate, Optional ByVal Args() As Object = Nothing) As Object
        If Args Is Nothing Then Args = New Object(-1) {}
        If obj.InvokeRequired Then Return obj.Invoke(ActionDelegate, Args)
        Return ActionDelegate.Method.Invoke(obj, Args)
    End Function

    ' Object.BeginInvokeIfRequired(Sub()|Function()|Delegate|Address of, Object) -> Befindet sich der ActionDelegate in einem anderen Thread, wird ein Invoke ausgelöst, jedoch nicht auf die Fertigstellung des ActionDelegates gewartet, andernfalls nur der ActionDelegate ohne Invoke ausgeführt. Dem ActionDelegate kann ein Objekt als Parameter übergeben werden, durch den asynchronen Aufruf kann er jedoch kein Objekt zurückliefern
    <Extension()>
    Public Sub BeginInvokeIfRequired(ByVal obj As System.ComponentModel.ISynchronizeInvoke, ByVal ActionDelegate As System.Delegate, Optional ByVal Args() As Object = Nothing)
        If Args Is Nothing Then Args = New Object(-1) {}
        If obj.InvokeRequired Then obj.BeginInvoke(ActionDelegate, Args) : Exit Sub
        ActionDelegate.Method.Invoke(obj, Args)
    End Sub

    '***** List *****

    ' String() -> String
    <Extension()>
    Public Function ToStringOfCommaSeparatedValues(ByVal arr As String(), Optional ByVal Separator As String = ",") As String
        If arr Is Nothing OrElse arr.Length = 0 Then Return ""
        Return String.Join(Separator, arr)
    End Function

    ' Integer() -> String
    <Extension()>
    Public Function ToStringOfCommaSeparatedValues(ByVal arr As Integer(), Optional ByVal Separator As String = ",") As String
        If arr Is Nothing OrElse arr.Length = 0 Then Return ""
        Return String.Join(Separator, System.Array.ConvertAll(arr, Function(Value As Integer) Value.ToString()))
    End Function

    ' List(Of String) -> String
    <Extension()>
    Public Function ToStringOfCommaSeparatedValues(ByVal lst As List(Of String), Optional ByVal Separator As String = ",") As String
        If lst Is Nothing OrElse lst.Count = 0 Then Return ""
        Return String.Join(Separator, lst.ToArray)
    End Function

    ' List(Of Integer) -> String
    <Extension()>
    Public Function ToStringOfCommaSeparatedValues(ByVal lst As List(Of Integer), Optional ByVal Separator As String = ",") As String
        If lst Is Nothing OrElse lst.Count = 0 Then Return ""
        Return String.Join(Separator, Array.ConvertAll(lst.ToArray, Function(Value As Integer) Value.ToString()))
    End Function

    ' String -> String()
    <Extension()>
    Public Function ToArrayOfString(ByVal CommaSeparatedValues As String, Optional ByVal Separator As String = ",") As String()
        If CommaSeparatedValues = "" Then Return New String() {}
        Return CommaSeparatedValues.Split(New String() {Separator}, StringSplitOptions.RemoveEmptyEntries)
    End Function

    ' String -> Integer()
    <Extension()>
    Public Function ToArrayOfInteger(ByVal CommaSeparatedValues As String, Optional ByVal Separator As String = ",") As Integer()
        If CommaSeparatedValues = "" Then Return New Integer() {}
        Return Array.ConvertAll(CommaSeparatedValues.Split(New String() {Separator}, StringSplitOptions.RemoveEmptyEntries), AddressOf Int32.Parse)
    End Function

    ' String -> List(Of String)
    <Extension()>
    Public Function ToListOfString(ByVal CommaSeparatedValues As String, Optional ByVal Separator As String = ",") As List(Of String)
        If CommaSeparatedValues = "" Then Return New List(Of String)
        Return CommaSeparatedValues.Split(New String() {Separator}, StringSplitOptions.RemoveEmptyEntries).ToList
    End Function

    ' String -> List(Of Integer)
    <Extension()>
    Public Function ToListOfInteger(ByVal CommaSeparatedValues As String, Optional ByVal Separator As String = ",") As List(Of Integer)
        If CommaSeparatedValues = "" Then Return New List(Of Integer)
        Return CommaSeparatedValues.Split(New String() {Separator}, StringSplitOptions.RemoveEmptyEntries).ToList.ConvertAll(AddressOf Int32.Parse)
    End Function

    '***** ListView *****

    ' ListView.SelectAll -> Selektiert alle Einträge der ListView
    <Extension()>
    Public Sub SelectAll(ByVal ListView As ListView)
        If ListView Is Nothing Then Exit Sub

        For Each Item As ListViewItem In ListView.Items
            Item.Selected = True
        Next
    End Sub

    ' ListView.SelectNone -> Deselektiert alle Einträge der ListView
    <Extension()>
    Public Sub SelectNone(ByVal ListView As ListView)
        If ListView Is Nothing Then Exit Sub

        For Each Item As ListViewItem In ListView.Items
            Item.Selected = False
        Next
    End Sub

    ' ListView.SelectInverse -> Deselektiert alle Einträge der ListView
    <Extension()>
    Public Sub SelectInverse(ByVal ListView As ListView)
        If ListView Is Nothing Then Exit Sub

        For Each Item As ListViewItem In ListView.Items
            Item.Selected = Not Item.Selected
        Next
    End Sub

    '***** String *****

    ' String.ToDurchmesser -> "Ø String"
    <Extension()>
    Public Function ToDurchmesser(ByVal Value As String) As String
        Return "Ø " & Value
    End Function

    ' String.ToEbenheit -> "⏥ String"
    <Extension()>
    Public Function ToEbenheit(ByVal Value As String) As String
        Return Value ' Unicodezeichen von ArialUnicode noch nicht unterstützt
    End Function

    ' String.ToGeradheit -> "⏤ String"
    <Extension()>
    Public Function ToGeradheit(ByVal Value As String) As String
        Return Value ' Unicodezeichen von ArialUnicode noch nicht unterstützt
    End Function

    ' String.ToGesamtlauf -> "⌰ String"
    <Extension()>
    Public Function ToGesamtlauf(ByVal Value As String) As String
        Return Value ' Unicodezeichen von ArialUnicode noch nicht unterstützt
    End Function

    ' String.ToGewindebohrung -> "M String"
    <Extension()>
    Public Function ToGewindebohrung(ByVal Value As String) As String
        Return "M " & Value
    End Function

    ' String.ToKlammermass -> "(String)"
    <Extension()>
    Public Function ToKlammermass(ByVal Value As String) As String
        Return "(" & Value & ")"
    End Function

    ' String.ToKonzentrizitaet -> "◎ String"
    <Extension()>
    Public Function ToKonzentrizitaet(ByVal Value As String) As String
        Return "◎ " & Value
    End Function

    ' String.ToLinienprofil -> "⌒ String"
    <Extension()>
    Public Function ToLinienprofil(ByVal Value As String) As String
        Return "⌒ " & Value
    End Function

    ' String.ToNeigung -> "∠ String"
    <Extension()>
    Public Function ToNeigung(ByVal Value As String) As String
        Return Value ' Unicodezeichen von ArialUnicode noch nicht unterstützt
    End Function

    ' String.ToOberflaechenprofil -> "⌓ String"
    <Extension()>
    Public Function ToOberflaechenprofil(ByVal Value As String) As String
        Return Value ' Unicodezeichen von ArialUnicode noch nicht unterstütz
    End Function

    ' String.ToParallelitaet -> "∥ String"
    <Extension()>
    Public Function ToParallelitaet(ByVal Value As String) As String
        Return "∥ " & Value
    End Function

    ' String.ToPosition -> "⌖ String"
    <Extension()>
    Public Function ToPosition(ByVal Value As String) As String
        Return Value ' Unicodezeichen von ArialUnicode noch nicht unterstützt
    End Function

    ' String.ToRadius -> "R String"
    <Extension()>
    Public Function ToRadius(ByVal Value As String) As String
        Return "R " & Value
    End Function

    ' String.ToRechtwinkligkeit -> "⟂ String"
    <Extension()>
    Public Function ToRechtwinkligkeit(ByVal Value As String) As String
        Return Value ' Unicodezeichen von ArialUnicode noch nicht unterstützt
    End Function

    ' String.ToRundheit -> "○ String"
    <Extension()>
    Public Function ToRundheit(ByVal Value As String) As String
        Return "○ " & Value
    End Function

    ' String.ToRundlauf -> "↗ String"
    <Extension()>
    Public Function ToRundlauf(ByVal Value As String) As String
        Return "↗ " & Value
    End Function

    ' String.ToSymmetrie -> "⌯ String"
    <Extension()>
    Public Function ToSymmetrie(ByVal Value As String) As String
        Return Value ' Unicodezeichen von ArialUnicode noch nicht unterstützt
    End Function

    ' String.ToWinkel -> "String°"
    <Extension()>
    Public Function ToWinkel(ByVal Value As String) As String
        Return Value & "°"
    End Function

    ' String.ToZylindrizitaet -> "⌭ String"
    <Extension()>
    Public Function ToZylindrizitaet(ByVal Value As String) As String
        Return Value ' Unicodezeichen von ArialUnicode noch nicht unterstützt
    End Function

    ' ListView
    <Extension()>
    Public Function HasItems(ByVal Value As ListView) As Boolean
        Return Value.Items.Count > 0
    End Function

    <Extension()>
    Public Function SelectedTagOrNothing(ByVal Value As ListView, Optional ShowErrorOnNothing As Boolean = True, Optional ErrorMessage As String = Nothing) As Object
        If Value.SelectedItems.Count = 0 Then
            If ShowErrorOnNothing Then
                ErrorMessage = If(ErrorMessage Is Nothing, My.Resources.resGlobal.MsgPleaseSelectEntry, ErrorMessage)
                clsShow.ErrorMsg(ErrorMessage)
            End If

            Return Nothing
        End If

        Return Value.SelectedItems(0).Tag
    End Function

    <Extension()>
    Public Function SelectedTagsOrNothing(ByVal Value As ListView, Optional ShowErrorOnNothing As Boolean = True, Optional ErrorMessage As String = Nothing) As Object()
        If Value.SelectedItems.Count = 0 Then
            If ShowErrorOnNothing Then
                ErrorMessage = If(ErrorMessage Is Nothing, My.Resources.resGlobal.MsgPleaseSelectEntry, ErrorMessage)
                clsShow.ErrorMsg(ErrorMessage)
            End If

            Return Nothing
        End If

        Dim SelectedTags(Value.SelectedItems.Count - 1) As Object

        For i As Integer = 0 To Value.SelectedItems.Count - 1
            SelectedTags(i) = Value.SelectedItems(i).Tag
        Next

        Return SelectedTags
    End Function

    <Extension()>
    Public Sub EnableSortByColumn(ByVal Value As ListView, Optional NumericColIdxs As Integer() = Nothing, Optional WithColoring As Boolean = True)
        If NumericColIdxs Is Nothing Then NumericColIdxs = {}
        If Value.ListViewItemSorter Is Nothing Then clsListView.SortAscDescInit(Value, 0, SortOrder.Ascending)
        If WithColoring Then clsListView.AlternateRowColor(Value)

        AddHandler Value.ColumnClick, Function(sender As System.Object, e As System.Windows.Forms.ColumnClickEventArgs)
                                          clsListView.SortAscDescByColIdx(Value, e.Column, NumericColIdxs.Contains(e.Column))
                                          If WithColoring Then clsListView.AlternateRowColor(Value)
                                          Return True
                                      End Function
    End Sub

    ' Combobox
    <Extension()>
    Public Sub FillWithEnum(ByVal Value As ComboBox, EnumType As Type, EnumPrinter As Func(Of Object, String), Optional WithAll As Boolean = False)
        If EnumPrinter Is Nothing Then EnumPrinter = Function(EnumItem As Object) EnumItem.ToString()

        Value.Items.Clear()

        If WithAll Then Value.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextAll, 0))

        For Each EnumValue In System.Enum.GetValues(EnumType)
            Value.Items.Add(New clsListBoxItem(EnumPrinter(EnumValue), EnumValue))
        Next

        If WithAll Then clsListBoxItem.setSelID(Value, 0)
    End Sub

    <Extension()>
    Public Sub FillWithListItems(ByVal Value As ComboBox, Items As IEnumerable(Of Object), ItemName As Func(Of Object, String), ItemID As Func(Of Object, Integer), Optional WithAll As Boolean = False)
        Dim selID As Integer = If(Value.Items.Count > 0, clsListBoxItem.getSelID(Value), 0)
        selID = If((From i In Items Where ItemID(i) = selID Select i).Count > 0, selID, 0)

        Value.Items.Clear()

        If WithAll Then Value.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextAll, 0))

        Value.Items.AddRange((From i In Items Select New clsListBoxItem(ItemName(i), ItemID(i))).ToArray)

        If WithAll Then clsListBoxItem.setSelID(Value, selID)
    End Sub

    <Extension()>
    Public Sub FillWithSqlRecords(ByVal cmb As ComboBox, Records As IEnumerable(Of clsSQLRecord), Optional withAll As Boolean = False, Optional ordered As Boolean = True)
        Dim selID As Integer = If(cmb.Items.Count > 0, cmb.GetSelID(), 0)
        selID = If((From r In Records Where r.ID = selID Select r).Count > 0, selID, 0)

        cmb.Items.Clear()

        If withAll Then cmb.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextAll, 0))
        If ordered Then Records = Records.OrderBy(Function(r) r.ToString)

        cmb.Items.AddRange((From r In Records Select New clsListBoxItem(r.ToString, r.ID)).ToArray)

        If withAll Then cmb.SetSelID(selID)
    End Sub

    <Extension()>
    Public Sub FillWithSqlRecords(ByVal cmb As ComboBox, Records As IEnumerable(Of clsSQLRecord), Flags As Integer)
        Dim selID As Integer = If(cmb.Items.Count > 0, cmb.GetSelID(), 0)
        selID = If((From r In Records Where r.ID = selID Select r).Count > 0, selID, 0)

        cmb.Items.Clear()

        If Flags And EnumCmbFlags.WithPleaseSelect Then cmb.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextPleaseSelect, 0)) _
        Else If Flags And EnumCmbFlags.WithAll Then cmb.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextAll, 0)) _
        Else If Flags And EnumCmbFlags.WithNoChoice Then cmb.Items.Add(New clsListBoxItem(" - ", 0))
        If Flags And EnumCmbFlags.Ordered Then Records = Records.OrderBy(Function(r)
                                                                             Dim s As String = r.ToString()
                                                                             Dim d As Double
                                                                             If Double.TryParse(s.Replace(",", "."),
                                                                                                Globalization.NumberStyles.Float,
                                                                                                Globalization.CultureInfo.InvariantCulture,
                                                                                                d) Then
                                                                                 Return d
                                                                             Else
                                                                                 Return s
                                                                             End If
                                                                         End Function)

        cmb.Items.AddRange((From r In Records Select New clsListBoxItem(r.ToString, r.ID)).ToArray)

        If Flags And (EnumCmbFlags.WithPleaseSelect Or EnumCmbFlags.WithAll Or EnumCmbFlags.WithNoChoice) Then cmb.SetSelID(selID)
    End Sub

    <Extension()>
    Public Function GetSelID(ByVal c As ComboBox, Optional ByVal nullWert As Integer = 0) As Integer
        If c.SelectedItem Is Nothing Then
            Return nullWert
        Else
            Return CType(c.SelectedItem, clsListBoxItem).ID
        End If
    End Function

    <Extension()>
    Public Sub SetSelID(ByVal c As ComboBox, ByVal id As Integer)
        Dim o As Object
        For Each o In c.Items
            Dim lbi As clsListBoxItem = CType(o, clsListBoxItem)
            If lbi.ID = id Then c.SelectedItem = o : Return
        Next
        c.SelectedIndex = -1
    End Sub

    <Extension()>
    Public Sub BindValidationLabel(ByVal Value As ComboBox, lbl As Label)
        lbl.ForeColor = GetMandatoryColor(clsListBoxItem.getSelID(Value) = 0)
        AddHandler Value.SelectedIndexChanged, Function(sender As Object, e As EventArgs)
                                                   lbl.ForeColor = GetMandatoryColor(clsListBoxItem.getSelID(CType(sender, ComboBox)) = 0)
                                                   Return True
                                               End Function
    End Sub

    ' Checked list box
    <Extension()>
    Public Sub FillWithEnum(ByVal Value As CheckedListBox, EnumType As Type, EnumPrinter As Func(Of Object, String), Optional WithAll As Boolean = False)
        If EnumPrinter Is Nothing Then EnumPrinter = Function(EnumItem As Object) EnumItem.ToString()

        Value.Items.Clear()

        If WithAll Then Value.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextAll, 0))

        For Each EnumValue In System.Enum.GetValues(EnumType)
            Value.Items.Add(New clsListBoxItem(EnumPrinter(EnumValue), EnumValue))
        Next

        If WithAll Then clsListBoxItem.setSelID(Value, 0)
    End Sub

    <Extension()>
    Public Sub FillWithSqlRecords(ByVal clb As CheckedListBox, Records As IEnumerable(Of clsSQLRecord), Flags As Integer, Optional Ordering As Func(Of clsSQLRecord, Object) = Nothing)
        Dim checkIDs = clsListBoxItem.getCheckIDs(clb)

        clb.Items.Clear()

        If Flags And EnumCmbFlags.Ordered Then
            If Ordering Is Nothing Then
                Ordering = Function(r)
                               Dim s As String = r.ToString()
                               Dim d As Double
                               If Double.TryParse(s.Replace(",", "."),
                                                  Globalization.NumberStyles.Float,
                                                  Globalization.CultureInfo.InvariantCulture,
                                                  d) Then
                                   Return d
                               Else
                                   Return s
                               End If
                           End Function
            End If

            Records = Records.OrderBy(Ordering)
        End If

        If Flags And EnumCmbFlags.WithAll Then clb.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextAll, 0))

        clb.Items.AddRange((From r In Records Select New clsListBoxItem(r.ToString(), r.ID)).ToArray)

        clsListBoxItem.setCheckIDs(clb, checkIDs)
    End Sub

    <Extension()>
    Public Sub BindValidationLabel(clb As CheckedListBox, lbl As Label)
        lbl.ForeColor = GetMandatoryColor(clsListBoxItem.getCheckIDs(clb) = "")
        AddHandler clb.SelectedIndexChanged, Function(sender As Object, e As EventArgs)
                                                 lbl.ForeColor = GetMandatoryColor(clsListBoxItem.getCheckIDs(CType(sender, CheckedListBox)) = "")
                                                 Return True
                                             End Function
    End Sub

    ' TextBox
    <Extension()>
    Public Sub BindValidationLabel(ByVal Value As TextBox, lbl As Label)
        lbl.ForeColor = GetMandatoryColor(Value.Text.Trim = "")
        AddHandler Value.TextChanged, Function(sender As Object, e As EventArgs)
                                          lbl.ForeColor = GetMandatoryColor(CType(sender, TextBox).Text.Trim = "")
                                          Return True
                                      End Function
    End Sub

End Module
