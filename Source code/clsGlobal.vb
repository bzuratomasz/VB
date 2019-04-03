Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO.Compression
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D

Public Class clsGlobal
    Public Shared Function BeginntMit(ByVal lvw As ListView, ByVal was As String) As Boolean
        Dim Sorter As clsListViewItemComparer = lvw.ListViewItemSorter
        Dim colIdx As Integer = 0
        If Sorter IsNot Nothing Then colIdx = Sorter.ColIdx

        Dim s As String = umlaute_raus(was.ToUpper)
        If lvw.MultiSelect Then lvw.SelectedItems.Clear()

        For Each itmX As ListViewItem In lvw.Items
            Dim t As String = tex.umlaute_raus(itmX.SubItems(colIdx).Text.ToUpper)
            If t.EndsWith(" !") And IsNumeric(t.Replace(" !", "")) Then t = t.Replace(" !", "")
            If IsNumeric(t) And IsNumeric(s) Then t = zahl.getLng(t).ToString : s = zahl.getLng(s).ToString
            Dim l As Integer = s.Length
            If t.Length < l Then l = t.Length

            If t.Substring(0, l) = s Then
                itmX.Selected = True
                itmX.Focused = True
                itmX.EnsureVisible()
                Return True
            End If
        Next

        Return False
    End Function
End Class

Public Class clsListBoxItem

    Public Name
    Public ID As Integer

    Public Sub New(ByVal Name As Object, ByVal ID As Integer)
        Me.Name = Name
        Me.ID = ID
    End Sub

    Public Overrides Function ToString() As String
        Return Name.ToString
    End Function

    'für ListBox
    Public Shared Function getSelID(ByVal c As ListBox, Optional ByVal nullWert As Integer = 0) As Integer
        If c.SelectedItem Is Nothing Then
            Return nullWert
        Else
            Return CType(c.SelectedItem, clsListBoxItem).ID
        End If
    End Function

    Public Shared Sub setSelID(ByVal c As ListBox, ByVal id As Integer)
        Dim o As Object
        For Each o In c.Items
            Dim lbi As clsListBoxItem = CType(o, clsListBoxItem)
            If lbi.ID = id Then c.SelectedItem = o : Return
        Next
        c.SelectedIndex = -1
    End Sub

    Public Shared Function getSelIDs(ByVal c As CheckedListBox) As String
        Dim ret As String = ""
        Dim i As Integer
        For i = 0 To c.Items.Count - 1
            If c.GetSelected(i) Then
                Dim lbi As clsListBoxItem = CType(c.Items(i), clsListBoxItem)
                tex.Cat(ret, lbi.ID.ToString(), ",")
            End If
        Next
        Return ret
    End Function

    Public Shared Sub setSelIDs(ByVal c As CheckedListBox, ByVal ids As String)
        Dim i As Integer
        For i = 0 To c.Items.Count - 1
            Dim lbi As clsListBoxItem = CType(c.Items(i), clsListBoxItem)
            c.SetSelected(i, menge.enthaelt(ids, lbi.ID))
        Next
    End Sub

    Public Shared Function getCheckIDs(ByVal c As CheckedListBox) As String
        Dim ret As String = ""
        Dim i As Integer
        For i = 0 To c.Items.Count - 1
            'If c.GetSelected(i) Then
            If c.GetItemChecked(i) Then
                Dim lbi As clsListBoxItem = CType(c.Items(i), clsListBoxItem)
                tex.Cat(ret, lbi.ID.ToString(), ",")
            End If
        Next
        Return ret
    End Function

    Public Shared Function getCheckNames(ByVal c As CheckedListBox) As String
        Dim ret As String = ""
        Dim i As Integer
        For i = 0 To c.Items.Count - 1
            'If c.GetSelected(i) Then
            If c.GetItemChecked(i) Then
                Dim lbi As clsListBoxItem = CType(c.Items(i), clsListBoxItem)
                tex.Cat(ret, lbi.Name, ", ")
            End If
        Next
        Return ret
    End Function

    Public Shared Sub setCheckIDs(ByVal c As CheckedListBox, ByVal ids As String)
        Dim i As Integer
        For i = 0 To c.Items.Count - 1
            Dim lbi As clsListBoxItem = CType(c.Items(i), clsListBoxItem)
            'c.SetSelected(i, menge.enthaelt(ids, lbi.ID))
            c.SetItemChecked(i, menge.enthaelt(ids, lbi.ID))
        Next
    End Sub
    'für ComboBox
    Public Shared Function getSelID(ByVal c As ComboBox, Optional ByVal nullWert As Integer = 0) As Integer
        If c.SelectedItem Is Nothing Then
            Return nullWert
        Else
            Return CType(c.SelectedItem, clsListBoxItem).ID
        End If
    End Function

    Public Shared Sub setSelID(ByVal c As ComboBox, ByVal id As Integer)
        Dim o As Object
        For Each o In c.Items
            Dim lbi As clsListBoxItem = CType(o, clsListBoxItem)
            If lbi.ID = id Then c.SelectedItem = o : Return
        Next
        c.SelectedIndex = -1
    End Sub

    Public Shared Function getIDtoString(ByVal c As ComboBox, ByVal exaktstring As String) As Integer
        Dim ret As Integer
        ret = c.FindStringExact(exaktstring)
        Return ret
    End Function

    Public Shared Sub LoadJaNein(ByVal c As ComboBox, Optional ByVal SelectedIndex As Integer = -1)
        With c.Items
            .Clear()
            .Add(New clsListBoxItem(" ", 0))
            .Add(New clsListBoxItem(My.Resources.resGlobal.TextYes, 1))
            .Add(New clsListBoxItem(My.Resources.resGlobal.TextNo, 2))
        End With
    End Sub

    Public Shared Sub DeselectCombo(ByVal ComboBox As Object, ByVal btnAbbrechen As Control)

        Dim cmb As ComboBox = CType(ComboBox, ComboBox)
        If ControlIsLocked(cmb) Then
            btnAbbrechen.Focus() 'Notlösung, da kein ReadOnly für Combos existiert
        End If
    End Sub

    Public Shared Sub FillCmbFromDictionary(cmb As ComboBox, dict As Dictionary(Of Integer, String), Optional withAll As Boolean = False)
        cmb.Items.Clear()

        If withAll Then cmb.Items.Add(New clsListBoxItem(My.Resources.resGlobal.TextAll, 0))

        For Each Pair In dict
            If Pair.Key = 0 Then Continue For
            cmb.Items.Add(New clsListBoxItem(Pair.Value, Pair.Key))
        Next
    End Sub
End Class

Public Class clsListViewItem

    Public Shared Function getAllIDs(ByVal lvw As ListView) As String
        Dim ret As String = ""
        Dim itmX As ListViewItem
        For Each itmX In lvw.Items
            tex.Cat(ret, itmX.Tag.ToString(), ",")
        Next
        Return ret
    End Function

    Public Shared Function getCheckIDs(ByVal lvw As ListView) As String
        Dim ret As String = ""
        Dim itmX As ListViewItem
        For Each itmX In lvw.CheckedItems
            tex.Cat(ret, itmX.Tag.ToString(), ",")
        Next
        Return ret
    End Function

    Public Shared Sub setCheckIDs(ByVal lvw As ListView, ByVal ids As String)
        Dim itmX As ListViewItem
        For Each itmX In lvw.Items
            itmX.Checked = (menge.enthaelt(ids, CStr(itmX.Tag)))
        Next
    End Sub

    Public Shared Function getSelIDs(ByVal lvw As ListView) As String
        Dim ret As String = ""
        Dim itmX As ListViewItem
        For Each itmX In lvw.SelectedItems
            tex.Cat(ret, itmX.Tag.ToString(), ",")
        Next
        Return ret
    End Function

    Public Shared Sub setSelIDs(ByVal lvw As ListView, ByVal ids As String)
        Dim itmX As ListViewItem
        Dim idx As Integer = 0
        For Each itmX In lvw.Items
            Dim b As Boolean = menge.enthaelt(ids, CStr(itmX.Tag))
            itmX.Selected = b
            If b Then idx = itmX.Index 'letztes Element darstellen, srcollt bei zerstückelten Selektionen nach unten
        Next
        If idx > 0 Then lvw.Items(idx).EnsureVisible()
    End Sub

    Public Shared Sub setSelID(ByVal lvw As ListView, ByVal key As Object)
        Dim lvi As ListViewItem
        Dim i As Integer
        For i = 0 To lvw.Items.Count - 1
            lvi = lvw.Items(i)
            Dim compkey As Object = Nothing
            If key.GetType.Name = "String" Then
                compkey = CType(lvi.Tag, String)
            Else
                compkey = zahl.getLng(CType(lvi.Tag, String))
            End If
            If IsNothing(compkey) Then Continue For
            If compkey = key Then
                lvi.Selected = True
                lvi.EnsureVisible()
                lvi.Focused = True
                Exit Sub
            End If
        Next
    End Sub

    Public Shared Function GetListviewSelItem(ByVal l As ListView) As ListViewItem
        Dim c As ListView.SelectedListViewItemCollection = l.SelectedItems
        If c.Count = 0 Then
            Return Nothing
        Else
            Dim lvi As ListViewItem
            For Each lvi In c
                If lvi.Selected Then Return lvi
            Next
            Return Nothing
        End If
    End Function

    Public Shared Function GetNextIntegerItmIDForSelectAfterDelete(ByVal lvw As ListView) As Integer
        If lvw.SelectedItems.Count <> 1 Or lvw.Items.Count < 2 Then Return 0

        Dim idx As Integer = lvw.SelectedItems(0).Index
        If idx = lvw.Items.Count - 1 Then
            idx -= 1
        Else
            idx += 1
        End If
        Dim ret As Integer = 0
        'If idx >= 0 Then
        ret = CInt(lvw.Items(idx).Tag)
        'End If
        Return ret
    End Function

    Public Shared Function GetNextStringItmIDForSelectAfterDelete(ByVal lvw As ListView) As String
        If lvw.SelectedItems.Count <> 1 Or lvw.Items.Count < 2 Then Return ""

        Dim idx As Integer = lvw.SelectedItems(0).Index
        If idx = lvw.Items.Count - 1 Then
            idx -= 1
        Else
            idx += 1
        End If
        Dim ret As String = ""
        ret = CStr(lvw.Items(idx).Tag)
        Return ret
    End Function

    Public Shared Sub Export2CSV(ByVal lvw As ListView)
        Dim dstPfad As String = ""

        Dim sfDialog As New SaveFileDialog
        With sfDialog
            .CheckFileExists = False
            .CheckPathExists = False
            .DefaultExt = "csv"
            .FileName = ""
            .Filter = My.Resources.resGlobal.TextCSVFiles & " (*.csv)|*.csv|" & My.Resources.resGlobal.TextAllFiles & " (*.*)|*.*"
            '.InitialDirectory = "..."

            If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                dstPfad = .FileName
            Else
                Exit Sub
            End If
        End With

        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("iso-8859-1") '"ibm850", "iso-8859-1", "IBM437"
        Dim sw As New System.IO.StreamWriter(dstPfad, False, enc)

        Dim i As Integer
        Dim r As String = ""
        Dim itmX As ListViewItem

        For i = 0 To lvw.Columns.Count - 1
            tex.Cat(r, lvw.Columns(i).Text, ";")
        Next
        sw.WriteLine(r)
        For Each itmX In lvw.Items
            r = ""
            For i = 0 To itmX.SubItems.Count - 1
                tex.Cat(r, itmX.SubItems(i).Text.Replace(";", ","), ";")
            Next
            sw.WriteLine(r)
        Next
        sw.Close()

        datei.Open(dstPfad)
    End Sub

    Public Shared Sub SelectAll(ByVal lvw As ListView)
        Dim itmX As ListViewItem
        For Each itmX In lvw.Items
            itmX.Selected = True
        Next
    End Sub

    Public Shared Sub DeselectAll(ByVal lvw As ListView)
        Dim itmX As ListViewItem
        For Each itmX In lvw.Items
            itmX.Selected = False
        Next
    End Sub

    Public Shared Sub CheckedAll(ByVal lvw As ListView)
        Dim itmX As ListViewItem
        For Each itmX In lvw.Items
            itmX.Checked = True
        Next
    End Sub

    Public Shared Sub UncheckedtAll(ByVal lvw As ListView)
        Dim itmX As ListViewItem
        For Each itmX In lvw.Items
            itmX.Checked = False
        Next
    End Sub

    Public Shared Sub SortAscDescInit(ByVal lvw As ListView, ByVal ColIdx As Integer, ByVal Sort As SortOrder, Optional ByVal IsNumeric As Boolean = False)
        clsListViewColumnSortImage.SetSortImage(lvw, ColIdx, Sort)
        lvw.ListViewItemSorter = New clsListViewItemComparer(ColIdx, IsNumeric, Sort)
    End Sub

    Public Shared Function SortAscDescByColIdx(ByVal lvw As ListView, ByVal colIdx As Integer, Optional ByVal IsNumeric As Boolean = False, Optional ByVal sort As SortOrder = SortOrder.None) As SortOrder
        Dim oldLviSorter As clsListViewItemComparer = CType(lvw.ListViewItemSorter, clsListViewItemComparer)
        If oldLviSorter Is Nothing Then oldLviSorter = New clsListViewItemComparer(0, False, SortOrder.Ascending)

        Dim oldColIdx As Integer = oldLviSorter.ColIdx
        Dim oldSortOrder As SortOrder = oldLviSorter.Sort

        Dim newColIdx As Integer = colIdx
        Dim newSortOrder As SortOrder = CType(IIf(newColIdx <> oldColIdx, SortOrder.Ascending, IIf(oldSortOrder = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)), SortOrder)

        clsListViewColumnSortImage.SetSortImage(lvw, oldColIdx, SortOrder.None)
        clsListViewColumnSortImage.SetSortImage(lvw, colIdx, newSortOrder)

        lvw.ListViewItemSorter = New clsListViewItemComparer(newColIdx, IsNumeric, newSortOrder)
        Return newSortOrder
    End Function
End Class

Public Class clsListViewItemComparer 'zum Sortieren von Spalten innerhalb von Listviews
    Implements IComparer

    Public ColIdx As Integer
    Private IsNumeric As Boolean = False
    Public Sort As SortOrder

    Public Sub New(ByVal ColIdx As Integer, ByVal IsNumeric As Boolean, ByVal Sort As SortOrder)
        Me.ColIdx = ColIdx
        Me.IsNumeric = IsNumeric
        Me.Sort = Sort
    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Dim lviX As ListViewItem = CType(x, ListViewItem)
        Dim lviY As ListViewItem = CType(y, ListViewItem)

        If IsNumeric Then
            Dim lviXVal As Double = zahl.getDecimal(lviX.SubItems(ColIdx).Text)
            Dim lviYVal As Double = zahl.getDecimal(lviY.SubItems(ColIdx).Text)

            Return Math.Sign(lviXVal - lviYVal) * zahl.IIfInt(Sort = SortOrder.Ascending, 1, -1)
        Else
            Dim lviXText As String = lviX.SubItems(ColIdx).Text
            Dim lviYText As String = lviY.SubItems(ColIdx).Text

            Return String.Compare(lviXText, lviYText) * zahl.IIfInt(Sort = SortOrder.Ascending, 1, -1)
        End If
    End Function
End Class

Public Class clsListViewColumnSortImage
    <DllImport("user32")> _
    Private Shared Function SendMessage(ByVal Handle As IntPtr, ByVal msg As Int32, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32", EntryPoint:="SendMessage")> _
    Private Shared Function SendMessage(ByVal Handle As IntPtr, ByVal msg As Int32, ByVal wParam As IntPtr, ByRef lParam As HDITEM) As IntPtr
    End Function

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure HDITEM
        Public mask As Int32
        Public cxy As Int32

        <MarshalAs(UnmanagedType.LPTStr)> _
        Public pszText As [String]

        Public hbm As IntPtr
        Public cchTextMax As Int32
        Public fmt As Int32
        Public lParam As Int32
        Public iImage As Int32
        Public iOrder As Int32
    End Structure

    Public Const HDI_WIDTH As Int32 = &H1
    Public Const HDI_HEIGHT As Int32 = HDI_WIDTH
    Public Const HDI_TEXT As Int32 = &H2
    Public Const HDI_FORMAT As Int32 = &H4
    Public Const HDI_LPARAM As Int32 = &H8
    Public Const HDI_BITMAP As Int32 = &H10
    Public Const HDI_IMAGE As Int32 = &H20
    Public Const HDI_DI_SETITEM As Int32 = &H40
    Public Const HDI_ORDER As Int32 = &H80
    Public Const HDI_FILTER As Int32 = &H100
    ' 0x0500
    Public Const HDF_LEFT As Int32 = &H0
    Public Const HDF_RIGHT As Int32 = &H1
    Const HDF_CENTER As Int32 = &H2

    Public Const HDF_JUSTIFYMASK As Int32 = &H3
    Public Const HDF_RTLREADING As Int32 = &H4
    Public Const HDF_OWNERDRAW As Int32 = &H8000
    Public Const HDF_STRING As Int32 = &H4000
    Public Const HDF_BITMAP As Int32 = &H2000
    Public Const HDF_BITMAP_ON_RIGHT As Int32 = &H1000
    Public Const HDF_IMAGE As Int32 = &H800
    Public Const HDF_SORTUP As Int32 = &H400
    ' 0x0501
    Public Const HDF_SORTDOWN As Int32 = &H200
    ' 0x0501
    Public Const I_IMAGENONE As Integer = -2

    Public Const LVM_FIRST As Int32 = &H1000
    ' List messages
    Public Const LVM_GETHEADER As Int32 = LVM_FIRST + 31

    Public Const HDM_FIRST As Int32 = &H1200
    ' Header messages
    Public Const HDM_SETIMAGELIST As Int32 = HDM_FIRST + 8
    Public Const HDM_GETIMAGELIST As Int32 = HDM_FIRST + 9
    Public Const HDM_GETITEM As Int32 = HDM_FIRST + 11
    Public Const HDM_SETITEM As Int32 = HDM_FIRST + 12

    Public Shared Sub SetSortImage(ByVal lvw As ListView, ByVal ColumnIndex As Integer, ByVal Sort As SortOrder)
        If ColumnIndex < 0 OrElse ColumnIndex >= lvw.Columns.Count Then Exit Sub

        Dim hHeader As IntPtr = SendMessage(lvw.Handle, LVM_GETHEADER, IntPtr.Zero, IntPtr.Zero)
        Dim colHdr As ColumnHeader = lvw.Columns(ColumnIndex)
        Dim hd As New HDITEM()
        Dim align As HorizontalAlignment = colHdr.TextAlign

        hd.mask = HDI_FORMAT

        If align = HorizontalAlignment.Left Then
            hd.fmt = HDF_LEFT Or HDF_STRING Or HDF_BITMAP_ON_RIGHT
        ElseIf align = HorizontalAlignment.Center Then
            hd.fmt = HDF_CENTER Or HDF_STRING Or HDF_BITMAP_ON_RIGHT
        Else
            ' HorizontalAlignment.Right
            hd.fmt = HDF_RIGHT Or HDF_STRING
        End If

        If Sort = SortOrder.Ascending Then
            hd.fmt = hd.fmt Or HDF_SORTUP
        ElseIf Sort = SortOrder.Descending Then
            hd.fmt = hd.fmt Or HDF_SORTDOWN
        ElseIf Sort = SortOrder.None Then
            hd.iImage = I_IMAGENONE
        End If

        SendMessage(hHeader, HDM_SETITEM, New IntPtr(ColumnIndex), hd)

        If lvw.Items.Count > 0 Then lvw.Items(0).Focused = False
    End Sub
End Class

Public Class clsListView
    Private Shared BackColorPrimary As Color = Color.FromArgb(197, 217, 241)
    Private Shared BackColorSecondary As Color = Color.FromArgb(83, 142, 213)

    Public Shared Sub AutoSizeColumnByHeaderOrContent(ByRef lvw As ListView, ByRef ColumnIndex As Integer)
        If (ColumnIndex + 1) > lvw.Columns.Count Then Exit Sub
        If lvw.Columns(ColumnIndex) Is Nothing Then Exit Sub
        lvw.Columns(ColumnIndex).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
        Dim headerwidth As Integer = lvw.Columns(ColumnIndex).Width
        lvw.Columns(ColumnIndex).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
        Dim contentwidth As Integer = lvw.Columns(ColumnIndex).Width
        If headerwidth > contentwidth Then lvw.Columns(ColumnIndex).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
    End Sub

    Public Shared Sub AlternateRowColor(ByRef lvw As ListView, Optional ByVal ColumnIndex As Integer = -1)
        If lvw.ListViewItemSorter Is Nothing Then Exit Sub
        If ColumnIndex = -1 Then ColumnIndex = CType(lvw.ListViewItemSorter, clsListViewItemComparer).ColIdx

        Dim ItemText As String = ""
        Dim AlternateColor As Boolean = True

        Select Case lvw.Groups.Count
            Case 0 To 1
                For Each Item As ListViewItem In lvw.Items
                    If Item.SubItems(ColumnIndex).Text <> ItemText Then AlternateColor = Not AlternateColor
                    ItemText = Item.SubItems(ColumnIndex).Text
                    'If Item.BackColor = Color.Orange Then Continue For ' Erledigte Vorgänge aus Wuxi mit MatText Änderung oder Löschung werden Orange dargestellt, diese Farbe nicht überschreiben
                    'If IsSiemensColorForAlternatingRowColorWarning(Item.BackColor) Then Item.BackColor = GetSiemensColorForAlternatingRowColorWarning(session.Benutzer.AlternatingRowColor) : Continue For ' Zeilen die mit Warnfarbe markiert sind mit gegebenenfalls neuer Warnfarbe markieren
                    If AlternateColor Then Item.BackColor = BackColorSecondary Else Item.BackColor = BackColorPrimary
                Next
            Case Else
                For Each Group As ListViewGroup In lvw.Groups
                    For Each Item As ListViewItem In lvw.Items
                        If Item.Group IsNot Group Then Continue For

                        If Item.SubItems(ColumnIndex).Text <> ItemText Then AlternateColor = Not AlternateColor
                        ItemText = Item.SubItems(ColumnIndex).Text
                        'If Item.BackColor = Color.Orange Then Continue For ' Erledigte Vorgänge aus Wuxi mit MatText Änderung oder Löschung werden Orange dargestellt, diese Farbe nicht überschreiben
                        'If IsSiemensColorForAlternatingRowColorWarning(Item.BackColor) Then Item.BackColor = GetSiemensColorForAlternatingRowColorWarning(session.Benutzer.AlternatingRowColor) : Continue For ' Zeilen die mit Warnfarbe markiert sind mit gegebenenfalls neuer Warnfarbe markieren
                        If AlternateColor Then Item.BackColor = BackColorSecondary Else Item.BackColor = BackColorPrimary
                    Next
                Next
        End Select
    End Sub

    Public Shared Sub SortAscDescInit(ByVal lvw As ListView, ByVal ColIdx As Integer, ByVal Sort As SortOrder, Optional ByVal IsNumeric As Boolean = False)
        clsListViewColumnSortImage.SetSortImage(lvw, ColIdx, Sort)
        lvw.ListViewItemSorter = New clsListViewItemComparer(ColIdx, IsNumeric, Sort)
    End Sub

    Public Shared Function SortAscDescByColIdx(ByVal lvw As ListView, ByVal colIdx As Integer, Optional ByVal IsNumeric As Boolean = False, Optional ByVal sort As SortOrder = SortOrder.None) As SortOrder
        Dim oldLviSorter As clsListViewItemComparer = CType(lvw.ListViewItemSorter, clsListViewItemComparer)
        If oldLviSorter Is Nothing Then oldLviSorter = New clsListViewItemComparer(0, False, SortOrder.Ascending)

        Dim oldColIdx As Integer = oldLviSorter.ColIdx
        Dim oldSortOrder As SortOrder = oldLviSorter.Sort

        Dim newColIdx As Integer = colIdx
        Dim newSortOrder As SortOrder = CType(IIf(newColIdx <> oldColIdx, SortOrder.Ascending, IIf(oldSortOrder = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)), SortOrder)

        clsListViewColumnSortImage.SetSortImage(lvw, oldColIdx, SortOrder.None)
        clsListViewColumnSortImage.SetSortImage(lvw, colIdx, newSortOrder)

        lvw.ListViewItemSorter = New clsListViewItemComparer(newColIdx, isNumeric, newSortOrder)
        Return newSortOrder
    End Function

    Public Shared Sub SortAscDescBySorter(lvw As ListView, sorter As clsListViewItemComparer)
        Dim oldColIdx As Integer = 0
        Dim oldLviSorter As clsListViewItemComparer = CType(lvw.ListViewItemSorter, clsListViewItemComparer)
        If oldLviSorter IsNot Nothing Then oldColIdx = oldLviSorter.ColIdx

        clsListViewColumnSortImage.SetSortImage(lvw, oldColIdx, SortOrder.None)
        If sorter IsNot Nothing Then clsListViewColumnSortImage.SetSortImage(lvw, sorter.ColIdx, sorter.Sort)

        lvw.ListViewItemSorter = sorter
    End Sub

    'Public Shared Sub ColumnToTextBoxAutoComplete(ByVal ListView As ListView, ByVal Sorter As clsListViewItemComparer, ByVal TextBox As TextBox)
    '    Dim StringCollection As New AutoCompleteStringCollection

    '    If Sorter.IsNumeric Then
    '        For Each Item As ListViewItem In ListView.Items
    '            StringCollection.Add(Item.SubItems(Sorter.ColIdx).Text.TrimStart("0"c))
    '        Next
    '    Else
    '        For Each Item As ListViewItem In ListView.Items
    '            StringCollection.Add(Item.SubItems(Sorter.ColIdx).Text)
    '        Next
    '    End If

    '    TextBox.AutoCompleteSource = AutoCompleteSource.CustomSource
    '    TextBox.AutoCompleteMode = AutoCompleteMode.Suggest
    '    TextBox.AutoCompleteCustomSource = StringCollection
    'End Sub
End Class

' === dat =====================================================================
Public Class dat
    Public Shared Function NullDate() As Date
        Return DateValue("1900-01-01")
    End Function

    Public Shared Function DBValue(ByVal d As Date) As Object
        Return IIf(dat.IsNull(d), System.DBNull.Value, d)
    End Function

    Public Shared Function Week(ByVal d As Date) As Integer
        Dim ci As New System.Globalization.CultureInfo("de-DE")
        Dim cal As System.Globalization.Calendar = ci.Calendar
        Return cal.GetWeekOfYear(d, ci.DateTimeFormat.CalendarWeekRule, ci.DateTimeFormat.FirstDayOfWeek)
    End Function

    Public Shared Function IsNull(ByVal d As Date) As Boolean
        Return d.Equals(NullDate()) OrElse d.Equals(DateValue("1900-01-02"))
    End Function

    Public Shared Function IsTime(ByVal d As Date) As Boolean
        Return d.Year = 1 And d.Month = 1 And d.Day = 1 And Not IsNull(d)
    End Function

    Public Shared Function Obj2Date(ByVal o As Object) As Date
        If o Is System.DBNull.Value Then Return NullDate()
        Return CDate(o)
    End Function

    Public Shared Function Format(ByVal d As Date, ByVal form As String) As String
        If IsNull(d) Then Return ""

        Return d.ToString(form)
    End Function

    Public Shared Function printDIN(ByVal d As Date, Optional IfNull As String = "") As String '-> "2006-01-31"
        If IsNull(d) Then Return IfNull

        Return d.ToString("yyyy\-MM\-dd")
    End Function

    Public Shared Function PrintSAP(ByVal d As Date) As String '->20060131
        If IsNull(d) Then Return ""
        Return d.ToString("yyyyMMdd")
    End Function

    Public Shared Function printDINYearWeek(ByVal d As Date) As String
        If IsNull(d) Then Return ""

        Return d.ToString("yyyy") + "-" + Week(d).ToString("00")
    End Function

    Public Shared Function printDINYearMonth(ByVal d As Date, Optional IfNull As String = "") As String
        If IsNull(d) Then Return IfNull

        Return d.ToString("yyyy") + "-" + d.Month.ToString("00")
    End Function

    Public Shared Function printDINTime2(ByVal d As Date) As String '-> "16-00-00"
        If IsNull(d) Then Return ""

        Return d.ToString("HH\-mm\-ss")
    End Function

    Public Shared Function printDINTime(ByVal d As Date) As String '-> "2006-01-31 16:00:00"
        If IsNull(d) Then Return ""

        Return d.ToString("yyyy\-MM\-dd HH\:mm\:ss")
    End Function

    Public Shared Function printDINTimeForFile(ByVal d As Date) As String '-> "2006-01-31_16-00-00"
        If IsNull(d) Then Return ""

        Return d.ToString("yyyy\-MM\-dd_HH\-mm\-ss")
    End Function

    Public Shared Function printTimeStamp(ByVal d As Date) As String '-> "20060131160000"
        If IsNull(d) Then Return ""

        Return d.ToString("yyyyMMddHHmmss")
    End Function

    Public Shared Function printTimeStamp2(ByVal d As Date) As String '-> "20060131_160000"
        If IsNull(d) Then Return ""

        Return d.ToString("yyyyMMdd_HHmmss")
    End Function

    Public Shared Function PrintStandard(ByVal d As Date) As String '-> "31.01.2006"
        If IsNull(d) Then Return ""

        Return d.ToString("dd.MM.yyyy")
    End Function

    Public Shared Function PrintStandardTime(ByVal d As Date) As String '->"30.01.2006 17:43"
        If IsNull(d) Then Return ""
        Return d.ToString("dd.MM.yyyy HH\:mm\:ss")
    End Function

    Public Shared Function PrintCNShort(ByVal d As Date) As String '-> "2014/3/13"
        If IsNull(d) Then Return ""
        Return d.ToString("yyyy\/M\/d")
    End Function

    Public Shared Function PrintCNLong(ByVal d As Date) As String '-> "2014年3月13日"
        If IsNull(d) Then Return ""
        Return d.ToString("yyyy年M月d日")
    End Function

    Public Shared Function printTimeSpan(ts As TimeSpan) As String
        Return ts.Hours.ToString.PadLeft(2, "0") & ":" & ts.Minutes.ToString.PadLeft(2, "0") & ":" & ts.Seconds.ToString.PadLeft(2, "0")
    End Function

    Public Shared Function printFileTimeToTime(ByVal s As String) As Date
        If s = "" Then Return dat.NullDate

        Dim Y As Integer = CInt(s.Substring(0, 4))
        Dim m As Integer = CInt(s.Substring(5, 2))
        Dim d As Integer = CInt(s.Substring(8, 2))

        Dim H As Integer = CInt(s.Substring(11, 2))
        Dim Mi As Integer = CInt(s.Substring(14, 2))
        Dim Se As Integer = CInt(s.Substring(17, 2))

        Dim newDate As New Date(Y, m, d, H, Mi, Se)

        Return newDate
    End Function

    Public Shared Function checkDIN(ByVal s As String) As Boolean
        s = Obj2Str(s).Trim() '2009-02-17 vorsichtshalber, falls bei Aufruf Nothing übergeben wird
        If s = "" Then Return True
        If s.IndexOf(".") >= 0 Then Return False
        If s.Length() <> 10 Then Return False
        If s.Substring(4, 1) <> "-" Or s.Substring(7, 1) <> "-" Then Return False

        Dim Y As Integer = CInt(s.Substring(0, 4))
        Dim m As Integer = CInt(s.Substring(5, 2))
        Dim d As Integer = CInt(s.Substring(8, 2))
        If Y < 1900 Or Y > 2100 Then Return False

        Dim t As Date : t = DateSerial(Y, m, d)
        Return (t.Year = Y And t.Month = m And t.Day = d)
    End Function

    Public Shared Function DINStr2Date(ByVal datstr As String) As Date
        If datstr Is Nothing Or datstr = "" Then Return NullDate()
        If Not checkDIN(datstr) Then Return NullDate()

        Dim y As Integer = CInt(datstr.Substring(0, 4))
        Dim m As Integer = CInt(datstr.Substring(5, 2))
        Dim d As Integer = CInt(datstr.Substring(8, 2))

        Return DateSerial(y, m, d)
    End Function

    Public Shared Function SapStr2Dat(ByVal datum As Object) As Date
        If datum Is Nothing Then Return NullDate()
        If TypeOf (datum) Is Date Then Return CType(datum, Date)
        If TypeOf (datum) Is String Then
            Dim datstr As String = CType(datum, String)
            If datstr = "" Or datstr = "00:00:00" Then Return NullDate()

            Dim retDat As Date = Any2Dat(datstr)
            If retDat = NullDate() And datstr.Length = 8 Then
                Dim y As Integer = CInt(datstr.Substring(0, 4))
                Dim m As Integer = CInt(datstr.Substring(4, 2))
                Dim d As Integer = CInt(datstr.Substring(6, 2))

                Return DateSerial(y, m, d)
            Else
                Return retDat
            End If
        End If
        Return NullDate()
    End Function

    Public Shared Function Any2Dat(ByVal datstr As String) As Date
        Dim ret As Date = NullDate()
        If datstr = "" Then Return ret

        Try
            ret = CType(datstr, Date)
        Catch ex As Exception
            ret = NullDate()
        End Try

        Return ret
    End Function

    Public Shared Function checkTime(ByVal s As String) As Boolean
        s = Obj2Str(s).Trim()
        If s = "" Then Return True
        If s.Length() <> 8 Then Return False
        If s.Substring(2, 1) <> ":" Or s.Substring(5, 1) <> ":" Then Return False

        Dim hh As Integer = CInt(s.Substring(0, 2))
        Dim mm As Integer = CInt(s.Substring(3, 2))
        Dim ss As Integer = CInt(s.Substring(6, 2))
        If hh < 0 Or hh > 23 Or mm < 0 Or mm > 59 Or ss < 0 Or ss > 59 Then Return False

        Dim t As Date : t = TimeSerial(hh, mm, ss)
        Return (t.Hour = hh And t.Minute = mm And t.Second = ss)
    End Function

    Public Shared Function TimeStr2Date(ByVal timstr As String) As Date
        If timstr = "" Then Return NullDate()
        If Not checkTime(timstr) Then Return NullDate()

        Dim hh As Integer = CInt(timstr.Substring(0, 2))
        Dim mm As Integer = CInt(timstr.Substring(3, 2))
        Dim ss As Integer = CInt(timstr.Substring(6, 2))

        Return TimeSerial(hh, mm, ss)
    End Function

    Public Shared Function DINTimeStr2Date(ByVal dtstr As String) As Date
        If dtstr = "" Then Return NullDate()
        Dim datstr As String = tex.Part(dtstr, 1, " ")
        Dim timstr As String = tex.Part(dtstr, 2, " ")
        If datstr = "" Then Return NullDate()
        If timstr = "" Then Return NullDate()
        If Not checkDIN(datstr) Then Return NullDate()
        If Not checkTime(timstr) Then Return NullDate()

        Dim datum As Date = DINStr2Date(datstr)
        Dim uhrzeit As Date = TimeStr2Date(timstr)
        datum = datum.AddHours(uhrzeit.Hour)
        datum = datum.AddMinutes(uhrzeit.Minute)
        datum = datum.AddSeconds(uhrzeit.Second)

        Return datum
    End Function

    Public Shared Function BetweenDate(ByVal datumTest As Date, ByVal datumBeginn As Date, ByVal datumEnde As Date) As Boolean
        Select Case datumTest.Date
            Case datumBeginn.Date : Return True
            Case datumEnde.Date : Return True
            Case datumBeginn.Date To datumEnde.Date : Return True
            Case Else : Return False
        End Select
    End Function

    Public Shared Function MonthDiff(ByVal dateStart As Date, ByVal dateEnd As Date) As Integer
        If dateStart.Date > dateEnd.Date Then
            Dim e As New ArgumentException(My.Resources.resGlobal.MsgEndDateIsBeforeStartDate)
            Throw e
        End If
        Dim ret As Integer = 0
        If dateStart.Year < dateEnd.Year Then ret = 12 * (dateEnd.Year - dateStart.Year)
        ret += dateEnd.Month - dateStart.Month
        Return ret
    End Function

    Public Shared Function StdStr2Date(ByVal datstr As String) As Date
        If datstr = "" Then Return NullDate()
        If Not checkStd(datstr) Then Return NullDate()

        Dim d As Integer = CInt(Left$((tex.Part(datstr, 1, ".")).Trim, 2))
        If d = 0 Then Return NullDate()
        Dim m As Integer = CInt(Left$((tex.Part(datstr, 2, ".")).Trim, 2))
        If m = 0 Then Return NullDate()
        Dim h As String : h = Left$(Trim$(tex.Part(datstr, 3, ".")), 4)
        Dim y As Integer = CInt(h)
        If h = "" Then y = Year(Now)
        If y >= 0 And y < 100 Then y = y + CInt(IIf(y < 50, 2000, 1900))
        If y < 1900 Or y > 2100 Then Return NullDate()

        Return DateSerial(y, m, d)
    End Function

    Public Shared Function checkStd(ByVal s As String) As Boolean
        s = s.Trim()
        If s = "" Then Return True
        If s.Contains("-") Then Return False
        Dim d As Integer = CInt(Left$((tex.Part(s, 1, ".")).Trim, 2))
        If d = 0 Then Return False
        Dim m As Integer = CInt(Left$((tex.Part(s, 2, ".")).Trim, 2))
        If m = 0 Then Return False
        Dim h As String = Left$((tex.Part(s, 3, ".")).Trim, 4)
        Dim y As Integer = CInt(h)
        If h = "" Then y = Year(Now)
        If y >= 0 And y < 100 Then y = y + CInt(IIf(y < 50, 2000, 1900))
        If y < 1900 Or y > 2100 Then Return False

        Dim t As Date : t = DateSerial(y, m, d)
        Return (t.Year = y And t.Month = m And t.Day = d)
    End Function

    Public Shared Function checkDat(ByVal s As String) As Boolean
        If s = "" Then Return True
        If s.Contains("-") Then Return checkDIN(s)
        If s.Contains(".") Then Return checkStd(s)
        Return False
    End Function

    Public Overloads Shared Function Equals(ByVal d1 As Date, ByVal d2 As Date) As Boolean
        Return (d1.Date = d2.Date) And (d1.Hour = d2.Hour) And (d1.Minute = d2.Minute)
    End Function

    Public Shared Function DiffSeconds(ByVal dateold As DateTime, ByVal dateNew As DateTime) As Long
        Dim ts As TimeSpan = dateNew.Subtract(dateold)
        Return ts.Seconds
    End Function

    Public Shared Function DiffMinutes(ByVal dateold As DateTime, ByVal dateNew As DateTime) As Long
        Dim ts As TimeSpan = dateNew.Subtract(dateold)
        Return ts.Minutes
    End Function

    Public Shared Function DiffTimeSpan(ByVal dateold As DateTime, ByVal dateNew As DateTime) As TimeSpan
        Return dateNew.Subtract(dateold)
    End Function

    Public Shared Function IIfDat(ByVal Expression As Boolean, ByVal TruePart As Date, ByVal FalsePart As Date) As Date
        If Expression Then Return TruePart
        Return FalsePart
    End Function
End Class

' === zahl ====================================================================
Public Class zahl
    Private Shared k As String 'Komma
    Private Shared p As String 'Punkt

    Public Shared ReadOnly Property Dezimaltrenner() As String
        Get
            init()
            Return k
        End Get
    End Property

    Public Shared ReadOnly Property Tausendertrenner() As String
        Get
            init()
            Return p
        End Get
    End Property

    Public Shared Sub init()
        Static firstTime As Boolean = True
        If Not firstTime Then Exit Sub

        firstTime = False
        Dim kultur As Globalization.CultureInfo = Globalization.CultureInfo.CurrentCulture
        Dim formatInfo As Globalization.NumberFormatInfo = kultur.NumberFormat
        k = formatInfo.NumberDecimalSeparator()
        p = formatInfo.NumberGroupSeparator()
    End Sub

    Public Shared Function checkUDecimal(ByVal s As String) As Boolean 'prüft ob positive Dezimalkommazahl
        init()

        Dim komma As String = "\" & k
        Dim punkt As String = "\" & p
        Dim ziffern As String = "\d+"
        Dim zifferngruppen As String = "(\d{1,3}(" & punkt & "\d{3})*)"
        Dim udecimal As String = "^((" & komma & ziffern & ")|(" & ziffern & komma & ziffern & ")|(" & ziffern & ")|(" & zifferngruppen & komma & ziffern & ")|(" & zifferngruppen & "))$"

        Dim reg As New System.Text.RegularExpressions.Regex(udecimal)
        Return reg.IsMatch(s)
    End Function

    Public Shared Function checkUNumber(ByVal s As String) As Boolean 'prüft ob positive Ganzzahl
        init()

        Dim ziffern As String = "\d+"
        Dim uNumber As String = "^(" & ziffern & ")$"

        Dim reg As New System.Text.RegularExpressions.Regex(uNumber)
        Return reg.IsMatch(s)
    End Function

    Public Shared Function getDecimal(ByVal s As String) As Double '"17,5 kV" bzw. "17.5 kV" -> 17.5 (gemäß Ländereinstellung)
        If s Is Nothing Then s = "" 'Erweiterung, damit ich mich nicht als Aufrufender darum kümmern muss ob eine andere Schnittstelle "" oder Nothing zurückgegeben hat MD 2008-06-30
        init()

        Dim ret As Double
        If p <> "" Then s = s.Replace(p, "")
        If k <> "" Then s = s.Replace(k, ".")
        ret = Val(s)

        Return ret
    End Function

    Public Shared Function getInt(ByVal s As String) As Integer '"2.250 mm" bzw. "2,250 mm" -> 2250 (gemäß Ländereinstellung)
        Dim ret As Integer = CInt(Math.Round(getDecimal(s)))
        Return ret
    End Function

    Public Shared Function getLng(ByVal s As String) As Long '"2.250 mm" bzw. "2,250 mm" -> 2250 (gemäß Ländereinstellung)
        Dim ret As Long = CLng(Math.Round(getDecimal(s)))
        Return ret
    End Function

    Public Shared Function getSapExpFormatLng(ByVal s As String) As Integer '1=1.0000E+00, 10=1.0000E+01
        Dim f As Double = zahl.getEngDecimal(tex.Part(s, 1, "E+"))
        Dim e As Integer = CInt(zahl.getLng(tex.Part(s, 2, "E+")))
        Return CInt(Math.Round(f * 10 ^ e))
    End Function

    Public Shared Function getGerDecimal(ByVal s As String) As Double '"17,5 kV" -> 17.5 (unabhängig von Ländereinstellung)
        If s Is Nothing Then s = "" 'falls Aufrufender Nothing übergibt, kriege ich sonst NullPointerException
        init()

        Dim ret As Double
        s = s.Replace(".", "")
        s = s.Replace(",", ".")
        ret = Val(s)

        Return ret
    End Function

    Public Shared Function getEngDecimal(ByVal s As String) As Double '"17.5 kV" -> 17.5 (unabhängig von Ländereinstellung)
        If s Is Nothing Then s = "" 'falls Aufrufender Nothing übergibt, kriege ich sonst NullPointerException
        init()

        Dim ret As Double
        s = s.Replace(",", "")
        ret = Val(s)

        Return ret
    End Function

    Public Shared Function printGerDecimal(ByVal d As Double) As String '17.5 -> "17,5" (unabhängig von Ländereinstellung)
        init()

        Dim s As String = d.ToString()
        s = s.Replace(k, ",") 'Dezimal-Trenner gegen Komma austauschen

        Return s
    End Function

    Public Shared Function printPreis(ByVal d As Double) As String
        Return d.ToString("#,##0.00")
    End Function

    Public Shared Function printPreisMitEUR(ByVal d As Double) As String
        Return d.ToString("#,##0.00") & " EUR"
    End Function

    Public Shared Function printStueckzahlMitEinheit(ByVal d As Double, ByVal Einheit As String) As String
        Dim ret As String = d.ToString("#,##0")
        If Einheit <> "" Then tex.Cat(ret, Einheit, " ")
        Return ret
    End Function

    Public Shared Function printSapMenge(ByVal d As Double) As String
        Return d.ToString("#.000")
    End Function

    Public Shared Function istKlammermass(ByVal Wert As String) As Boolean
        Return Wert.StartsWith("(") And Wert.EndsWith(")")
    End Function

    Public Shared Function IIfInt(ByVal Expression As Boolean, ByVal TruePart As Integer, ByVal FalsePart As Integer) As Integer
        If Expression Then Return TruePart
        Return FalsePart
    End Function

    Public Shared Function PrintInt10(ByVal Int As Integer, Optional ByVal ReturnIfNull As String = "0000000000") As String
        If Int = 0 Then Return ReturnIfNull
        Return Int.ToString("0000000000")
    End Function

    Public Shared Function PrintInt9(ByVal Int As Integer, Optional ByVal ReturnIfNull As String = "000000000") As String
        If Int = 0 Then Return ReturnIfNull
        Return Int.ToString("000000000")
    End Function

    Public Shared Function PrintInt8(ByVal Int As Integer, Optional ByVal ReturnIfNull As String = "00000000") As String
        If Int = 0 Then Return ReturnIfNull
        Return Int.ToString("00000000")
    End Function

    Public Shared Function PrintInt6(ByVal Int As Integer, Optional ByVal ReturnIfNull As String = "000000") As String
        If Int = 0 Then Return ReturnIfNull
        Return Int.ToString("000000")
    End Function

    Public Shared Function PrintInt4(ByVal Int As Integer, Optional ByVal ReturnIfNull As String = "0000") As String
        If Int = 0 Then Return ReturnIfNull
        Return Int.ToString("0000")
    End Function

    Public Shared Function PrintInt3(ByVal Int As Integer, Optional ByVal ReturnIfNull As String = "000") As String
        If Int = 0 Then Return ReturnIfNull
        Return Int.ToString("000")
    End Function

    Public Shared Function PrintInt2(ByVal Int As Integer, Optional ByVal ReturnIfNull As String = "00") As String
        If Int = 0 Then Return ReturnIfNull
        Return Int.ToString("00")
    End Function

    Public Shared Function PrintInt(ByVal Int As Integer, Optional ByVal ReturnIfNull As String = "") As String
        If Int = 0 Then Return ReturnIfNull
        Return Int.ToString()
    End Function

    Public Shared Function PrintDbl(ByVal dbl As Double, Optional ByVal ReturnIfNull As String = "") As String
        If dbl = 0 Then Return ReturnIfNull
        Return dbl.ToString()
    End Function

End Class

' === key =====================================================================
Public Class key
    Public Shared Function AllowedForDINDate(ByVal c As Char) As Boolean 'lässt nur Zeichen für DIN-Daten zu (Ziffern und Minus)
        Return (Char.IsDigit(c) Or c = "-" Or Char.IsControl(c))
    End Function

    Public Shared Function AllowedForDINDateTime(ByVal c As Char) As Boolean
        Return (Char.IsDigit(c) Or c = "-" Or c = ":" Or Char.IsControl(c))
    End Function

    Public Shared Function AllowedForUFloat(ByVal c As Char) As Boolean
        Return (Char.IsDigit(c) Or {".", ","}.Contains(c) Or Char.IsControl(c))
    End Function

    Public Shared Function AllowedForInteger(ByVal c As Char) As Boolean
        Return (Char.IsDigit(c) Or c = "-" Or Char.IsControl(c))
    End Function

    Public Shared Function AllowedForUInteger(ByVal c As Char) As Boolean
        Return (Char.IsDigit(c) Or Char.IsControl(c))
    End Function

    Public Shared Function AllowedForUAlphaNum(ByVal c As Char) As Boolean
        Return (Char.IsLetterOrDigit(c) Or Char.IsControl(c))
    End Function
End Class

' === clsIniFile ==============================================================
Public Class clsIniFile
    Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function FlushPrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Integer, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer

    Private strFilename As String

    Public Sub New(ByVal Filename As String)
        strFilename = Filename
    End Sub

    Public Function GetString(ByVal Section As String, ByVal Key As String, Optional ByVal [Default] As String = "") As String
        Dim objResult As New System.Text.StringBuilder(256)
        Dim intCharCount As Integer = GetPrivateProfileString(Section, Key, [Default], objResult, objResult.Capacity, strFilename)
        If intCharCount > 0 Then Return Left(objResult.ToString, intCharCount)
        Return ""
    End Function

    Public Function GetInteger(ByVal Section As String, ByVal Key As String, Optional ByVal [Default] As Integer = 0) As Integer
        Return GetPrivateProfileInt(Section, Key, [Default], strFilename)
    End Function

    Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal Value As String)
        WritePrivateProfileString(Section, Key, Value, strFilename)
        Flush()
    End Sub

    Public Sub WriteInteger(ByVal Section As String, ByVal Key As String, ByVal Value As Integer)
        WriteString(Section, Key, Value.ToString())
        Flush()
    End Sub

    Private Sub Flush()
        FlushPrivateProfileString(0, 0, 0, strFilename)
    End Sub
End Class

' === dat =====================================================================
Public Class menge
    Public Shared Function IstTeilmenge(ByVal m1 As String, ByVal m2 As String) As Boolean
        Dim i As Integer
        Dim h As String

        Dim m1n As String : m1n = Replace(m1, " ", "")
        Dim m2n As String : m2n = Replace(m2, " ", "")

        If m2n = "" Then IstTeilmenge = False : Exit Function
        m2n = "," & m2n & ","

        For i = 1 To tex.PartCount(m1, ",")
            h = Trim$(tex.Part(m1, i, ","))
            If InStr(m2n, "," & h & ",") = 0 Then IstTeilmenge = False : Exit Function
        Next i

        IstTeilmenge = True
    End Function

    Public Shared Function Schnittmenge(ByVal m1 As String, ByVal m2 As String) As String
        Dim i As Integer
        Dim h As String
        Dim ret As String : ret = ""

        Dim m1n As String : m1n = Replace(m1, " ", "")
        Dim m2n As String : m2n = Replace(m2, " ", "")

        For i = 1 To tex.PartCount(m1n, ",")
            h = tex.Part(m1n, i, ",")
            If InStr("," & m2n & ",", "," & h & ",") > 0 Then
                tex.Cat(ret, h, ",")
            End If
        Next i

        Schnittmenge = ret
    End Function

    Public Shared Function MengePlus(ByVal m1 As String, ByVal m2 As String) As String
        Dim i As Integer
        Dim h As String
        Dim ret As String : ret = ""

        Dim m1n As String : m1n = Replace(m1, " ", "")
        Dim m2n As String : m2n = Replace(m2, " ", "")

        For i = 1 To tex.PartCount(m1n, ",")
            h = tex.Part(m1n, i, ",")
            If InStr("," & ret & ",", "," & h & ",") = 0 Then tex.Cat(ret, h, ",")
        Next i
        For i = 1 To tex.PartCount(m2n, ",")
            h = Trim$(tex.Part(m2n, i, ","))
            If InStr("," & ret & ",", "," & h & ",") = 0 Then tex.Cat(ret, h, ",")
        Next i

        MengePlus = ret
    End Function

    Public Shared Function MengeMinus(ByVal m1 As String, ByVal m2 As String) As String
        Dim ret As String : ret = "" 'm1 - m2

        Dim m1n As String : m1n = Replace(m1, " ", "")
        Dim m2n As String : m2n = Replace(m2, " ", "")

        Dim h As String
        Dim i As Integer
        For i = 1 To tex.PartCount(m1n, ",")
            h = tex.Part(m1n, i, ",")
            If InStr("," & m2n & ",", "," & h & ",") = 0 Then
                tex.Cat(ret, h, ",")
            End If
        Next i

        MengeMinus = ret
    End Function

    Public Shared Function enthaelt(ByVal m1 As String, ByVal x As String) As Boolean
        If m1 Is Nothing Then m1 = ""
        Dim m1n As String : m1n = Obj2Str(Replace(m1, " ", ""))
        enthaelt = (InStr("," & m1n & ",", "," & x & ",") > 0)
    End Function

    Public Shared Function Anzahl(ByVal m1 As String) As Integer
        Dim m1n As String : m1n = Replace(m1, " ", "")
        Return tex.PartCount(m1n, ",")
    End Function

    Public Shared Function Anzahl(ByVal m1 As String, ByVal x As Integer) As Integer
        Dim ret As Integer = 0

        Dim m1n As String : m1n = Replace(m1, " ", "")

        Dim i As Integer
        For i = 1 To tex.PartCount(m1n, ",")
            If zahl.getLng(tex.Part(m1n, i, ",")) = x Then ret += 1
        Next

        Return ret
    End Function

    Public Shared Function SortierteListe(ByVal m As String) As String
        Dim ret As String = ""

        Dim l As New ArrayList
        Dim i As Integer
        For i = 1 To tex.PartCount(m, ",")
            Dim z As Integer = zahl.getLng(tex.Part(m, i, ","))
            l.Add(z)
        Next

        l.Sort()
        For i = 0 To l.Count - 1
            tex.Cat(ret, CType(l.Item(i), Integer).ToString(), ",")
        Next

        Return ret
    End Function
End Class

'=== String ===================================================================
Public Class tex

    Public Shared Function IIfStr(ByVal Expression As Boolean, ByVal TruePart As String, ByVal FalsePart As String) As String
        If Expression Then Return TruePart
        Return FalsePart
    End Function

    Private Shared Function ContainsChineseChars(ByVal text As String) As Boolean
        For Each digit As Char In text
            If Convert.ToInt64(digit) > 256 Then Return True
        Next
        Return False
    End Function

    Public Shared Function CheckConvertChinese2ISO(ByVal text As String) As String
        Dim str As String = ConvertChinese2ISO(text)
        If ContainsChineseChars(str) Then Return str
        Return text
    End Function

    Public Shared Function CheckConvertISO2Chinese(ByVal text As String) As String
        Dim str As String = ConvertISO2Chinese(text)
        If ContainsChineseChars(str) Then Return str
        Return text
    End Function

    Public Shared Function ConvertChinese2ISO(ByVal text As String) As String
        If Not ContainsChineseChars(text) Then Return text

        Dim enciso As System.Text.Encoding = System.Text.Encoding.GetEncoding("ISO-8859-1") 'ISO-8859-1
        Dim encchin As System.Text.Encoding = System.Text.Encoding.GetEncoding("GB2312") 'GB2312, vereinf. Chinesisch

        Return enciso.GetString(encchin.GetBytes(text))
    End Function

    Public Shared Function ConvertISO2Chinese(ByVal text As String) As String
        If Not ContainsChineseChars(text) Then Return text

        Dim enciso As System.Text.Encoding = System.Text.Encoding.GetEncoding("ISO-8859-1") 'ISO-8859-1
        Dim encchin As System.Text.Encoding = System.Text.Encoding.GetEncoding("GB2312") 'GB2312, vereinf. Chinesisch

        Return encchin.GetString(enciso.GetBytes(text))
    End Function

    Public Shared Function Print(ByVal s As String, Optional ByVal ReturnIfNull As String = "") As String
        If s = "" Then Return ReturnIfNull
        Return s
    End Function

    Public Shared Sub Cat(ByRef s As String, ByVal wort As String, ByVal trenn As String)
        If s Is Nothing Then s = "" 'Sicherheitsnetz
        If wort Is Nothing Then wort = ""
        If trenn Is Nothing Then trenn = ""

        Dim builder As New System.Text.StringBuilder(s, s.Length + wort.Length + trenn.Length)
        If s.Length <> 0 Then builder.Append(trenn)
        builder.Append(wort)
        s = builder.ToString
    End Sub

    Public Shared Function PartCount(ByVal s As String, ByVal t As String) As Integer
        If s = "" Then Return 0

        s = Replace(s, t, ControlChars.NullChar)
        Dim parts() As String = s.Split(ControlChars.NullChar)

        Return parts.Length
    End Function

    Public Shared Function Part(ByVal v As Object, ByVal p As Integer, ByVal t As String) As String
        Dim s As String = Obj2Str(v)
        If s = "" Then Return ""

        s = Replace(s, t, ControlChars.NullChar)
        Dim parts() As String = s.Split(ControlChars.NullChar)

        If p >= 1 And p <= parts.Length Then Return parts(p - 1)

        Return ""
    End Function

    Public Shared Function umlaute_raus(ByVal s As String) As String
        Dim ret As String = s
        If s.IndexOf("ä") >= 0 Then ret = ret.Replace("ä", "ae")
        If s.IndexOf("Ä") >= 0 Then ret = ret.Replace("Ä", "Ae")
        If s.IndexOf("ö") >= 0 Then ret = ret.Replace("ö", "oe")
        If s.IndexOf("Ö") >= 0 Then ret = ret.Replace("Ö", "Oe")
        If s.IndexOf("ü") >= 0 Then ret = ret.Replace("ü", "ue")
        If s.IndexOf("Ü") >= 0 Then ret = ret.Replace("Ü", "Ue")
        If s.IndexOf("ß") >= 0 Then ret = ret.Replace("ß", "ss")
        umlaute_raus = ret
    End Function

    Public Shared Function GrossKleinSchreibung(ByVal s As String) As String
        Dim ret As String : ret = ""
        If s <> "" Then
            Dim i As Integer
            Dim h As String
            Dim gross As Boolean : gross = True

            For i = 1 To Len(s)
                h = Mid$(s, i, 1)
                If h = " " Or h = "-" Then
                    ret = ret & h
                    gross = True
                Else
                    If gross Then
                        ret = ret & UCase$(h)
                        gross = False
                    Else
                        ret = ret & LCase$(h)
                    End If
                End If
            Next i
        End If
        Return ret
    End Function

    Public Shared Function PrintFixLength(ByVal s As String, ByVal i As Integer) As String
        Return Left$(s + Space(i), i)
    End Function

End Class

' === Datei ===================================================================
Public Class datei

    Public Shared Function Delete(ByVal f As String) As Boolean
        Try
            File.Delete(f)
        Catch ex As Exception
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileDeletingFile_PleaseDeleteFileManually.Replace("{0}", f) & vbCrLf & vbCrLf & PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Copy(ByVal souPath As String, ByVal desPath As String, Optional ByVal overwrite As Boolean = False, Optional ByVal Silent As Boolean = False) As Boolean
        If souPath = desPath Then Return True 'Datei auf sich selbst kopieren geht zwar nicht, da in dem Fall aber nichts zu tun ist melde Erfolg
        Try
            If clsDirectory.IsDirectory(desPath) Then
                If Not IO.Directory.Exists(desPath) Then
                    clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileCopyingFile_DestinationDirectory_NotAvailableOrUserIsNotAuthorized.Replace("{0}", souPath).Replace("{1}", desPath))
                    Return False
                End If
                desPath = IO.Path.Combine(desPath, IO.Path.GetFileName(souPath))
            Else
                If Not clsDirectory.Exists(IO.Path.GetDirectoryName(desPath)) Then
                    clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileCopyingFile_DestinationDirectory_NotAvailableOrUserIsNotAuthorized.Replace("{0}", souPath).Replace("{1}", IO.Path.GetDirectoryName(desPath)))
                    Return False
                End If
            End If

            If Not overwrite AndAlso File.Exists(desPath) Then
                clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileCopyingFile_FileIsAlreadyAvailableOverwritingNotIntended.Replace("{0}", desPath))
                Return False
            End If
            IO.File.Copy(souPath, desPath, overwrite)
        Catch secex As System.Security.SecurityException
            clsShow.InternalError(My.Resources.resGlobal.MsgErrorWhileCopyingFile_DestinationDirectory_NotAvailableOrUserIsNotAuthorized.Replace("{0}", souPath).Replace("{1}", desPath))
            Return False
        Catch ex As Exception
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileCopyingFile_DestinationDirectory_NotAvailableOrUserIsNotAuthorized.Replace("{0}", souPath).Replace("{1}", desPath) & vbCrLf & PrintException(ex))
            Return False 'Exit Function zeigt ein zu großes Vertrauen in den Compiler, den Rückgabewert richtig initialisiert zu haben
        End Try
        Return True
    End Function

    Public Shared Function Time(ByVal pfad As String) As Date
        Dim ret As Date = dat.NullDate()
        Try
            ret = IO.File.GetLastWriteTime(pfad)
        Catch ex As Exception
            clsShow.InternalError(My.Resources.resGlobal.MsgErrorWhileAccessingLastWriteTime4File_.Replace("{0}", pfad) & vbCrLf & PrintException(ex))
        End Try
        Return ret
    End Function

    Public Shared Function Move(ByVal sour As String, ByVal dest As String, Optional ByVal Override As Boolean = False) As Boolean
        Try
            'File.Move(sour, dest) 'sonst werden Berechtigungen des Netzlaufwerks mit verschoben statt die des Zieles zu übernehmen
            If Copy(sour, dest, Override) Then Delete(sour)
        Catch ex As Exception
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileMovingFile_To_.Replace("{0}", sour).Replace("{1}", dest) & vbCrLf & PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Find(ByVal verz As String, ByVal suchMuster As String) As String
        Dim s() As String = IO.Directory.GetFiles(verz, suchMuster, SearchOption.TopDirectoryOnly)
        If s.GetLength(0) > 0 Then
            Return s(0)
        End If
        Return ""
    End Function

    Public Shared Function Update(ByVal srcFile As String, ByVal dstFile As String) As Boolean
        'wenn Zieldatei nicht vorhanden oder veraltet dann überschreiben
        Dim ret As Boolean = True
        If IO.File.Exists(dstFile) Then
            Dim fiaktuell As New IO.FileInfo(srcFile)
            Dim fiVorhanden As New IO.FileInfo(dstFile)

            If (fiaktuell.LastWriteTime.Ticks > fiVorhanden.LastWriteTime.Ticks) Then
                If Not Copy(srcFile, dstFile, True) Then ret = False
            End If
            fiaktuell = Nothing
            fiVorhanden = Nothing
        Else
            If Not Copy(srcFile, dstFile, True) Then ret = False
        End If

        Return ret
    End Function

    Public Shared Function Exists(ByVal path As String) As Boolean
        Dim ret As Boolean
        Try
            ret = IO.File.Exists(path)
        Catch ex As Exception
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileSearchingFile_.Replace("{0}", path) & vbCrLf & PrintException(ex))
            ret = False
        End Try
        Return ret
    End Function

    Public Function CopyMultiple(ByVal srcPfade As String, ByVal dstVerz As String) As Boolean

        Dim ok As Boolean = True

        Dim anz As Integer = tex.PartCount(srcPfade, ",")
        For i As Integer = 1 To anz ' sonst muss bei JEDEM Schleifendurchlauf erneut tex.PartCount aufgerufen werden
            Dim srcPfad As String = tex.Part(srcPfade, i, ",")
            Dim srcDatei As String = IO.Path.GetFileName(srcPfad)
            Dim dstPfad As String = IO.Path.Combine(dstVerz, srcDatei)
            If Not Copy(srcPfad, dstPfad, True) Then ok = False
        Next

        Return ok
    End Function

    Public Shared Function readBytes(ByVal Pfad As String) As Byte()
        If Not datei.Exists(Pfad) Then Return Nothing
        Return File.ReadAllBytes(Pfad)

    End Function

    Public Shared Function writeBytes(ByVal Pfad As String, ByVal Bytes As Byte()) As Boolean
        If IsNothing(Bytes) Then Return False

        Try
            File.WriteAllBytes(Pfad, Bytes)
            Return True
        Catch ex As Exception
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgCantWriteFile_.Replace("{0}", Pfad) & vbCrLf & PrintException(ex))
            Return False
        End Try
    End Function

    Public Shared Function Print(ByVal Pfad As String) As Boolean
        If Not datei.Exists(Pfad) Then Return False

        Dim myProcess As Process
        Dim psi As ProcessStartInfo = New ProcessStartInfo()

        Try
            psi.Verb = "print"
            psi.UseShellExecute = True
            psi.WindowStyle = ProcessWindowStyle.Hidden
            psi.FileName = Pfad
            myProcess = Process.Start(psi)
        Catch ex As Exception
            clsShow.ErrorMsg(PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Open(ByVal Pfad As String) As Boolean
        If Not datei.Exists(Pfad) Then Return False

        Dim myProcess As Process
        Try
            myProcess = Process.Start(Pfad)
        Catch ex As Exception
            clsShow.ErrorMsg(PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Shared Function CheckName4FileName(ByVal FileName As String) As String
        Dim ret As String = FileName.Replace("/", "_")
        ret = ret.Replace("\", "_")
        ret = ret.Replace(":", "_")
        ret = ret.Replace("*", "_")
        ret = ret.Replace("?", "_")
        ret = ret.Replace(">", "_")
        ret = ret.Replace("<", "_")
        ret = ret.Replace("|", "_")
        ret = ret.Replace(Chr(13), "_")
        Return umlaute_raus(ret)
    End Function
End Class

' === clsDirectory ============================================================
Public Class clsDirectory
    Public Shared Function IsDirectory(ByVal pfad As String) As Boolean
        pfad = pfad.TrimEnd("\"c)
        If System.IO.Path.HasExtension(pfad) Then Return False
        If Not System.IO.Path.IsPathRooted(pfad) Then Return False ' Derzeit nur für absolute Pfade
        Dim b As String = System.IO.Path.GetFullPath(pfad)
        If b <> pfad Then Return False ' Es sollte der selbe/gleiche Pfad zurück gegeben werden, wenn dies nicht der Fall ist, dann war schon der Input Parameter nicht korrekt konstruiert.
        Return True
    End Function

    Public Shared Function Move(ByVal Quelle As String, ByVal Ziel As String) As Boolean
        Quelle = Quelle.TrimEnd("\"c)
        Ziel = Ziel.TrimEnd("\"c)

        Dim InhaltVerzeichnis As New clsVerzeichnisInhalt
        InhaltVerzeichnis.ReadDatenKomplett(Quelle, True) ', frm)
        'If Not IsNothing(frm) AndAlso frm.Abbrechen Then Return False


        Dim InhaltVerzeichnisTauschen As New clsVerzeichnisInhalt
        InhaltVerzeichnisTauschen.ReadDatenKomplett(Quelle, True) ', frm)
        'If Not IsNothing(frm) AndAlso frm.Abbrechen Then Return False

        For i As Integer = 0 To InhaltVerzeichnisTauschen.Verzeichnisse.Count - 1
            'If Not IsNothing(frm) Then frm.ShowInfo(My.Resources.resGlobal.MsgMovingFile, -1)

            Dim Pfad As String = InhaltVerzeichnisTauschen.Verzeichnisse.Item(i)

            InhaltVerzeichnisTauschen.Verzeichnisse.Item(i) = Pfad.Replace(Quelle, Ziel)
            If Not clsDirectory.Make(InhaltVerzeichnisTauschen.Verzeichnisse.Item(i)) Then Return False
        Next

        For i As Integer = 0 To InhaltVerzeichnis.Dateien.Count - 1
            'If Not IsNothing(frm) Then frm.ShowInfo(My.Resources.resGlobal.MsgMovingFile, -1)


            Dim Pfad As String = InhaltVerzeichnisTauschen.Dateien.Item(i)
            Pfad = Pfad.Replace(Quelle, Ziel)

            If Not datei.Move(InhaltVerzeichnis.Dateien.Item(i), Pfad, True) Then Return False
        Next

        For i As Integer = 0 To InhaltVerzeichnis.Verzeichnisse.Count - 1
            'If Not IsNothing(frm) Then frm.ShowInfo(My.Resources.resGlobal.MsgMovingFile, -1)

            If Not clsDirectory.Delete(InhaltVerzeichnis.Verzeichnisse.Item(i)) Then Return False
        Next

        Return True
    End Function

    Public Shared Function Open(ByVal verz As String) As Boolean
        Try
            Shell("explorer.exe """ & verz & """", vbNormalFocus)
        Catch ex As Exception
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileOpeningDirectory_.Replace("{0}", verz) & vbCrLf & PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Delete(ByVal verz As String, Optional ByVal withsubs As Boolean = True, Optional ByVal ShowError As Boolean = True) As Boolean
        If Not IO.Directory.Exists(verz) Then Return True ' wenn es gar nicht existiert, ist das Ergebnis des Löschens das gleiche
        Try
            IO.Directory.Delete(verz, withsubs)
        Catch ex As Exception
            clsShow.InternalError(My.Resources.resGlobal.MsgErrorWhileDeletingDirectory_.Replace("{0}", verz) & vbCrLf & PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Make(ByVal verz As String) As Boolean
        Try
            IO.Directory.CreateDirectory(verz)
        Catch ex As Exception
            clsShow.InternalError(My.Resources.resGlobal.MsgErrorWhileCreatingDirectory_.Replace("{0}", verz) & vbCrLf & PrintException(ex))
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Exists(ByVal verz As String) As Boolean
        Return IO.Directory.Exists(verz)
    End Function

    Public Shared Function Exists(ByVal rootDir As String, ByVal dirPattern As String) As Boolean
        Try
            Dim DirList As String() = IO.Directory.GetDirectories(rootDir, dirPattern)
            Return DirList.Length() > 0
        Catch ex As DirectoryNotFoundException
            Return False
        Catch e As Exception
            clsShow.ErrorMsg(My.Resources.resGlobal.MsgErrorWhileCheckingAvailabilityOfDirectory4_WithSearchPattern_.Replace("{0}", rootDir).Replace("{1}", dirPattern) & vbCrLf & PrintException(e))
            Return False
        End Try
    End Function

    Public Shared Sub Copy(ByVal SourcePath As String, ByVal DestPath As String, Optional ByVal Overwrite As Boolean = False)
        Dim SourceDir As New DirectoryInfo(SourcePath)
        Dim DestDir As New DirectoryInfo(DestPath)

        ' the source directory must exist, otherwise throw an exception
        If SourceDir.Exists Then
            ' if destination SubDir's parent SubDir does not exist throw an exception
            If Not DestDir.Parent.Exists Then
                Throw New DirectoryNotFoundException _
                    (My.Resources.resGlobal.MsgTargetDirectoryDoesNotExist_.Replace("{0}", DestDir.Parent.FullName))
            End If

            If Not DestDir.Exists Then
                DestDir.Create()
            End If

            ' copy all the files of the current directory
            Dim ChildFile As FileInfo
            For Each ChildFile In SourceDir.GetFiles()
                If Overwrite Then
                    ChildFile.CopyTo(Path.Combine(DestDir.FullName, ChildFile.Name), True)
                Else
                    ' if Overwrite = false, copy the file only if it does not exist
                    ' this is done to avoid an IOException if a file already exists
                    ' this way the other files can be copied anyway...
                    If Not datei.Exists(Path.Combine(DestDir.FullName, ChildFile.Name)) Then
                        ChildFile.CopyTo(Path.Combine(DestDir.FullName, ChildFile.Name), False)
                    End If
                End If
            Next

            ' copy all the sub-directories by recursively calling this same routine
            Dim SubDir As DirectoryInfo
            For Each SubDir In SourceDir.GetDirectories()
                Copy(SubDir.FullName, Path.Combine(DestDir.FullName, _
                    SubDir.Name), Overwrite)
            Next
        Else
            Throw New DirectoryNotFoundException(My.Resources.resGlobal.MsgSourceDirectoryDoesNotExist_.Replace("{0}", SourceDir.FullName))
        End If
    End Sub

End Class

Public Class clsVerzeichnisInhalt
    Public Dateien As New List(Of String)
    Public Verzeichnisse As New List(Of String)
    Public DatenInfos As New List(Of FileInfo)
    Public VerzeichnisseInfo As New List(Of DirectoryInfo)
    Public Verzeichnis As String

    Public Sub New()
        ' nix
    End Sub

    Public Sub ReadDatenKomplett(ByVal Pfad As String, ByVal LeseRekursiv As Boolean)
        Verzeichnis = Pfad
        Dateien.Clear()
        Verzeichnisse.Clear()

        Dim tempVerzeichnis As List(Of String) = GetDirectories(Pfad)

        Dateien.AddRange(GetFiles(Pfad))

        If LeseRekursiv Then
            For i As Integer = 0 To tempVerzeichnis.Count - 1
                'If Not IsNothing(frm) AndAlso frm.Abbrechen Then Exit Sub
                Application.DoEvents()
                ReadDatenKomplettRekursiv(tempVerzeichnis.Item(i))
            Next
        End If

        Verzeichnisse.AddRange(tempVerzeichnis)

        For i As Integer = 0 To Dateien.Count - 1
            'If Not IsNothing(frm) AndAlso frm.Abbrechen Then Exit Sub
            DatenInfos.Add(New FileInfo(Dateien(i)))
        Next

        For i As Integer = 0 To Verzeichnisse.Count - 1
            'If Not IsNothing(frm) AndAlso frm.Abbrechen Then Exit Sub
            VerzeichnisseInfo.Add(New DirectoryInfo(Verzeichnisse(i)))
        Next
    End Sub

    Private Sub ReadDatenKomplettRekursiv(ByVal Pfad As String)
        Application.DoEvents()
        Dim tempVerzeichnis As List(Of String) = GetDirectories(Pfad)
        Dim tempDateien As List(Of String) = GetFiles(Pfad)

        Verzeichnisse.AddRange(tempVerzeichnis)
        Dateien.AddRange(tempDateien)

        For i As Integer = 0 To tempVerzeichnis.Count - 1
            ReadDatenKomplettRekursiv(tempVerzeichnis.Item(i))
        Next
    End Sub
End Class

Public Class clsImage

    Public Overloads Shared Function ResizeImageAndSave(ByVal souPath As String, ByVal desDir As String, ByVal sizeFactorInPercent As Integer) As String
        If desDir = "" Then desDir = IO.Path.GetDirectoryName(souPath)
        Dim FileName As String = IO.Path.GetFileName(souPath)
        Dim desPath As String = IO.Path.Combine(desDir, "25" & FileName)

        Using bmp As New Bitmap(souPath)
            Using newBmp As Bitmap = ResizeImage(bmp, sizeFactorInPercent)
                newBmp.Save(desPath, Imaging.ImageFormat.Jpeg)
                Application.DoEvents()
            End Using
        End Using

        Return desPath
    End Function

    Public Shared Function ResizeImage(ByVal img As Image, ByVal sizeFactorInPercent As Integer) As Bitmap
        Dim oldW As Integer = img.Width
        Dim oldH As Integer = img.Height
        Dim newW As Integer = CInt((oldW / 100) * sizeFactorInPercent)
        Dim newH As Integer = CInt((oldH / 100) * sizeFactorInPercent)

        Dim newBmp As New Bitmap(newW, newH)

        Using g As Graphics = Graphics.FromImage(newBmp)
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.DrawImage(img, 0, 0, newW, newH)
        End Using

        Return newBmp
    End Function

End Class

Public Class clsShow
    Public Shared Sub Message(ByVal txt As String, Optional ByVal titel As String = "")
        MsgBox(txt, MsgBoxStyle.Information, IIf(titel = "", My.Resources.resGlobal.TextInformation, titel))
    End Sub

    Public Shared Sub InternalMessage(ByVal txt As String, Optional ByVal titel As String = "")
        MsgBox(txt, MsgBoxStyle.Information, IIf(titel = "", My.Resources.resGlobal.TextInformation, titel))
    End Sub

    Public Shared Sub ErrorMsg(ByVal s As String)
        'If session.SILENT Then
        '    clsLog.LogLineError(My.Resources.resGlobal.TextError, s)
        'Else
        MsgBox(s, MsgBoxStyle.Exclamation, My.Resources.resGlobal.TextError)
        'End If
    End Sub

    Public Shared Function MsgYesNoCancel(ByVal s As String, Optional ByVal titel As String = "") As TriState
        Dim r As MsgBoxResult = (MsgBox(s, CType(MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, MsgBoxStyle), IIf(titel = "", My.Resources.resGlobal.TextQuestion, titel)))
        : Return TriState.True
        If r = MsgBoxResult.Yes Then
        ElseIf r = MsgBoxResult.No Then
            : Return TriState.False
        Else : Return TriState.UseDefault
        End If
    End Function

    Public Shared Sub InternalError(ByVal txt As String)
        'If session.SILENT Then
        '    clsLog.LogLineError(My.Resources.resGlobal.TextError, txt)
        'Else
        MsgBox(txt, MsgBoxStyle.Exclamation, My.Resources.resGlobal.TextInternalError)
        'End If
    End Sub

    Public Shared Function Retry(ByVal s As String, Optional ByVal t As String = "") As Boolean
        Return (MsgBox(s, CType(MsgBoxStyle.Question + MsgBoxStyle.RetryCancel, MsgBoxStyle), IIf(t = "", My.Resources.resGlobal.TextRetry, t)) = MsgBoxResult.Retry)
    End Function

    Public Shared Function Question(ByVal s As String, Optional ByVal t As String = "") As Boolean
        Return (MsgBox(s, CType(MsgBoxStyle.Question + MsgBoxStyle.YesNo, MsgBoxStyle), IIf(t = "", My.Resources.resGlobal.TextQuestion, t)) = MsgBoxResult.Yes)
    End Function

    Public Shared Sub InternalError(ByVal txt As String, ByVal sendToAdmin As Boolean)
        '  If clsOutlook.OutlookIsOpen(False) Then
        'clsOutlook.CreateEMail(session.IniErrorMailRecipient, My.Resources.Res_Andere_Texte.Fehler_EmailTitel, My.Resources.Res_Andere_Texte.Fehler_EmailText & vbCrLf & txt)
        'MsgBox(txt, MsgBoxStyle.Critical, My.Resources.Res_Andere_Texte.Fehler_EmailTitel)
        'Else
        '    MsgBox(txt & vbCrLf & My.Resources.Res_Andere_Texte.Fehler_EmailScreenshot, MsgBoxStyle.Critical, My.Resources.Res_Andere_Texte.Fehler_EmailTitel)
        'End If
    End Sub
End Class