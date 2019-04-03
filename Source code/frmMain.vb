
Imports Test.Common.Interfaces
Imports Test.Common.Model
Imports Test.Common.Services

Public Class frmMain

    Private ResultFunction As FileReaderResponse

    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click
        If fbdInputPath.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            txtInputPath.Text = VerifyPath(fbdInputPath.SelectedPath)
        End If
    End Sub

    Private Function VerifyPath(selectedPath As String) As String
        Dim count = selectedPath.Count(Function(x) Char.IsDigit(x))
        If count < 10 Then
            Throw New Exception("Invalid path!")
        Else

            Dim instance = New FileReaderService()
            Dim result = instance.ProcceedFiles(selectedPath)
            ResultFunction = result
            dgvSummary.DataSource = ResultFunction.GridSource

        End If
        Return selectedPath
    End Function

    Private Sub btnImport_Click(sender As System.Object, e As System.EventArgs) Handles btnImport.Click
#Disable Warning BC42025
        Dim ProjectEntity As List(Of clsSQLProject) = New clsSQLProject().GetAll()

        If ProjectEntity.Any(Function(s As clsSQLProject) s.OrderNo = ResultFunction.ProjectOrderNo) Then

            MessageBox.Show("Db already contains record clsSQLProject with declared parameters! Skip")

        Else

            Dim B As clsSQLProject = New clsSQLProject()
            B.OrderNo = ResultFunction.ProjectOrderNo
            B.ProductID = 1
            B.EntryDate = Date.Now
            B.Update()

        End If

        MessageBox.Show(ResultFunction.ToString())

        Dim ConnLocationEntity As List(Of clsSQLConnLocation) = New clsSQLConnLocation().GetAll()

        ResultFunction.FileModels.ForEach(Function(item)

                                              If Not ConnLocationEntity.Any(Function(s As clsSQLConnLocation) s.ShortName = item.WireDef.ConnectionLocation) Then

                                                  Dim connLocation As clsSQLConnLocation = New clsSQLConnLocation()
                                                  connLocation.ShortName = item.WireDef.ConnectionLocation
                                                  connLocation.Update()

                                              End If
#Disable Warning BC42105
                                          End Function)
#Enable Warning BC42105

        Dim WireEntity As List(Of clsSQLWire) = New clsSQLWire().GetAll()

        ResultFunction.FileModels.ForEach(Function(item)

                                              If Not WireEntity.Any(Function(s As clsSQLWire) s.TypeName = item.WireDef.WireTypeName And
                                                                        s.CrossSection = item.WireDef.WireCrossSection And
                                                                        s.Color = item.WireDef.WireColor) Then

                                                  Dim wire As clsSQLWire = New clsSQLWire()

                                                  wire.Color = item.WireDef.WireColor
                                                  wire.CrossSection = item.WireDef.WireCrossSection
                                                  wire.TypeName = item.WireDef.WireTypeName

                                                  wire.Update()

                                              End If
#Disable Warning BC42105
                                          End Function)
#Enable Warning BC42105

        Dim projectId As Int32 = New clsSQLProject().GetAll().Single(Function(s As clsSQLProject) s.OrderNo = ResultFunction.ProjectOrderNo).ID

        ResultFunction.FileModels.ForEach(Function(item)


                                              Dim wireId As Int32 = New clsSQLWire().GetAll().Single(Function(s As clsSQLWire) s.TypeName = item.WireDef.WireTypeName And s.Color = item.WireDef.WireColor And s.CrossSection = item.WireDef.WireCrossSection).ID
                                              Dim connLocationId As Int32 = New clsSQLConnLocation().GetAll().Single(Function(s As clsSQLConnLocation) s.ShortName = item.WireDef.ConnectionLocation).ID

                                              Dim connQty As clsSQLConnQuantity = New clsSQLConnQuantity()

                                              connQty.ConnLocationID = connLocationId
                                              connQty.WireID = wireId
                                              connQty.ProjectID = projectId
                                              connQty.Quantity = item.CountOfElements

                                              connQty.Update()

#Disable Warning BC42105
                                          End Function)
#Enable Warning BC42105

#Enable Warning BC42025
    End Sub
End Class
