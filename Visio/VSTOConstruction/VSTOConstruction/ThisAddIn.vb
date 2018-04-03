Public Class ThisAddIn
    Dim xcoord As Double = 2
    Dim ycoord As Double = 2

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        Dim commandBars As Office.CommandBars
        Dim commandBar As Office.CommandBar
        Dim runStoreReport As Office.CommandBarButton
        commandBars = CType(Application.CommandBars, Microsoft.Office.Core.CommandBars)
        commandBar = commandBars.Add("VSTOAddinToolbar", Office.MsoBarPosition.msoBarTop, , True)
        ' Set the context when the toolbar is visible.
        commandBar.Context = Visio.VisUIObjSets.visUIObjSetDrawing & "*"

        ' Add a button with an icon that looks like a report.
        runStoreReport = CType(commandBar.Controls.Add( _
            Office.MsoControlType.msoControlButton),  _
            Microsoft.Office.Core.CommandBarButton)
        runStoreReport.Tag = "Store Report"

        AddHandler runStoreReport.Click, AddressOf VisualizeSales

    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Private Sub VisualizeSales()
        Dim connectionString As String
        connectionString = "Data Source=LICH\SQL2K5;Initial Catalog=AdventureWorks;Integrated Security=True"
        Dim commandString As String
        commandString = "Select  Name, SquareFeet, AnnualRevenue from vStoreWithDemographics WHERE EmailPromotion = 1"
        Dim currentDocument As Visio.Document
        currentDocument = Application.ActiveDocument
        Dim storeData As Visio.DataRecordset
        storeData = currentDocument.DataRecordsets.Add(connectionString, commandString, 0, "Store Data")

        Dim currentRows As Array
        currentRows = storeData.GetDataRowIDs("")

        Dim currentDataRow As Object

        For counter As Integer = 1 To currentRows.Length
            currentDataRow = storeData.GetRowData(counter)
            DrawStore(CDbl(currentDataRow(1)), CDbl(currentDataRow(2)), counter)
            xcoord = xcoord + 2
        Next
        Application.ActiveWindow.Windows.ItemFromID(Visio.VisWinTypes.visWinIDExternalData).Visible = True
    End Sub

    Private Sub DrawStore(ByVal sqFoot As Double, ByVal sales As Double, ByVal rowId As Integer)
        Dim currentPage As Visio.Page = Application.ActivePage
        Dim currentDocument As Visio.Document = Application.ActiveDocument
        Dim storeShape As Visio.Shape

        Dim newx As Integer = xcoord + (10 * (sales / 1000000))
        Dim newy As Integer = ycoord + (10 * (sales / 1000000))

        storeShape = currentPage.DrawRectangle(xcoord, ycoord, newx, newy)
        storeShape.LinkToData(1, rowId, True)

    End Sub
End Class
