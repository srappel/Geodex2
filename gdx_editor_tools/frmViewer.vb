Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports System.Windows.Forms

Public Class frmViewer
    Dim exception As Boolean

    Private _application As IApplication
    Public Property ArcMapApplication() As IApplication
        Get
            Return _application
        End Get
        Set(ByVal value As IApplication)
            _application = value
        End Set
    End Property

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        ' Enable key preview.
        Me.KeyPreview = True
    End Sub

    Private Sub TabProcessorDialog_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        ' No need to process keypress if form is modal.
        If Me.Modal Then Return

        ' Process the tab key.
        If e.KeyCode = Keys.Tab Then
            If e.Modifiers = Keys.Shift Then
                ' Reverse tabbing if Shift key is held down.
                Me.ProcessTabKey(False)
            Else
                ' Forward tabbing if Shift key is not held down.
                Me.ProcessTabKey(True)
            End If
        End If
    End Sub

    'On Load
    'If the geodex layer isn't there, don't load the form.
    'if it is, populate the files drop down using populatefiles()
    Public Sub frmViewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim pLayer As ILayer = GDXTools.getInstance.GetGeodexLayer
        If pLayer Is Nothing Then
            MsgBox("Layer Geodex is not loaded, use the load data button to load the layer.")
            exception = True

            Me.Close()
            Exit Sub
        End If
        exception = True
        rbRecord.Checked = True
        populateFields()
        My.Settings.sheet = "record"
        Try
            PopulateFiles()
        Catch ex As Exception
            MsgBox("Layer not loaded, add layer 'Geodex' to Map Document", MsgBoxStyle.OkOnly)

            exception = True 'set exception to true so that if someone closes the form, it doesn't ask if they want to reset the def query
            'because at this point it hasn't been changed yet.
            Me.Close()
        End Try
        btFlash.Enabled = False
        btSelect.Enabled = False
        btFilter.Enabled = False
        cmbSheet.Enabled = False
    End Sub

    'Populate the files when opening the form (this function is called during on_load)
    'Fills the drop down box for "series"
    'Iterates through the "subtypes" in the layer and populates the drop down list with the descriptions.
    'this uses the Series object type to override the tostring() function for the object.
    'Remember, the drop down box is not populated with text, its popoulated with objects.
    'This allows us to get more information than from just a index number.
    Private Sub PopulateFiles()
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer

        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass

        'this will only work if your feature class has subtypes
        Dim pSubtypes As ISubtypes = pFeatureClass

        Dim eSubtypes As IEnumSubtype = pSubtypes.Subtypes

        eSubtypes.Reset()

        Dim code As Integer
        Dim sSubtypetext As String

        sSubtypetext = eSubtypes.Next(code)

        cmbFile.Items.Clear()


        Do While sSubtypetext <> ""
            Dim pSeries As Series = New Series
            pSeries.GdxCode = code
            pSeries.GdxName = sSubtypetext
            cmbFile.Items.Add(pSeries)
            sSubtypetext = eSubtypes.Next(code)
        Loop

    End Sub

    'On selecting the series from the combo box
    'Change the defintion query of the layer to query the map and the table down to just this series.
    Private Sub cmbFile_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFile.SelectionChangeCommitted
        'Set the exception to FALSE because the def query will be changed now.  If we close this form,
        'the user will have the option to reset the def query.
        exception = False

        'Clear any existing items in the sheets drop down list.
        cmbSheet.Items.Clear()

        Dim pSelectedSeries As Series = cmbFile.SelectedItem
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer

        setDefinitionQuery("GDX_FILE", pSelectedSeries.GdxCode())

        btFilter.Enabled = True
        cmbSheet.Enabled = False
    End Sub

    'Populate the sheets after selecting the series from the combo box
    'Very similar to populating files, except its just filling it with some of the most important values.
    'Similar to the file drop down list being full of "series" objects, these are "sheet" objects.
    Private Sub PopulateSheets(pFCursor As IFeatureCursor)
        cmbSheet.Enabled = True
        Dim pcursor As System.Windows.Forms.Cursor = Cursor

        Dim pFeature As IFeature
        pFeature = pFCursor.NextFeature()
        cmbSheet.Items.Clear()

        'This loop uses a cursor to grab all sheets form the file and populate the sheet box
        'it uses the "sheet" class in the util.vb file and checks for null values to avoid casting errors

        Do Until pFeature Is Nothing
            pcursor = Cursors.WaitCursor
            Dim pSheet As Sheet = New Sheet

            Dim index As Integer = pFeature.Fields.FindField("RECORD")
            Dim sRecord As String
            If IsDBNull(pFeature.Value(index)) = False Then
                sRecord = pFeature.Value(index)
                pSheet.Record = sRecord
            Else
                pSheet.Record = "<Null>"
            End If

            Dim index2 As Integer = pFeature.Fields.FindField("LOCATION")
            Dim sLocation As String
            If IsDBNull(pFeature.Value(index2)) = False Then
                sLocation = pFeature.Value(index2)
                pSheet.Location = sLocation
            Else
                pSheet.Location = "<Null>"
            End If

            Dim index3 As Integer = pFeature.Fields.FindField("DATE")
            Dim sDate As Short
            If IsDBNull(pFeature.Value(index3)) = False Then
                sDate = pFeature.Value(index3)
                pSheet.SheetDate = sDate
            Else
                pSheet.SheetDate = 0
            End If

            Dim index4 As Integer = pFeature.Fields.FindField("OBJECTID")
            Dim sObjectID As String
            If IsDBNull(pFeature.Value(index4).ToString) = False Then
                sObjectID = pFeature.Value(index4).ToString
                pSheet.ObjectID = sObjectID
            End If

            cmbSheet.Items.Add(pSheet)
            pFeature = pFCursor.NextFeature


        Loop
        pcursor = Cursors.Default
    End Sub

    'This populates the drop down list of fields in the layer.  Basically, it just makes an entry for every field in the file.  NBD.
    Private Sub populateFields()
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pfeatureclass As IFeatureClass = pLayer.FeatureClass
        Dim pFields As IFields = pfeatureclass.Fields
        Dim fieldCount As Integer = pFields.FieldCount
        Dim pfield As String = ""
        cbFields.Items.Clear()
        For count As Integer = 0 To fieldCount - 1
            pfield = pFields.Field(count).Name
            cbFields.Items.Add(pfield)
        Next


    End Sub

    'Set the definition query to nothing (used in the refresh function
    Private Sub setDefinitionQuery()
        setDefinitionQuery("")
    End Sub

    'Set defintiion query that accepts fieldname as a string and category code, or as just a where clause
    Private Sub setDefinitionQuery(fieldname As String, categoryCode As Integer)
        setDefinitionQuery(fieldname & " =" & categoryCode)
    End Sub

    Private Sub setDefinitionQuery(whereclause As String)
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pMxDoc As IMxDocument = _application.Document

        Dim pSelectedLayer As ILayer = pLayer
        'This can give an error if the selected layer is NOT a "feature layer"
        Dim pFLayerDef As IFeatureLayerDefinition = pSelectedLayer
        pFLayerDef.DefinitionExpression = whereclause
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()

    End Sub

    'On selecting a sheet from the combo box
    'Zooms to the expanded envelope around the sheet, doesn't select it (just in case the selection is important to the user)
    'They can easily select the feature by clicking "select" once the feature is in focus in the combo box.
    Private Sub cmbSheet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSheet.SelectedIndexChanged
        Dim pSelectedSeries As Series = cmbFile.SelectedItem
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        If pSelectedSeries Is Nothing Then
            MsgBox("Select a series before selecting a sheet", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        Dim pSelectedSheet As Sheet = cmbSheet.SelectedItem
        If pSelectedSheet Is Nothing Then
            MsgBox("An error occured when selecting the sheet.  Try another sheet or try reselecting the series.", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pFeature As IFeature = pFeatureClass.GetFeature(pSelectedSheet.ObjectID)
        exception = False
        Dim pactiveView As IActiveView = pMap
        Dim pEnvelope As IEnvelope = pFeature.Extent
        pEnvelope.Expand(5, 5, True)
        btFlash.Enabled = True
        btSelect.Enabled = True
        pactiveView.Extent = pEnvelope
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()
    End Sub

    'On clicking the refresh button
    'Refresh the form, refill the drop down boxes.  Reset the definition query.  Zoom to the whole layer.  Refresh the map.
    Private Sub btRefresh_Click(sender As Object, e As EventArgs) Handles btRefresh.Click
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pFLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim player As ILayer = pFLayer
        cbOp.SelectedItem = "="
        tbFilter.Text = ""
        pMap.ClearSelection()
        setDefinitionQuery()
        cmbSheet.Items.Clear()
        cmbFile.Text = ""
        cmbSheet.SelectedText = ""
        cmbSheet.Text = ""
        'PopulateFiles()
        pMxDoc.UpdateContents()
        pMxDoc.ActivatedView.Refresh()
        exception = True
        'Zoom to layer
        Dim pactiveView As IActiveView = pMap
        Dim pEnvelope As IEnvelope = pFLayer.AreaOfInterest
        'pEnvelope.Expand(2.5, 2.5, True)
        PopulateFiles()
        pactiveView.Extent = pEnvelope
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()
        btFlash.Enabled = False
        btSelect.Enabled = False
        btFilter.Enabled = False
        cmbSheet.Enabled = False
    End Sub

    'On closing the form
    'Will ask the user if they want to reset the def query and the selection.  If exception is true, this is all skipped and the
    'form just closes.
    Private Sub frmViewer_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If exception = False Then
            Dim msg As MsgBoxResult = MsgBox("Do you want to reset the definition query to show all features and clear the current selection?", MsgBoxStyle.YesNoCancel)
            Dim pMxDoc As IMxDocument = _application.Document
            Dim pMap As IMap = pMxDoc.FocusMap

            If msg = MsgBoxResult.Yes Then
                setDefinitionQuery()
                pMap.ClearSelection()
                pMxDoc.UpdateContents()
                pMxDoc.ActivatedView.Refresh()
            ElseIf msg = MsgBoxResult.No Then
                pMxDoc.UpdateContents()
                pMxDoc.ActivatedView.Refresh()
            Else
                e.Cancel = True
            End If
        End If
    End Sub

    '
    Private Sub rbRecord_CheckedChanged(sender As Object, e As EventArgs) Handles rbRecord.CheckedChanged
        If rbRecord.Checked = True Then
            My.Settings.sheet = "record"
        Else
            My.Settings.sheet = "location"
        End If
    End Sub

    Private Sub rbRecord_MouseUp(sender As Object, e As Windows.Forms.MouseEventArgs) Handles rbRecord.MouseUp, rbLocation.MouseUp
        cmbSheet.Items.Clear()
        Dim pSelectedSeries As Series = cmbFile.SelectedItem
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pSelectedLayer As ILayer = pLayer
        Dim pFeatureLayer As IFeatureLayer = pSelectedLayer
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace
        Dim pSQLSyntax As ISQLSyntax = pWorkspace
        Dim ql As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierPrefix)
        Dim qr As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierSuffix)
        Dim pQFilter As IQueryFilter = New QueryFilter

        pQFilter.WhereClause = ql & "GDX_FILE" & qr & " = " & pSelectedSeries.GdxCode()
        Dim pFeatureCursor As IFeatureCursor = pFeatureClass.Search(pQFilter, False)

        'Populate the Sheets
        PopulateSheets(pFeatureCursor)
    End Sub

    Private Sub btFilter_Click(sender As Object, e As EventArgs) Handles btFilter.Click
        Try
            cmbSheet.Enabled = True
            Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()
            Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
            Dim pDataset As IDataset = pFeatureClass
            Dim pWorkspace As IWorkspace = pDataset.Workspace
            Dim pSQLSyntax As ISQLSyntax = pWorkspace

            Dim ql As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierPrefix)
            Dim qr As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierSuffix)

            Dim sField As String = cbFields.SelectedItem
            Dim pSelectedSeries As Series = cmbFile.SelectedItem

            Dim sOp As String = cbOp.SelectedItem
            Dim sClause As String = tbFilter.Text.Trim()
            Dim where As String

            exception = False
            cmbSheet.Items.Clear()
            If sClause = "" Then
                where = ql & "GDX_FILE" & qr & " = " & pSelectedSeries.GdxCode.ToString()
            Else
                where = ql & "GDX_FILE" & qr & " = " & pSelectedSeries.GdxCode.ToString() & " And " & ql & sField & qr & " " & sOp & " " & sClause
            End If

            setDefinitionQuery(where)

            Dim pQFilter As IQueryFilter = New QueryFilter

            pQFilter.WhereClause = where

            'Get the feature cursor to iterate through the features:
            Dim pFeatureCursor As IFeatureCursor = pFeatureClass.Search(pQFilter, False)

            'Populate the venues
            PopulateSheets(pFeatureCursor)

        Catch ex As Exception
            MsgBox(ex.ToString(), MsgBoxStyle.OkOnly)
        End Try
    End Sub

    Private Sub btFlash_Click(sender As Object, e As EventArgs) Handles btFlash.Click
        Dim pSelectedSheet As Sheet = cmbSheet.SelectedItem

        Dim pFlayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()

        Dim pFeatureClass As IFeatureClass = pFlayer.FeatureClass

        Dim pFeature As IFeature = pFeatureClass.GetFeature(pSelectedSheet.ObjectID)

        Dim pMxdoc As IMxDocument = _application.Document
        'Get the screen display so we can start drawing
        Dim pScreenDisplay As IScreenDisplay = pMxdoc.ActivatedView.ScreenDisplay

        pScreenDisplay.StartDrawing(pScreenDisplay.hDC, ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache)

        Dim pPolygon As IGeometry = pFeature.Shape

        Dim pMarkerSymbol As ISimpleFillSymbol = New SimpleFillSymbol

        Dim pColor As IRgbColor = New RgbColor
        pColor.RGB = RGB(255, 0, 0) 'Red
        pMarkerSymbol.Color = pColor

        Dim pSymbol As ISymbol = pMarkerSymbol
        pSymbol.ROP2 = esriRasterOpCode.esriROPNotXOrPen

        'set the symbol to the marker above, FLASH A BUNCH OF TIMES!
        pScreenDisplay.SetSymbol(pSymbol)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        Threading.Thread.Sleep(100)
        pScreenDisplay.DrawPolygon(pPolygon)
        pScreenDisplay.FinishDrawing()

    End Sub

    Private Sub btSelect_Click(sender As Object, e As EventArgs) Handles btSelect.Click
        Dim pSelectedSheet As Sheet = cmbSheet.SelectedItem
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pFeature As IFeature = pFeatureClass.GetFeature(pSelectedSheet.ObjectID)
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        pMap.ClearSelection()
        pMap.SelectFeature(pLayer, pFeature)
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()

    End Sub

 
    
End Class

