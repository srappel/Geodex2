Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geometry
Imports System.Windows.Forms

Public Class frmReconcile

    'Public Function GetSpecialCharacter(ByVal sqlsc As esriSQLSpecialCharacters) As String
    'End Function

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

    'On form load
    'Test if Geodex layer is added, if not return a msg box and close the form.
    'If Geodex layer is there, populate the series drop down list (populatefiles())
    'set the default sort order and jump type
    'make the check, x, and hold label invisible
    Private Sub frmReconcile_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim pLayer As ILayer = GDXTools.getInstance.GetGeodexLayer
        If pLayer Is Nothing Then
            MsgBox("Layer """ & My.Settings.FeatureClassName & """ is not loaded, use the load data button to load the layer." & vbCrLf & "ERROR CODE 101")
            Me.Close()
            Exit Sub
        End If

        cmbJumpType.SelectedItem = "OID"
        cmbSortOrder.SelectedItem = "Record > Location > p_date"
        Try
            PopulateFiles()
        Catch ex As Exception
            MsgBox("There was an error loading the series list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
            Exit Sub
        End Try

        My.Settings.CurrentOID = 0
        cmbSortOrder.SelectedIndex = 1
        cmbJumpType.SelectedIndex = 2
        picYes.Visible = False
        lblHold.Visible = False
        picNo.Visible = False

    End Sub

    'Handles clicking "Edit Record"
    'Opens a new frmRecordViewer.  Since the load function of record viewer checks if
    'my.settings.currentOID is set, this will pass the current oid, essentially.
    Private Sub cmdEdit_Click(sender As Object, e As EventArgs) Handles cmdEdit.Click
        'TODO: Add cmd_recordViewer.OnClick implementation
        'TODO: Add cmdIndexEdit.OnClick implementation

        Dim pRecordEdit As New frmRecordViewer

        pRecordEdit.ArcMapApplication = _application

        'create the arc map wrapper
        Dim pArcMapApplication As New ArcMapWrapper
        pArcMapApplication.ArcMapApplication = _application

        'Show the form
        pRecordEdit.Show(pArcMapApplication)

        Dim pMxDoc As IMxDocument = _application.Document

        'get the map from the document
        Dim pMap As IMap = pMxDoc.FocusMap
    End Sub

    'Populate the files when opening the form (this function is called during on_load)
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

    'This function sets the layers definition query when it is passed a whereclause as a string.
    'This is used when the series is selected to reduce the map down to just the selected series.
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

    'This is a little messy, but this alows the passing of a "setdefinitionquery" based on some variables.
    Private Sub setDefinitionQuery(fieldname As String, categoryCode As Integer)
        setDefinitionQuery(fieldname & " =" & categoryCode)
    End Sub

    'This is called when the selection is changed in the cmbFile drop down list (series)
    'the SelectionChangeCommitted option is used because it responds only when the user changes it, not when it is
    'changed programatically.
    Private Sub cmbFile_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbFile.SelectionChangeCommitted
        Dim pSelectedSeries As Series = cmbFile.SelectedItem
        setDefinitionQuery("GDX_FILE", pSelectedSeries.GdxCode())
    End Sub

    'When chosing a sort order.
    Private Sub cmbSortOrder_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbSortOrder.SelectionChangeCommitted
        'Nothing happens right now.
        'Considering making this sort the table, but haven't figured out how to do this yet.
        'I'll just leave this here in case.
    End Sub

    'On clicking "Search"
    'If there is no series selected, it will search all files.
    'There is a check that if there is a selected series, after it returns the search results, it will return the first
    'result that matches the file.
    Private Sub cmdSearch_Click(sender As Object, e As EventArgs) Handles cmdSearch.Click
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pfeature As IFeature = Nothing
        Dim pSelectedSeries As Series = cmbFile.SelectedItem
        Dim pFile As String = Nothing
        If Not pSelectedSeries Is Nothing Then
            pFile = pSelectedSeries.GdxCode.ToString()
        End If


        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace


        'This is actually super important.  This interface gets the correct SQL syntax for the workspace
        'Since this program can run on either SDE or File Geodatabases, you must include this type of coding.
        'Basically field names in SQL statements for File Geodatabases need to be surrounded by "" ie. "GDX_FILE"
        'in an SDE Geodatabase, the field is just GDX_FILE with no quotes.  ql is the "delimiter prefix.  ql is the delimiter suffix.
        Dim pSQLSyntax As ISQLSyntax = pWorkspace

        Dim ql As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierPrefix)
        Dim qr As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierSuffix)
        Dim wild As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_WildcardManyMatch)

        Dim sField As String = ""

        'This is the field that will be queried in the PQFilter whereclause
        If cmbJumpType.SelectedIndex = 0 Then
            sField = "OBJECTID"
        ElseIf cmbJumpType.SelectedIndex = 1 Then
            sField = "GDX_FILE"
        ElseIf cmbJumpType.SelectedIndex = 2 Then
            sField = "RECORD"
        ElseIf cmbJumpType.SelectedIndex = 3 Then
            sField = "LOCATION"
        Else
            MsgBox("Please select a field to search!")
            Exit Sub
        End If

        'Get the search string from the text
        Dim sSearch As String = ""
        If Not txtQuery.Text Is Nothing Then
            sSearch = txtQuery.Text.Trim().Trim("'")
        Else
            MsgBox("Please enter a query.  Ommit any quotation marks")
        End If

        'IQueryFilterDefinition interface allows us to add an "ORDER BY" statement
        'This is essential for sorting the records in the same order as they are stored in the map drawer.
        'For example, nautical charts are sorted by "RECORD" ascending (the map number) and then by DATE descending (the pub date)
        Dim pQFilter As IQueryFilter = New QueryFilter
        Dim pQFilterDefintion As IQueryFilterDefinition = pQFilter

        'Depending on the selection in cmbSortOrder, the ORDER BY clause in the query filter will change accordingly.
        If cmbSortOrder.SelectedIndex = 0 Then
            pQFilterDefintion.PostfixClause = "ORDER BY OBJECTID"
        ElseIf cmbSortOrder.SelectedIndex = 1 Then
            pQFilterDefintion.PostfixClause = "ORDER BY RECORD, DATE DESC"
        ElseIf cmbSortOrder.SelectedIndex = 2 Then
            pQFilterDefintion.PostfixClause = "ORDER BY RECORD, DATE"
        ElseIf cmbSortOrder.SelectedIndex = 3 Then
            pQFilterDefintion.PostfixClause = "ORDER BY LOCATION, DATE DESC"
        ElseIf cmbSortOrder.SelectedIndex = 4 Then
            pQFilterDefintion.PostfixClause = "ORDER BY LOCATION, DATE"
        ElseIf cmbSortOrder.SelectedIndex = 5 Then
            pQFilterDefintion.PostfixClause = "ORDER BY DATE DESC"
        ElseIf cmbSortOrder.SelectedIndex = 6 Then
            pQFilterDefintion.PostfixClause = "ORDER BY DATE"
        ElseIf cmbSortOrder.SelectedIndex = 7 Then
            pQFilterDefintion.PostfixClause = "ORDER BY SERIES_TIT, DATE DESC"
        ElseIf cmbSortOrder.SelectedIndex = 8 Then
            pQFilterDefintion.PostfixClause = "ORDER BY SERIES_TIT, DATE"
        End If

        'This defines the where clause.  see the notes on ql and qr above.  sField is the field selected and sSearch is from the
        'search text box.  "wild" comes from the pSQLsyntax interface as well.  This allows a partial search rather than exact strings.
        'Note that the search is still case sensitive, which sucks but I haven't figured out a good workaround.
        Dim sWhere As String = ql & sField & qr & " Like '" & wild & sSearch & wild & "'"
        pQFilter.WhereClause = sWhere

        'Here we define the cursor using the pQFilter.  Note that GDX_FILE is not included in the pQFILTER, its addressed below:
        pFeatureCursor = pFeatureClass.Search(pQFilter, False)

        'get the first record matching the where clause

        pfeature = pFeatureCursor.NextFeature()

        If Not pFile Is Nothing Then
            Do Until pfeature.Value(pfeature.Fields.FindField("GDX_FILE")) = pFile
                pfeature = pFeatureCursor.NextFeature()
            Loop
        End If

        'This will prevent an exception in the search.  It essentially means that the search was unsuccessfull
        If pfeature Is Nothing Then
            MsgBox("No feature was found using this query.")
            Exit Sub
        End If

        'Set the current OID to the oid of the selected record
        My.Settings.CurrentOID = pfeature.OID

        'Fills the form with the feature's attributes.
        fillreconcile(My.Settings.CurrentOID, pfeature, pFeatureClass)

    End Sub

    'Fill reconcile function
    'Similar to the fill function in the record viewer form
    'Get the values from the record, add them to a "Record" object
    'Take the values from the object and add them to the form.
    'Attributes that have coded domains are a little bit different, they enumerate through the domain values until the right one is found
    Private Sub fillreconcile(ByVal pOID As Integer, pFeature As IFeature, pFeatureClass As IFeatureClass)
        'Get the values from the record, add them to a pRecord

        'Take records from pRecord and populate the form

        Dim pRecord As New Record

        Try
            txtOID.Text = pOID

            'GDX_FILE
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("GDX_FILE"))) = False Then
                pRecord.GDX_FILE = pFeature.Value(pFeature.Fields.FindField("GDX_FILE"))
            End If
            If Not pRecord.GDX_FILE = Nothing Then
                txtFile.Text = pRecord.GDX_FILE.ToString()
            Else
                txtFile.Text = "<Null>"
            End If

            'Record
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("RECORD"))) = False Then
                pRecord.RECORD = pFeature.Value(pFeature.Fields.FindField("RECORD"))
            End If
            If Not pRecord.RECORD Is Nothing Then
                txtRecord.Text = pRecord.RECORD
            Else
                txtRecord.Text = "<Null>"
            End If

            'Location
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("LOCATION"))) = False Then
                pRecord.LOCATION = pFeature.Value(pFeature.Fields.FindField("LOCATION"))
            End If
            If Not pRecord.LOCATION Is Nothing Then
                txtLocation.Text = pRecord.LOCATION
            Else
                txtLocation.Text = "<Null>"
            End If

            'Series Title
            If Not pFeature.Value(pFeature.Fields.FindField("SERIES_TIT")) Is DBNull.Value Then
                pRecord.SERIES_TIT = pFeature.Value(pFeature.Fields.FindField("SERIES_TIT"))
            End If

            If Not pRecord.SERIES_TIT Is Nothing Then
                txtSeriesTit.Text = pRecord.SERIES_TIT
            Else
                txtSeriesTit.Text = "<Null>"
            End If

            'Publisher
            If Not pFeature.Value(pFeature.Fields.FindField("PUBLISHER")) Is DBNull.Value Then
                pRecord.PUBLISHER = pFeature.Value(pFeature.Fields.FindField("PUBLISHER"))
            End If

            If Not pRecord.PUBLISHER Is Nothing Then
                txtPublisher.Text = pRecord.PUBLISHER
            Else
                txtPublisher.Text = "<Null>"
            End If

            'Scale
            If Not pFeature.Value(pFeature.Fields.FindField("SCALE")) Is DBNull.Value Then
                pRecord.SCALE = pFeature.Value(pFeature.Fields.FindField("SCALE"))
            End If

            If Not pRecord.SCALE = Nothing Then
                txtScale.Text = pRecord.SCALE
            Else
                txtScale.Text = "<Null>"
            End If

            'Catloc
            If Not pFeature.Value(pFeature.Fields.FindField("CATLOC")) Is DBNull.Value Then
                pRecord.CATLOC = pFeature.Value(pFeature.Fields.FindField("CATLOC"))
            End If

            If Not pRecord.CATLOC Is Nothing Then
                txtCataloc.Text = pRecord.CATLOC
            Else
                txtCataloc.Text = "<Null>"
            End If

            '####################
            '#### attributes ####
            '### with domains ###
            '####################

            'Production
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("PRODUCTION"))) = False Then
                pRecord.PRODUCTION = pFeature.Value(pFeature.Fields.FindField("PRODUCTION"))
            Else
                pRecord.PRODUCTION = Nothing
            End If

            Dim pSubtypes As ISubtypes = pFeatureClass

            If Not pRecord.PRODUCTION = Nothing Then
                Dim subcode As Integer = 150
                Dim pProductionDomain As IDomain = pSubtypes.Domain(subcode, "PRODUCTION")
                Dim pCodedValueDomain2 As ICodedValueDomain = pProductionDomain

                Dim x As Integer

                For x = 0 To pCodedValueDomain2.CodeCount - 1
                    Dim pProduction As Production = New Production
                    pProduction.Description = pCodedValueDomain2.Name(x)
                    pProduction.Code = pCodedValueDomain2.Value(x)
                    Dim pProductionCode As Short = pProduction.Code
                    If pProductionCode = pRecord.PRODUCTION Then
                        txtProduction.Text = pProduction.Description.ToString()
                        Exit For
                    End If
                Next
            Else
                txtProduction.Text = "<Null>"
            End If

            'Map_type
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("MAP_TYPE"))) = False Then
                pRecord.MAP_TYPE = pFeature.Value(pFeature.Fields.FindField("MAP_TYPE"))
            Else
                pRecord.MAP_TYPE = Nothing
            End If

            If Not pRecord.MAP_TYPE = Nothing Then
                Dim subcode As Integer = 2
                Dim pMAP_TYPEDomain As IDomain = pSubtypes.Domain(subcode, "MAP_TYPE")
                Dim pCodedValueDomain As ICodedValueDomain = pMAP_TYPEDomain

                Dim x As Integer
                For x = 0 To pCodedValueDomain.CodeCount - 1
                    Dim pMAP_TYPE As MapType = New MapType
                    pMAP_TYPE.Code = pCodedValueDomain.Value(x)
                    pMAP_TYPE.Description = pCodedValueDomain.Name(x)
                    Dim pMAP_TYPEcode As Short = pMAP_TYPE.Code
                    If pMAP_TYPEcode = pRecord.MAP_TYPE Then
                        txtMapType.Text = pMAP_TYPE.Description.ToString()
                        Exit For
                    End If
                Next
            Else
                txtMapType.Text = "<Null>"
            End If

            'Projection
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("PROJECT"))) = False Then
                pRecord.PROJECT = pFeature.Value(pFeature.Fields.FindField("PROJECT"))
            Else
                pRecord.PROJECT = Nothing
            End If

            If Not pRecord.PROJECT = Nothing Then
                Dim subcode As Integer = 2
                Dim pProjectDomain As IDomain = pSubtypes.Domain(subcode, "PROJECT")
                Dim pCodedValueDomain As ICodedValueDomain = pProjectDomain

                Dim x As Integer
                For x = 0 To pCodedValueDomain.CodeCount - 1
                    Dim pProjection As Projection = New Projection
                    pProjection.Code = pCodedValueDomain.Value(x)
                    pProjection.Description = pCodedValueDomain.Name(x)
                    Dim pProjectCode As Short = pProjection.Code
                    If pProjectCode = pRecord.PROJECT Then
                        txtProjection.Text = pProjection.Description.ToString()
                        Exit For
                    End If
                Next
            Else
                txtProjection.Text = "<Null>"
            End If

            'Prime Merididan
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("PRIME_MER"))) = False Then
                pRecord.PRIME_MER = pFeature.Value(pFeature.Fields.FindField("PRIME_MER"))
            Else
                pRecord.PRIME_MER = Nothing
            End If

            If Not pRecord.PRIME_MER = Nothing Then
                Dim subcode As Integer = 2
                Dim pPrimeMerDomain As IDomain = pSubtypes.Domain(subcode, "PRIME_MER")
                Dim pCodedValueDomain As ICodedValueDomain = pPrimeMerDomain

                Dim x As Integer
                For x = 0 To pCodedValueDomain.CodeCount - 1
                    Dim pPrimeMer As PrimeMer = New PrimeMer
                    pPrimeMer.Code = pCodedValueDomain.Value(x)
                    pPrimeMer.Description = pCodedValueDomain.Name(x)
                    Dim pPrimerCode As Short = pPrimeMer.Code
                    If pPrimerCode = pRecord.PRIME_MER Then
                        txtPrimeMer.Text = pPrimeMer.Description.ToString()
                        Exit For
                    End If
                Next
            Else
                txtPrimeMer.Text = "<Null>"
            End If

            'Map Format
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("MAP_FOR"))) = False Then
                pRecord.MAP_FOR = pFeature.Value(pFeature.Fields.FindField("MAP_FOR"))
            Else
                pRecord.MAP_FOR = Nothing
            End If

            If Not pRecord.MAP_FOR = Nothing Then
                Dim subcode As Integer = 2
                Dim pFormatDomain As IDomain = pSubtypes.Domain(subcode, "MAP_FOR")
                Dim pCodedValueDomain As ICodedValueDomain = pFormatDomain

                Dim x As Integer
                For x = 0 To pCodedValueDomain.CodeCount - 1
                    Dim pMapFor As Format = New Format
                    pMapFor.Code = pCodedValueDomain.Value(x)
                    pMapFor.Description = pCodedValueDomain.Name(x)
                    Dim pFormatCode As Short = pMapFor.Code
                    If pFormatCode = pRecord.MAP_FOR Then
                        txtFormat.Text = pMapFor.Description.ToString()
                        Exit For
                    End If
                Next
            Else
                txtFormat.Text = "<Null>"
            End If

            Dim year1 As String = " "
            Dim year1type As String = " "
            Dim year2 As String = " "
            Dim year2type As String = " "
            Dim year3 As String = " "
            Dim year3type As String = " "
            Dim year4 As String = " "
            Dim year4type As String = " "


            'Year1_Type
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("YEAR1_TYPE"))) = False Then
                pRecord.YEAR1_TYPE = pFeature.Value(pFeature.Fields.FindField("YEAR1_TYPE"))
            Else
                pRecord.YEAR1_TYPE = Nothing
            End If

            If Not pRecord.YEAR1_TYPE = Nothing Then

                Dim subcode As Integer = 2
                Dim pYearDomain As IDomain = pSubtypes.Domain(subcode, "YEAR1_TYPE")
                Dim pCodedValueDomain As ICodedValueDomain = pYearDomain

                Dim x As Integer
                For x = 0 To pCodedValueDomain.CodeCount - 1
                    Dim pYearType As YearType = New YearType
                    pYearType.Code = pCodedValueDomain.Value(x)
                    pYearType.Description = pCodedValueDomain.Name(x)
                    Dim sYearTypeCode As Short = pYearType.Code
                    If sYearTypeCode = pRecord.YEAR1_TYPE Then
                        year1type = pYearType.Description & ": "
                        Exit For
                    End If
                Next
            End If

            'Year2_Type
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("YEAR2_TYPE"))) = False Then
                pRecord.YEAR2_TYPE = pFeature.Value(pFeature.Fields.FindField("YEAR2_TYPE"))
            Else
                pRecord.YEAR2_TYPE = Nothing
            End If

            If Not pRecord.YEAR2_TYPE = Nothing Then

                Dim subcode As Integer = 2
                Dim pYearDomain As IDomain = pSubtypes.Domain(subcode, "YEAR2_TYPE")
                Dim pCodedValueDomain As ICodedValueDomain = pYearDomain

                Dim x As Integer
                For x = 0 To pCodedValueDomain.CodeCount - 1
                    Dim pYearType As YearType = New YearType
                    pYearType.Code = pCodedValueDomain.Value(x)
                    pYearType.Description = pCodedValueDomain.Name(x)
                    Dim sYearTypeCode As Short = pYearType.Code
                    If sYearTypeCode = pRecord.YEAR2_TYPE Then
                        year2type = pYearType.Description & ": "
                        Exit For
                    End If
                Next
            End If

            'Year3_Type
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("YEAR3_TYPE"))) = False Then
                pRecord.YEAR3_TYPE = pFeature.Value(pFeature.Fields.FindField("YEAR3_TYPE"))
            Else
                pRecord.YEAR3_TYPE = Nothing
            End If

            If Not pRecord.YEAR3_TYPE = Nothing Then

                Dim subcode As Integer = 2
                Dim pYearDomain As IDomain = pSubtypes.Domain(subcode, "YEAR3_TYPE")
                Dim pCodedValueDomain As ICodedValueDomain = pYearDomain

                Dim x As Integer
                For x = 0 To pCodedValueDomain.CodeCount - 1
                    Dim pYearType As YearType = New YearType
                    pYearType.Code = pCodedValueDomain.Value(x)
                    pYearType.Description = pCodedValueDomain.Name(x)
                    Dim sYearTypeCode As Short = pYearType.Code
                    If sYearTypeCode = pRecord.YEAR3_TYPE Then
                        year3type = pYearType.Description & ": "
                        Exit For
                    End If
                Next
            End If

            'Year4_Type
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("YEAR4_TYPE"))) = False Then
                pRecord.YEAR4_TYPE = pFeature.Value(pFeature.Fields.FindField("YEAR4_TYPE"))
            Else
                pRecord.YEAR4_TYPE = Nothing
            End If

            If Not pRecord.YEAR4_TYPE = Nothing Then

                Dim subcode As Integer = 2
                Dim pYearDomain As IDomain = pSubtypes.Domain(subcode, "YEAR4_TYPE")
                Dim pCodedValueDomain As ICodedValueDomain = pYearDomain

                Dim x As Integer
                For x = 0 To pCodedValueDomain.CodeCount - 1
                    Dim pYearType As YearType = New YearType
                    pYearType.Code = pCodedValueDomain.Value(x)
                    pYearType.Description = pCodedValueDomain.Name(x)
                    Dim sYearTypeCode As Short = pYearType.Code
                    If sYearTypeCode = pRecord.YEAR4_TYPE Then
                        year4type = pYearType.Description & ": "
                        Exit For
                    End If
                Next
            End If

            If Not pFeature.Value(pFeature.Fields.FindField("YEAR1")) Is DBNull.Value Then
                pRecord.YEAR1 = pFeature.Value(pFeature.Fields.FindField("YEAR1"))
            End If
            If Not pRecord.YEAR1 = Nothing Then
                year1 = pRecord.YEAR1
            End If


            If Not pFeature.Value(pFeature.Fields.FindField("YEAR2")) Is DBNull.Value Then
                pRecord.YEAR2 = pFeature.Value(pFeature.Fields.FindField("YEAR2"))
            End If
            If Not pRecord.YEAR2 = Nothing Then
                year2 = pRecord.YEAR2
            End If


            If Not pFeature.Value(pFeature.Fields.FindField("YEAR3")) Is DBNull.Value Then
                pRecord.YEAR3 = pFeature.Value(pFeature.Fields.FindField("YEAR3"))
            End If
            If Not pRecord.YEAR3 = Nothing Then
                year3 = pRecord.YEAR3
            End If

            If Not pFeature.Value(pFeature.Fields.FindField("YEAR4")) Is DBNull.Value Then
                pRecord.YEAR4 = pFeature.Value(pFeature.Fields.FindField("YEAR4"))
            End If
            If Not pRecord.YEAR4 = Nothing Then
                year4 = pRecord.YEAR4
            End If

            Dim syearstring As String = year1type & year1 & vbCrLf &
                year2type & year2 & vbCrLf &
                year3type & year3 & vbCrLf &
                year4type & year4 & vbCrLf

            txtDates.Text = syearstring

            If Not pRecord.YEAR1 = Nothing Then
                txtDate.Text = pRecord.YEAR1.ToString()
            Else
                txtDate.Text = "<Null>"
            End If

            If Not pFeature.Value(pFeature.Fields.FindField("EDITION_NO")) Is DBNull.Value Then
                pRecord.EDITION_NO = pFeature.Value(pFeature.Fields.FindField("EDITION_NO"))
            End If
            If Not pRecord.EDITION_NO = Nothing Then
                txtEdition.Text = pRecord.EDITION_NO.ToString()
            Else
                txtEdition.Text = "<Null>"
            End If

            'X and Y fields!

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("X1"))) = False Then
                pRecord.X1 = pFeature.Value(pFeature.Fields.FindField("X1"))
            Else
                pRecord.X1 = Nothing
            End If

            If Not pRecord.X1 = Nothing Then
                lblW.Text = pRecord.X1
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("X2"))) = False Then
                pRecord.X2 = pFeature.Value(pFeature.Fields.FindField("X2"))
            Else
                pRecord.X2 = Nothing
            End If

            If Not pRecord.X2 = Nothing Then
                lblE.Text = pRecord.X2
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("Y1"))) = False Then
                pRecord.Y1 = pFeature.Value(pFeature.Fields.FindField("Y1"))
            Else
                pRecord.Y1 = Nothing
            End If

            If Not pRecord.Y1 = Nothing Then
                lblN.Text = pRecord.Y1
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("Y2"))) = False Then
                pRecord.Y2 = pFeature.Value(pFeature.Fields.FindField("Y2"))
            Else
                pRecord.Y2 = Nothing
            End If

            If Not pRecord.Y2 = Nothing Then
                lblS.Text = pRecord.Y2
            End If

        Catch ex As Exception
            MsgBox("Error occured!" & vbCrLf & ex.ToString)
            Exit Sub
            picYes.Visible = False
            picNo.Visible = False
            lblHold.Visible = False
        End Try

        If pFeature.Value(pFeature.Fields.FindField("HOLD")) = 0 Then
            picYes.Visible = False
            picNo.Visible = True
            lblHold.Text = "No"
        ElseIf pFeature.Value(pFeature.Fields.FindField("HOLD")) > 0 Then
            picYes.Visible = True
            picNo.Visible = False
            lblHold.Text = "Yes"
        Else
            picYes.Visible = False
            picNo.Visible = False
            lblHold.Visible = False
        End If


    End Sub

    'On clicking the "Next" button
    'Finds the current feature and then, basedo n the sort order, goes to the next feature.
    Private Sub cmbNext_Click(sender As Object, e As EventArgs) Handles cmbNext.Click

        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pOID As Integer = My.Settings.CurrentOID
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass

        Dim pfeature As IFeature = Nothing
        Dim pFile As Short = Nothing
        Dim pOID2 As Integer = Nothing

        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace

        Dim pSQLSyntax As ISQLSyntax = pWorkspace

        Dim ql As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierPrefix)
        Dim qr As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierSuffix)

        If My.Settings.CurrentOID = 0 Then
            MsgBox("Please perform a search to pick a starting point.")
            Exit Sub
        Else
            pfeature = pFeatureClass.GetFeature(pOID)
            pFile = pfeature.Value(pfeature.Fields.FindField("GDX_FILE"))
        End If

        Dim pQFilter As IQueryFilter = New QueryFilter
        Dim pQFilterDefintion As IQueryFilterDefinition = pQFilter

        If cmbSortOrder.SelectedIndex = 0 Then
            pQFilterDefintion.PostfixClause = "ORDER BY OBJECTID"
        ElseIf cmbSortOrder.SelectedIndex = 1 Then
            pQFilterDefintion.PostfixClause = "ORDER BY RECORD, DATE DESC"
        ElseIf cmbSortOrder.SelectedIndex = 2 Then
            pQFilterDefintion.PostfixClause = "ORDER BY RECORD, DATE"
        ElseIf cmbSortOrder.SelectedIndex = 3 Then
            pQFilterDefintion.PostfixClause = "ORDER BY LOCATION, DATE DESC"
        ElseIf cmbSortOrder.SelectedIndex = 4 Then
            pQFilterDefintion.PostfixClause = "ORDER BY LOCATION, DATE"
        ElseIf cmbSortOrder.SelectedIndex = 5 Then
            pQFilterDefintion.PostfixClause = "ORDER BY DATE DESC"
        ElseIf cmbSortOrder.SelectedIndex = 6 Then
            pQFilterDefintion.PostfixClause = "ORDER BY DATE"
        ElseIf cmbSortOrder.SelectedIndex = 7 Then
            pQFilterDefintion.PostfixClause = "ORDER BY SERIES_TIT, DATE DESC"
        ElseIf cmbSortOrder.SelectedIndex = 8 Then
            pQFilterDefintion.PostfixClause = "ORDER BY SERIES_TIT, DATE"
        End If

        pQFilter.WhereClause = ql & "GDX_FILE" & qr & " = " & pFile
        pFeatureCursor = pFeatureClass.Search(pQFilter, False)

        'loop through the feature cursor until the current record is found
        Dim pfeaturesearch As IFeature = pFeatureCursor.NextFeature()


        Do Until pfeaturesearch.OID = pOID
            pfeaturesearch = pFeatureCursor.NextFeature()
        Loop

        pfeature = pFeatureCursor.NextFeature()
        pOID2 = pfeature.OID

        fillreconcile(pOID2, pfeature, pFeatureClass)

        My.Settings.CurrentOID = pOID2
    End Sub

    'On clicking "previous" button 
    'Almost identical to the "next" button, but the sort orders are reversed so that the NEXT feature is actually the previous
    Private Sub cmbPrev_Click(sender As Object, e As EventArgs) Handles cmbPrev.Click
        Dim pFeatureCursor As IFeatureCursor = Nothing
        Dim pOID As Integer = My.Settings.CurrentOID
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass

        Dim pfeature As IFeature = Nothing
        Dim pFile As Short = Nothing
        Dim pOID2 As Integer = Nothing


        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace

        Dim pSQLSyntax As ISQLSyntax = pWorkspace


        Dim ql As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierPrefix)
        Dim qr As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierSuffix)

        If My.Settings.CurrentOID = 0 Then
            MsgBox("Please perform a search to pick a starting point.")
            Exit Sub
        Else
            pfeature = pFeatureClass.GetFeature(pOID)
            pFile = pfeature.Value(pfeature.Fields.FindField("GDX_FILE"))
        End If

        Dim pQFilter As IQueryFilter = New QueryFilter
        Dim pQFilterDefintion As IQueryFilterDefinition = pQFilter

        If cmbSortOrder.SelectedIndex = 0 Then
            pQFilterDefintion.PostfixClause = "ORDER BY OBJECTID desc"
        ElseIf cmbSortOrder.SelectedIndex = 1 Then
            pQFilterDefintion.PostfixClause = "ORDER BY RECORD desc, DATE asc"
        ElseIf cmbSortOrder.SelectedIndex = 2 Then
            pQFilterDefintion.PostfixClause = "ORDER BY RECORD desc, DATE desc"
        ElseIf cmbSortOrder.SelectedIndex = 3 Then
            pQFilterDefintion.PostfixClause = "ORDER BY LOCATION desc, DATE asc"
        ElseIf cmbSortOrder.SelectedIndex = 4 Then
            pQFilterDefintion.PostfixClause = "ORDER BY LOCATION desc, DATE desc"
        ElseIf cmbSortOrder.SelectedIndex = 5 Then
            pQFilterDefintion.PostfixClause = "ORDER BY DATE asc"
        ElseIf cmbSortOrder.SelectedIndex = 6 Then
            pQFilterDefintion.PostfixClause = "ORDER BY DATE desc"
        ElseIf cmbSortOrder.SelectedIndex = 7 Then
            pQFilterDefintion.PostfixClause = "ORDER BY SERIES_TIT desc, DATE asc"
        ElseIf cmbSortOrder.SelectedIndex = 8 Then
            pQFilterDefintion.PostfixClause = "ORDER BY SERIES_TIT desc, DATE desc"
        End If

        pQFilter.WhereClause = ql & "GDX_FILE" & qr & " = " & pFile
        pFeatureCursor = pFeatureClass.Search(pQFilter, False)

        'loop through the feature cursor until the current record is found
        Dim pfeaturesearch As IFeature = pFeatureCursor.NextFeature()
        Do Until pfeaturesearch.OID = pOID
            pfeaturesearch = pFeatureCursor.NextFeature()
        Loop

        pfeature = pFeatureCursor.NextFeature()
        pOID2 = pfeature.OID

        fillreconcile(pOID2, pfeature, pFeatureClass)

        My.Settings.CurrentOID = pOID2
    End Sub

    'On clicking "get selected"
    'Gets the feature selected in the table as the starting point of the reconcile.
    'Only works if ONE feature is selected,  no more, no less.
    Private Sub cmdGetSelect_click(sender As Object, e As EventArgs) Handles cmdGetSelect.Click
        Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pMXDocument As IMxDocument = _application.Document
        Dim pMap As IMap = pMXDocument.FocusMap
        Dim count As Integer = pMap.SelectionCount
        Dim pFeature As IFeature = Nothing
        Dim pOID As Integer = 0
        Dim pFeatureSelection As IFeatureSelection = pFeatureLayer
        Dim pSelectionSet As ISelectionSet = pFeatureSelection.SelectionSet
        Dim eIDS As IEnumIDs = pSelectionSet.IDs
        If count <> 1 Then
            MsgBox("this function will only work if one feature is selected in the table")
            Exit Sub
        Else
            pOID = eIDS.Next
            My.Settings.CurrentOID = pOID
            pFeature = pFeatureClass.GetFeature(pOID)
            Try
                fillreconcile(pOID, pFeature, pFeatureClass)
            Catch ex As Exception
                MsgBox("There was an error when filling the form:" & vbCrLf & ex.ToString())
                Exit Sub
            End Try
        End If
    End Sub

    'Clears the form (Who wudda thunk it?)
    Private Sub clearform()
        My.Settings.CurrentOID = 0
        txtOID.Clear()
        cmbFile.SelectedItem = Nothing
        txtQuery.Clear()
        txtRecord.Clear()
        txtLocation.Clear()
        txtDate.Clear()
        txtEdition.Clear()
        txtCataloc.Clear()
        txtSeriesTit.Clear()
        txtPublisher.Clear()
        txtScale.Clear()
        txtMapType.Clear()
        txtProduction.Clear()
        txtPrimeMer.Clear()
        txtFormat.Clear()
        txtHold.Text = "1"
        txtFile.Clear()
        txtProjection.Clear()
        txtDates.Clear()
        lblN.Text = "North Extent"
        lblE.Text = "East Extent"
        lblW.Text = "West Extent"
        lblS.Text = "South Extent"
    End Sub

    'This function edits the value of the holding field.  This is the only edit function in the frmReconcile.
    'If othe fields need to be edited, the user will use the "edit record" button and it will take them to the frmRecordViewer
    'This is a pretty straighforward sub.  Pass it an OID and the short value for holdings 0 for none, or the # of copies.
    Private Sub editHolding(ByVal pOID As Integer, ByVal sHold As Short)
        Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace
        Dim pWorkspaceEdit As IWorkspaceEdit = pWorkspace
        Dim pFeature As IFeature = pFeatureClass.GetFeature(pOID)
        Try
            pWorkspaceEdit.StartEditing(True)
            pWorkspaceEdit.StartEditOperation()
            pFeature.Value(pFeature.Fields.FindField("HOLD")) = sHold
            pFeature.Value(pFeature.Fields.FindField("RUN_DATE")) = DateAndTime.DateString() & My.User.Name.ToString()
            pFeature.Store()
            pWorkspaceEdit.StartEditOperation()
            pWorkspaceEdit.StopEditing(True)
        Catch ex As Exception
            MsgBox("There was an error editing the holding information." & vbCrLf & ex.ToString())
        End Try
    End Sub

    'On the user clicking yes
    'Gets the number of copies entered in the copies text box.
    'runs the "edit holdings" function with the necessary values.
    Private Sub cmdYES_Click(sender As Object, e As EventArgs) Handles cmdYES.Click
        If Not My.Settings.CurrentOID = 0 Then
        Else
            MsgBox("No OID is selected.", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        Dim pOID As Integer = My.Settings.CurrentOID
        Dim sHold As Short = txtHold.Text.Trim()
        Try
            editHolding(pOID, sHold)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Exit Sub
        End Try
        Try
            Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()
            Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
            Dim pFeature As IFeature = pFeatureClass.GetFeature(pOID)
            fillreconcile(pOID, pFeature, pFeatureClass)
            picYes.Visible = True
            picNo.Visible = False
            lblHold.Text = "Yes"
        Catch ex As Exception
            MsgBox(ex.ToString())
            Exit Sub
        End Try
    End Sub

    'on the user clicking NO
    'same as YES except the value for holdings is hard coded to 0
    Private Sub cmdNO_Click(sender As Object, e As EventArgs) Handles cmdNO.Click
        If Not My.Settings.CurrentOID = 0 Then
        Else
            MsgBox("No OID is selected.", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        Dim pOID As Integer = My.Settings.CurrentOID
        Dim sHold As Short = 0
        Try
            editHolding(pOID, sHold)
        Catch ex As Exception
            MsgBox(ex.ToString())
            Exit Sub
        End Try
        Try
            Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()
            Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
            Dim pFeature As IFeature = pFeatureClass.GetFeature(pOID)
            fillreconcile(pOID, pFeature, pFeatureClass)
            picYes.Visible = False
            picNo.Visible = True
            lblHold.Text = "No"
        Catch ex As Exception
            MsgBox(ex.ToString())
            Exit Sub
        End Try
    End Sub

    'On closing the form.
    'Reset the definition expression
    'reset the active view
    'resent the my.settings.currentOID to 0
    Private Sub frmReconcile_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        'The first check makes sure that these things don't run if the geodex layer isn't open
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        If pLayer Is Nothing Then
            Exit Sub
        End If
        Try
            Dim pMxDoc As IMxDocument = _application.Document
            Dim pSelectedLayer As ILayer = pLayer
            Dim pFLayerDef As IFeatureLayerDefinition = pSelectedLayer
            pFLayerDef.DefinitionExpression = ""
            pMxDoc.UpdateContents()
            pMxDoc.ActiveView.Refresh()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        My.Settings.CurrentOID = 0
    End Sub

    'On clicking "DELETE"
    'OK so I lied, this edits a feature too, but it just deletes it.  This produces a confirmation box so it isn't clicked accidentally.
    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        If Not My.Settings.CurrentOID = 0 Then
        Else
            MsgBox("No OID is selected.", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        Dim msg As MsgBoxResult = MsgBox("Are you sure you want to delete the feature with OID:" & My.Settings.CurrentOID.ToString() & "?", MsgBoxStyle.YesNo)
        If msg = MsgBoxResult.Yes Then
            Dim pOID As Integer = My.Settings.CurrentOID
            Dim pFLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
            Dim pFeatureClass As IFeatureClass = pFLayer.FeatureClass
            Dim pFeature As IFeature = pFeatureClass.GetFeature(pOID)
            Dim pDataset As IDataset = pFeatureClass
            Dim pWorkspace As IWorkspace = pDataset.Workspace
            Dim pWorkspaceEdit As IWorkspaceEdit = pWorkspace
            Try
                pWorkspaceEdit.StartEditing(True)
                pWorkspaceEdit.StartEditOperation()
                pFeature.Delete()
                pFeature.Store()
                'pWorkspaceEdit.StopEditingoperation()
                pWorkspaceEdit.StopEditing(True)
            Catch ex As Exception
                MsgBox("There was an error deleting the feature: " & vbCrLf & ex.ToString(), MsgBoxStyle.OkOnly)
                Exit Sub
            End Try

            Try
                clearform()
            Catch ex As Exception
                MsgBox(ex.ToString(), MsgBoxStyle.OkOnly)
            End Try
            My.Settings.CurrentOID = 0
        ElseIf msg = MsgBoxResult.No Then
            Exit Sub
        Else
            Exit Sub
        End If
    End Sub

    'On clicking CLEAR.  Clears the form using the clearform() sub
    'Also refreshes the definition expression, refreshes the active view, and resets the currentOID in my.settings
    Private Sub cmbClear_Click(sender As Object, e As EventArgs) Handles cmbClear.Click
        Try
            clearform()
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pSelectedLayer As ILayer = pLayer
        Dim pFLayerDef As IFeatureLayerDefinition = pSelectedLayer
        pFLayerDef.DefinitionExpression = Nothing
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()
        My.Settings.CurrentOID = 0
    End Sub

    'Flash
    'Flashes a feature.  I barely understand this.
    Private Sub cmbFlash_Click(sender As Object, e As EventArgs) Handles cmbFlash.Click
        If My.Settings.CurrentOID = 0 Then
            Exit Sub
        End If

        Dim pFlayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()

        Dim pFeatureClass As IFeatureClass = pFlayer.FeatureClass

        Dim pFeature As IFeature = pFeatureClass.GetFeature(My.Settings.CurrentOID)

        Dim pMxdoc As IMxDocument = _application.Document
        'Get the screen display so we can start drawing
        Dim pScreenDisplay As IScreenDisplay = pMxdoc.ActivatedView.ScreenDisplay

        pScreenDisplay.StartDrawing(pScreenDisplay.hDC, ESRI.ArcGIS.Display.esriScreenCache.esriNoScreenCache)
        If pFeature.Shape Is Nothing Then
            MsgBox("The currently selected OID has no shape!")
            Exit Sub
        End If

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

    'Zoom To
    'Zooms to the focused record.  Defines an envelope to the size of the feature and "expands it" so it isn't TOO close.
    'Changes the values of "expand" if you really need to.  I won't be offended.
    Private Sub cmbZoom_Click(sender As Object, e As EventArgs) Handles cmbZoom.Click
        If My.Settings.CurrentOID = 0 Then
            Exit Sub
        End If
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pFeature As IFeature = pFeatureClass.GetFeature(My.Settings.CurrentOID)
        Dim pactiveView As IActiveView = pMap
        If pFeature.Shape Is Nothing Then
            MsgBox("Cannot calculate extent to zoom because feature has no shape!")
            Exit Sub
        End If
        Dim pEnvelope As IEnvelope = pFeature.Extent
        pEnvelope.Expand(5, 5, True)
        pactiveView.Extent = pEnvelope
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()
    End Sub

    'Select
    'Selects the feature on the map and in the table for the focused OID.
    Private Sub cmbSelect_Click(sender As Object, e As EventArgs) Handles cmbSelect.Click
        If My.Settings.CurrentOID = 0 Then
            Exit Sub
        End If
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pFeature As IFeature = pFeatureClass.GetFeature(My.Settings.CurrentOID)
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        pMap.ClearSelection()
        pMap.SelectFeature(pLayer, pFeature)
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()
    End Sub

End Class