'Imports:
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports gdx_editor_tools.frmIndexEdit
Imports ESRI.ArcGIS.Display
Imports System.Windows.Forms



'Main Class:
Public Class frmRecordViewer
    'This defines the application as Arc Map
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
                e.Handled = True

            End If
        End If

       

    End Sub


    ':::::::::::::::::Load Functions::::::::::::::::::
    'Runs when frmRecordViewer loads
    'Gets the GDX layer, populates the files, subtypes, year types, map type, map format, production
    'projection, and prime meridian drop down boxes.
    'ERROR CODE 100: A FUNCTION IN frmRecordViewer_Load() COMMAND RETURNED AN EXCEPTION:
    'ERROR CODE 101: The Geodex Layer is Not Loaded
    Private Sub frmRecordViewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        


        Dim pLayer As ILayer = GDXTools.getInstance.GetGeodexLayer
        If pLayer Is Nothing Then
            MsgBox("Layer """ & My.Settings.FeatureClassName & """ is not loaded, use the load data button to load the layer." & vbCrLf & "ERROR CODE 101")
            Me.Close()
            Exit Sub
        End If

        Try
            PopulateFiles()
        Catch ex As Exception
            MsgBox("There was an error loading the series list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
        End Try
        dudHold.SelectedIndex = 1

        Try
            PopulateSubs()
        Catch ex As Exception
            MsgBox("There was an error loading the Subtype list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
        End Try

        Try
            PopulateYearDomain()

        Catch ex As Exception
            MsgBox("There was an error loading the Year Type list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
        End Try

        Try
            PopulateIsoType()
        Catch ex As Exception
            MsgBox("There was an error loading the Iso_Type list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
        End Try

        Try
            PopulateMapFor()
        Catch ex As Exception
            MsgBox("There was an error loading the Format list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
        End Try

        Try
            PopulateMapType()
        Catch ex As Exception
            MsgBox("There was an error loading the Map Type list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
        End Try

        Try
            PopulatePrime()
        Catch ex As Exception
            MsgBox("There was an error loading the Prime Meridian list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
        End Try

        Try
            PopulateProduction()
        Catch ex As Exception
            MsgBox("There was an error loading the Production list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
        End Try

        Try
            PopulateProject()
        Catch ex As Exception
            MsgBox("There was an error loading the Projection list" & vbCrLf & ex.ToString() & vbCrLf & "ERROR CODE 100", MsgBoxStyle.OkOnly)
            Me.Close()
        End Try

        If Not My.Settings.CurrentOID = 0 Then
            Dim pOID As Integer = My.Settings.CurrentOID
            
            Try
                FillForm(pOID)
            Catch ex As Exception
                MsgBox("There was an error when filling the form:" & vbCrLf & ex.ToString())
                My.Settings.CurrentOID = 0
                Exit Sub
            End Try
        End If
        lblOID.Text = My.Settings.CurrentOID.ToString()
        '"dudhold" is the holding box, index 19 is actually "1"
        dudHold.SelectedIndex = 19





    End Sub


    'Got rid of this because it was more annoying than it was useful...
    'Check for non-numeric characters in reserved fields (7)
    'Private Sub NumCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtScale.KeyPress, txtIsoVal.KeyPress, txtYear1.KeyPress, txtYear2.KeyPress, txtYear3.KeyPress, txtYear4.KeyPress, txtEdition.KeyPress, txtOID.KeyPress

    '    If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
    '        MsgBox("Please enter numbers only", MsgBoxStyle.OkOnly)
    '        e.Handled = True
    '    End If
    'End Sub

    'This function resets the form.  It's called under a couple commands.
    Private Sub ClearForm()
        txtOID.Clear()
        cmbFile.SelectedItem = Nothing
        cmb_subtype.SelectedItem = Nothing
        txtRecord.Clear()
        txtLocation.Clear()
        tbP_Date.Clear()
        txtEdition.Clear()
        txtCatalog.Clear()
        txtSeriesTit.Clear()
        txtPublisher.Clear()
        txtScale.Clear()
        cmbMapType.SelectedItem = Nothing
        cmbProduction.SelectedItem = Nothing
        cmbProjection.SelectedItem = Nothing
        cmbPrimeMer.SelectedItem = Nothing
        cmbFormat.SelectedItem = Nothing
        cmbIsoType.SelectedItem = Nothing
        txtIsoVal.Clear()
        dudHold.SelectedItem = 1
        cmbYear1.SelectedItem = Nothing
        cmbYear2.SelectedItem = Nothing
        cmbYear3.SelectedItem = Nothing
        cmbYear4.SelectedItem = Nothing
        txtYear1.Clear()
        txtYear2.Clear()
        txtYear3.Clear()
        txtYear4.Clear()
        lblNorth.Text = "North Extent"
        lblEast.Text = "East Extent"
        lblWest.Text = "West Extent"
        lblSouth.Text = "South Extent"
        dudHold.SelectedIndex = 19
        lblOID.Text = "..."

    End Sub

    'Populate the isobar types drop down list
    'Note that the list is not just filled with text, it is filled with OBJECTS
    'These objects are defined in class "IsoType" and have a code and a description
    'There is a public override that forces .tostring() to spit out the text description
    Private Sub PopulateIsoType()
        Dim subcode As Integer = 2
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pSubtypes As ISubtypes = pFeatureClass
        Dim pIsoDomain As IDomain = pSubtypes.Domain(subcode, "ISO_TYPE")
        Dim pCodedValueDomain As ICodedValueDomain = pIsoDomain
        If pCodedValueDomain Is Nothing Then
            Exit Sub
        End If

        'clear the list of any items there currently
        cmbIsoType.Items.Clear()

        'iterates through the domain coded values and adds them to the year type combo boxes
        Dim x As Integer
        For x = 0 To pCodedValueDomain.CodeCount - 1
            Dim pIsoType As IsoType = New IsoType
            pIsoType.Code = pCodedValueDomain.Value(x)
            pIsoType.Description = pCodedValueDomain.Name(x)
            cmbIsoType.Items.Add(pIsoType)
        Next
    End Sub

    'Populate the Map Types drop down menu
    'Filled with objects of type MapType (See IsoType Above)
    Private Sub PopulateMapType()
        Dim subcode As Integer = 2
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pSubtypes As ISubtypes = pFeatureClass
        Dim pMapTypeDomain As IDomain = pSubtypes.Domain(subcode, "MAP_TYPE")
        Dim pCodedValueDomain As ICodedValueDomain = pMapTypeDomain
        If pCodedValueDomain Is Nothing Then Exit Sub

        cmbMapType.Items.Clear()

        'iterates through the domain coded values and adds them to the year type combo boxes
        Dim x As Integer
        For x = 0 To pCodedValueDomain.CodeCount - 1
            Dim pMapType As MapType = New MapType
            pMapType.Code = pCodedValueDomain.Value(x)
            pMapType.Description = pCodedValueDomain.Name(x)
            cmbMapType.Items.Add(pMapType)
        Next
    End Sub

    'Populates the Format drop down menu
    'fileld with objectsw of type "Format" (See IsoType Above)
    Private Sub PopulateMapFor()
        Dim subcode As Integer = 2
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pSubtypes As ISubtypes = pFeatureClass
        Dim pMapForDomain As IDomain = pSubtypes.Domain(subcode, "MAP_FOR")
        Dim pCodedValueDomain As ICodedValueDomain = pMapForDomain
        If pCodedValueDomain Is Nothing Then
            Exit Sub
        End If
        cmbFormat.Items.Clear()

        'iterates through the domain coded values and adds them to the year type combo boxes
        Dim x As Integer
        For x = 0 To pCodedValueDomain.CodeCount - 1
            Dim pFormat As Format = New Format
            pFormat.Code = pCodedValueDomain.Value(x)
            pFormat.Description = pCodedValueDomain.Name(x)
            cmbFormat.Items.Add(pFormat)
        Next
    End Sub

    'Populates the Projection drop down menu
    'filled with objectsw of type "Projection" (See IsoType Above)
    Private Sub PopulateProject()
        Dim subcode As Integer = 2
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pSubtypes As ISubtypes = pFeatureClass
        Dim pProjectDomain As IDomain = pSubtypes.Domain(subcode, "PROJECT")
        Dim pCodedValueDomain As ICodedValueDomain = pProjectDomain
        If pCodedValueDomain Is Nothing Then Exit Sub
        cmbProjection.Items.Clear()

        'iterates through the domain coded values and adds them to the year type combo boxes
        Dim x As Integer
        For x = 0 To pCodedValueDomain.CodeCount - 1
            Dim pProject As Projection = New Projection
            pProject.Code = pCodedValueDomain.Value(x)
            pProject.Description = pCodedValueDomain.Name(x)
            cmbProjection.Items.Add(pProject)
        Next
    End Sub

    'Populates the Prime Meridian drop down menu
    'filled with objectsw of type "PrimeMer" (See IsoType Above)
    Private Sub PopulatePrime()
        Dim subcode As Integer = 2
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pSubtypes As ISubtypes = pFeatureClass
        Dim pPrimeDomain As IDomain = pSubtypes.Domain(subcode, "PRIME_MER")
        Dim pCodedValueDomain As ICodedValueDomain = pPrimeDomain
        If pCodedValueDomain Is Nothing Then Exit Sub
        cmbPrimeMer.Items.Clear()

        'iterates through the domain coded values and adds them to the year type combo boxes
        Dim x As Integer
        For x = 0 To pCodedValueDomain.CodeCount - 1
            Dim pPrime As PrimeMer = New PrimeMer
            pPrime.Code = pCodedValueDomain.Value(x)
            pPrime.Description = pCodedValueDomain.Name(x)
            cmbPrimeMer.Items.Add(pPrime)
        Next
    End Sub

    ''Populates the Production drop down menu
    'filled with objectsw of type "Production" (See IsoType Above)
    Private Sub PopulateProduction()
        Dim subcode As Integer = 2
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pSubtypes As ISubtypes = pFeatureClass
        Dim pProductionDomain As IDomain = pSubtypes.Domain(subcode, "PRODUCTION")
        Dim pCodedValueDomain As ICodedValueDomain = pProductionDomain
        If pCodedValueDomain Is Nothing Then Exit Sub
        cmbProduction.Items.Clear()

        'iterates through the domain coded values and adds them to the year type combo boxes
        Dim x As Integer
        For x = 0 To pCodedValueDomain.CodeCount - 1
            Dim pProduction As Production = New Production
            pProduction.Code = pCodedValueDomain.Value(x)
            pProduction.Description = pCodedValueDomain.Name(x)
            cmbProduction.Items.Add(pProduction)
        Next
    End Sub

    ''Populates the Year type drop down menus
    'filled with objectsw of type "YearType" (See IsoType Above)
    'Notice there are 4 different combo boxes but they all share a class "YearType"
    Private Sub PopulateYearDomain()

        'This subcode could be changed if different year domains were used for different
        'subtypes.  Change the name of the sub to PopulateYearDomain(byVal subcode as integer)
        'And when calling the funciton, provide a subtye code
        Dim subcode As Integer = 2
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pSubtypes As ISubtypes = pFeatureClass
        Dim pYearDomain As IDomain = pSubtypes.Domain(subcode, "YEAR1_TYPE")
        Dim pCodedValueDomain As ICodedValueDomain = pYearDomain
        If pCodedValueDomain Is Nothing Then
            Exit Sub
        End If

        cmbYear1.Items.Clear()
        cmbYear2.Items.Clear()
        cmbYear3.Items.Clear()
        cmbYear4.Items.Clear()

        'iterates through the domain coded values and adds them to the year type combo boxes
        Dim x As Integer
        For x = 0 To pCodedValueDomain.CodeCount - 1
            Dim pYearType As YearType = New YearType
            pYearType.Code = pCodedValueDomain.Value(x)
            pYearType.Description = pCodedValueDomain.Name(x)
            cmbYear1.Items.Add(pYearType)
            cmbYear2.Items.Add(pYearType)
            cmbYear3.Items.Add(pYearType)
            cmbYear4.Items.Add(pYearType)
        Next

    End Sub

    'Populates the Series drop down menu
    'filled with objectsw of type "Series"
    'Unlike the other functions which are iterating through domains
    'this function iterates through subtypes since "GDX_FILE" is a subtype field
    Private Sub PopulateFiles()
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer

        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass

        'this will only work if your feature class has subtypes
        Dim pSubtypes As ISubtypes = pFeatureClass

        'This is a subtype enumerator:  You can iterate through the subtypes
        Dim eSubtypes As IEnumSubtype = pSubtypes.Subtypes

        eSubtypes.Reset()

        Dim code As Integer
        Dim sSubtypetext As String

        sSubtypetext = eSubtypes.Next(code)

        cmbFile.Items.Clear()

        'This loops through the subtypes in the eSubtypes enumerator and
        'adds each subtype as an object of type "Series" to the series combo box
        'drop down list.
        Do While sSubtypetext <> ""
            Dim pSeries As Series = New Series
            pSeries.GdxCode = code
            pSeries.GdxName = sSubtypetext
            cmbFile.Items.Add(pSeries)
            sSubtypetext = eSubtypes.Next(code)
        Loop

    End Sub

    'Populates the Sub Type drop down menu
    'filled with objectsw of type "GDX_SUB" (See IsoType Above)
    'NOTE: This is not actually a "subtype" as far as Arc is concerned.  I call it a subtype
    'but it could be considered a sub classification of geodex files (See GDX_SUB field in the feature class)
    Private Sub PopulateSubs()
        Dim subcode As Integer = 2
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pLayer.FeatureClass
        Dim pSubtypes As ISubtypes = pFeatureClass
        Dim pSubtypeDomain As IDomain = pSubtypes.Domain(subcode, "GDX_SUB")
        Dim pCodedValueDomain As ICodedValueDomain = pSubtypeDomain
        If pCodedValueDomain Is Nothing Then
            Exit Sub
        End If

        cmb_subtype.Items.Clear()

        'iterates through the domain coded values and adds them to the year type combo boxes
        Dim x As Integer
        For x = 0 To pCodedValueDomain.CodeCount - 1
            Dim pSubtype As GDX_SUB = New GDX_SUB
            pSubtype.Code = pCodedValueDomain.Value(x)
            pSubtype.Description = pCodedValueDomain.Name(x)
            cmb_subtype.Items.Add(pSubtype)
        Next
    End Sub

    'These load functions are all called when the form loads.  That's why it is necessary to have
    'the Geodex layer in arc map before calling the cmdRecordViewer to load the frm

    ':::::::::::::::::Major Functions::::::::::::::::::


    'Create a new Feature (27)
    'ERROR CODE 120: Error when running the newfeature() command
    Private Sub newfeature()
        Try
            Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
            Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
            Dim pDataset As IDataset = pFeatureClass
            Dim pWorkspace As IWorkspace = pDataset.Workspace
            Dim pWorkspaceEdit As IWorkspaceEdit = pWorkspace

            Try
                pWorkspaceEdit.StartEditing(True)
                pWorkspaceEdit.StartEditOperation()
            Catch ex As Exception
                MsgBox("Error starting an editing operation")
                Exit Sub
            End Try
            
            'Do editing


            Dim pNewFeature As IFeature = pFeatureClass.CreateFeature()

            'WHY IS THIS HERE??
            'Dim pPoly As IPolygon = New Polygon

            Try
                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(True)
            Catch ex As Exception
                MsgBox("Error ending an editing operation")
                Exit Sub
            End Try
            

            If pNewFeature.HasOID = True Then
                My.Settings.CurrentOID = pNewFeature.OID
            Else
                MsgBox("The new feature has no ObjectID", MsgBoxStyle.OkOnly)
            End If

        Catch ex As Exception
            MsgBox("AN ERROR OCCURED while creating a new feature: " & ex.ToString() & vbCrLf & "ERROR CODE 120", MsgBoxStyle.OkOnly)
        End Try
    End Sub

    'Edit a feature (called by commit, commit and copy) (381)
    'ERROR CODE 121: Edit feature was called but a GDX series was not assigned
    'ERROR CODE 122: Value for scale must only be numeric digits.  Exception will cancel the edit operation.

    'This is a pretty massive script, about 400 lines of code... but the general process is:
    '1. Get the values from the form
    '2. Check to make sure the values are valid
    '3. Assign the values to a class called "Record" which defines the needed variables for
    'a record in the table
    '   - some of the records have special cases like years that need to be sorted.  see that part of the code
    '     for more infomration on that.
    '4. Start an editing session on the workspace with the passed OID
    '5. Take that new "record" object (pRecord) and use its new attributes to assign values to fields in the table.
    '6. Close the editing session.

    Private Sub EditFeature(ByVal OID As Integer)
        '::::::::::::::::::::::::::::::
        '::::::SETTING THE VALUES::::::
        '::::::::::::::::::::::::::::::
        Try
            Dim pOID As Integer = OID
            'EDIT THE FEATURE
            'GetGeodexLayer is a public function that is called all over this application!
            'It lives in the GDXTools class.  It basically searches for a layer in the map document
            'that matches the name set as the geodex layer (usually "Geodex")
            Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
            Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
            Dim pDataset As IDataset = pFeatureClass
            'Workspace is a fancy word for "geodatabase"... or is it the other way around... 
            Dim pWorkspace As IWorkspace = pDataset.Workspace

            'Iworkspaceedit is a class that is used to make edits to the workspace
            'You basically just pass it an IWorkspace
            Dim pWorkspaceEdit As IWorkspaceEdit = pWorkspace

            'HERES WHERE YOU MAKE THE pRECORD
            'pRecord has a bunch of "properties" that correspond to fields in geodex
            'It acts as a control where you can pass the values to the object in the form of
            'properties and then take those properties and insert them into the table.
            'It seems convoluted but it works really well!  Think of it like
            'writing down all your answers in pencil before you etch it in stone.
            Dim pRecord As Record = New Record

            'Each of the below little subsections sets a property of the Record Object, pRecord

            'GDX_FILE
            If Not cmbFile.SelectedItem Is "" Then
                Dim pSelectedSeries As Series = cmbFile.SelectedItem
                pRecord.GDX_FILE = pSelectedSeries.GdxCode
            Else
                MsgBox("Please select a Series (Geodex File) to save the entry under.  This is a required field." & vbCrLf & "ERROR CODE 121")
            End If

            'GDX_NUM
            If Not cmbFile.SelectedItem Is "" Then
                Dim pSelectedSeries As Series = cmbFile.SelectedItem
                Dim fullstring As String = pSelectedSeries.ToString().Trim()
                Dim start As Integer = fullstring.Length() - 6
                Dim partstring As String = fullstring.Substring(start).TrimEnd(")")
                pRecord.GDX_NUM = partstring
            End If

            'GDX_SUB
            If Not cmb_subtype.SelectedItem Is Nothing Then
                Dim pSelectedSub As GDX_SUB = cmb_subtype.SelectedItem
                Try
                    pRecord.GDX_SUB = pSelectedSub.Code
                Catch ex As Exception
                    MsgBox(ex.ToString())
                End Try
            End If

            ''RECORD
            If Not txtRecord.Text.Trim() = "" Then
                pRecord.RECORD = txtRecord.Text.Trim()
            End If

            ''LOCATION
            If Not txtLocation.Text.Trim() = "" Then
                pRecord.LOCATION = txtLocation.Text.Trim()
            End If

            ''SERIES_TIT
            If Not txtSeriesTit.Text.Trim() = "" Then
                pRecord.SERIES_TIT = txtSeriesTit.Text.Trim()
            End If

            ''PUBLISHER
            If Not txtPublisher.Text.Trim() = "" Then
                pRecord.PUBLISHER = txtPublisher.Text.Trim()
            End If

            ''MAP_TYPE
            If Not cmbMapType.SelectedItem Is Nothing Then
                Dim pSelectedMapType As MapType = cmbMapType.SelectedItem
                pRecord.MAP_TYPE = pSelectedMapType.Code
            End If

            ''PRODUCTION
            If Not cmbProduction.SelectedItem Is Nothing Then
                Dim pSelectedProduction As Production = cmbProduction.SelectedItem
                pRecord.PRODUCTION = pSelectedProduction.Code
            End If

            ''MAP_FOR
            If Not cmbFormat.SelectedItem Is Nothing Then
                Dim pSelectedFormat As Format = cmbFormat.SelectedItem
                pRecord.MAP_FOR = pSelectedFormat.Code
            End If

            'PROJECT
            If Not cmbProjection.SelectedItem Is Nothing Then
                Dim pSelectedProjection As Projection = cmbProjection.SelectedItem
                pRecord.PROJECT = pSelectedProjection.Code
            End If

            'PRIME_MER
            If Not cmbPrimeMer.SelectedItem Is Nothing Then
                Dim pSelectedPrimeMer As PrimeMer = cmbPrimeMer.SelectedItem
                pRecord.PRIME_MER = pSelectedPrimeMer.Code
            End If

            ''SCALE
            'ERROR CODE 122: Value for scale must be numeric
            If Not txtScale.Text = "" Then
                Dim pScale As Long = txtScale.Text()
                Try
                    pRecord.SCALE = pScale
                Catch ex As Exception
                    MsgBox("Value for ""SCALE"" must contain numeric digits only!" & vbCrLf & "ERROR CODE 122" & ex.ToString(), MsgBoxStyle.OkOnly)
                    Exit Sub
                End Try
            End If

            ''CATLOC
            If Not txtCatalog.Text.Trim() = "" Then
                pRecord.CATLOC = txtCatalog.Text.Trim()
            End If

            ''HOLD

            Dim pHold As Short = 20 - dudHold.SelectedIndex
            pRecord.HOLD = pHold

            ''Edition
            If Not txtEdition.Text.Trim() = "" Then
                Try
                    pRecord.EDITION_NO = txtEdition.Text.Trim()
                Catch ex As Exception
                    MsgBox("Value for ""Edition Number"" must contain numeric digits only!" & vbCrLf & "ERROR CODE 123" & ex.ToString(), MsgBoxStyle.OkOnly)
                    Exit Sub
                End Try
            End If

            ''ISO_TYPE
            If Not cmbIsoType.SelectedItem Is Nothing Then
                Dim pSelectedIso As IsoType = cmbIsoType.SelectedItem
                Dim sISO_TYPE As Short = pSelectedIso.Code
                pRecord.ISO_TYPE = sISO_TYPE
            End If

            'ISO_VAL
            If Not txtIsoVal.Text = "" Then
                Dim sISO_VAL As Short = txtIsoVal.Text.Trim()
                Try
                    pRecord.ISO_VAL = sISO_VAL
                Catch ex As Exception
                    MsgBox("Value for ""Isobar Value"" must contain numeric digits only!" & vbCrLf & "ERROR CODE 124" & ex.ToString(), MsgBoxStyle.OkOnly)
                    Exit Sub
                End Try
            End If

            'LAT_DIMEN
            '::::::::::::::::::::::::::::
            '::LAT AND LON DIMEN STUFF!::
            '::::::::::::::::::::::::::::
            'Not sure if this is even going to be used at all.  If so, just
            'make these guys functional:
            pRecord.LAT_DIMEN = Nothing
            pRecord.LON_DIMEN = Nothing

            'RUN_DATE
            '::::::::::::
            'Not sure if this works yet...
            '::::::::::::
            
            pRecord.RUN_DATE = DateAndTime.DateString() & My.User.Name.ToString()


            'YEARS AND YEAR TYPES
            '1. Get the values from the year and year type controls
            '2. Link the corresponding year and year types in a list of tuples (1 tuple for each year/yeartype combo)
            '3. Sort a list of the years from high to low
            '4. Make a list of objects with properties i. year and ii. year type
            '5. Assign the objects, in order, to the variables so that Year 1 is the most recent year
            '6. ??????????
            '7. PROFIT

            Dim year1 As Short = Nothing
            Dim year2 As Short = Nothing
            Dim year3 As Short = Nothing
            Dim year4 As Short = Nothing

            Dim sYear1 As Short = Nothing
            Dim sYear2 As Short = Nothing
            Dim sYear3 As Short = Nothing
            Dim sYear4 As Short = Nothing

            'get the years
            'ERROR CODE 125: Not a valid entry for the year!

            'Let me explain for year 1:

            'If the year 1 text box is not empty then
            If Not txtYear1.Text = "" Then
                'Make a new string variable that contains the text from the box
                Dim sTxtYear1 As String = txtYear1.Text.Trim()
                'Make sure the length is 4 (to top lazy people from using two digit years.
                'DOESNT ANYONE REMEMBER Y2K???
                If sTxtYear1.Length <> 4 Then
                    MsgBox("Please enter a valid year for Year 1" & vbCrLf & "ERROR CODE 125")
                    Exit Sub
                End If
                'Try to covert the string to an integer.  If this doesn't work, the user probably put
                'lettesr or special characters in the text box.
                Try
                    year1 = Convert.ToInt16(sTxtYear1)
                Catch ex As Exception
                    MsgBox("Please enter a valid year for Year 1" & vbCrLf & "ERROR CODE 125" & vbCrLf & ex.ToString())
                    Exit Sub
                End Try
            End If
            'Voila.  If the user didn't put in invalid values, the variable "year1" has been assigned
            'a year value from the form.

            'The same thing happens for year 2, 3, and 4.   Remember the first if will catch them if they are not used.
            If Not txtYear2.Text = "" Then
                Dim sTxtYear2 As String = txtYear2.Text.Trim()
                If sTxtYear2.Length <> 4 Then
                    MsgBox("Please enter a valid year for Year 2" & vbCrLf & "ERROR CODE 125")
                    Exit Sub
                End If
                Try
                    year2 = Convert.ToInt16(sTxtYear2)
                Catch ex As Exception
                    MsgBox("Please enter a valid year for Year 2" & vbCrLf & "ERROR CODE 125" & vbCrLf & ex.ToString())
                    Exit Sub
                End Try
            End If

            If Not txtYear3.Text = "" Then
                Dim sTxtYear3 As String = txtYear3.Text.Trim()
                If sTxtYear3.Length <> 4 Then
                    MsgBox("Please enter a valid year for Year 3" & vbCrLf & "ERROR CODE 125")
                    Exit Sub
                End If
                Try
                    year3 = Convert.ToInt16(sTxtYear3)
                Catch ex As Exception
                    MsgBox("Please enter a valid year for Year 3" & vbCrLf & "ERROR CODE 125" & vbCrLf & ex.ToString())
                    Exit Sub
                End Try
            End If

            If Not txtYear4.Text = "" Then
                Dim sTxtYear4 As String = txtYear4.Text.Trim()
                If sTxtYear4.Length <> 4 Then
                    MsgBox("Please enter a valid year for Year 4" & vbCrLf & "ERROR CODE 125")
                    Exit Sub
                End If
                Try
                    year4 = Convert.ToInt16(sTxtYear4)
                Catch ex As Exception
                    MsgBox("Please enter a valid year for Year 4" & vbCrLf & "ERROR CODE 125" & vbCrLf & ex.ToString())
                    Exit Sub
                End Try
            End If


            'NOW we need to get the year types for each year.  We haven't changed the order yet
            'so year1 will be associated with the type of year1.  Again, I'll explain for year1type
            'and you can figure out the other 3.

            'If statement makes sure that something is selected in the box, if not, it's going to be assinged
            'a nothing value (which is AOK)
            If Not cmbYear1.SelectedItem Is Nothing Then
                'Make an OBJECT of type YearTYpe.
                'the YearType object has two properties, year and year type.
                'Here, we just want the "CODE" that's used in the domain for year type.
                Dim pSelectedYear1 As YearType = cmbYear1.SelectedItem
                sYear1 = pSelectedYear1.Code
            End If
            'Voila, we have assined the year type code to the variable sYear1.  SO...


            'Now we have two variables for Year 1:
            'year1 as a short integer that is the actual year value (ie. 1999)
            'year2 as a short integer that is the code used for the domain for year type (ie. 98 for pub_date)



            If Not cmbYear2.SelectedItem Is Nothing Then
                Dim pSelectedYear2 As YearType = cmbYear2.SelectedItem
                sYear2 = pSelectedYear2.Code
            End If

            If Not cmbYear3.SelectedItem Is Nothing Then
                Dim pSelecteYear3 As YearType = cmbYear3.SelectedItem
                sYear3 = pSelecteYear3.Code
            End If

            If Not cmbYear4.SelectedItem Is Nothing Then
                Dim pSelecteYear4 As YearType = cmbYear4.SelectedItem
                sYear4 = pSelecteYear4.Code
            End If


            'define a dictionary of key: year value: year type (both shorts) NOPE
            'INSTEAD define a list of two item tuples, the list of tuples can be sorted
            'unlike a dictionary
            Dim ltYearListofTyples As List(Of Tuple(Of Short, Short)) = New List(Of Tuple(Of Short, Short))

            If Not year1 = Nothing And Not sYear1 = Nothing Then
                Dim tYear1 = Tuple.Create(year1, sYear1)
                ltYearListofTyples.Add(tYear1)
            ElseIf Not year1 = Nothing Xor Not sYear1 = Nothing Then
                MsgBox("Years and year types must be paired in the entry boxes (Year 1)")
                Exit Sub
            End If

            If Not year2 = Nothing And Not sYear2 = Nothing Then
                Dim tYear2 = Tuple.Create(year2, sYear2)
                ltYearListofTyples.Add(tYear2)
            ElseIf Not year2 = Nothing Xor Not sYear2 = Nothing Then
                MsgBox("Years and year types must be paired in the entry boxes (Year 2)")
                Exit Sub
            End If
            If Not year3 = Nothing And Not sYear3 = Nothing Then
                Dim tYear3 = Tuple.Create(year3, sYear3)
                ltYearListofTyples.Add(tYear3)
            ElseIf Not year3 = Nothing Xor Not sYear3 = Nothing Then
                MsgBox("Years and year types must be paired in the entry boxes (Year 3)")
                Exit Sub
            End If
            If Not year4 = Nothing And Not sYear4 = Nothing Then
                Dim tYear4 = Tuple.Create(year4, sYear4)
                ltYearListofTyples.Add(tYear4)
            ElseIf Not year4 = Nothing Xor Not sYear4 = Nothing Then
                MsgBox("Years and year types must be paired in the entry boxes (Year 4)")
                Exit Sub
            End If

            ltYearListofTyples.Sort()
            ltYearListofTyples.Reverse()

            'get values from tuples
            'assign in order.

            Dim listcount As Integer = ltYearListofTyples.Count()
            Dim i As Integer = 0
            Dim lYearObjectList As List(Of Object) = New List(Of Object)

            For i = 0 To listcount - 1
                Dim pCountYear As CountYear = New CountYear
                pCountYear.Year = ltYearListofTyples(i).Item1
                pCountYear.YearType = ltYearListofTyples(i).Item2
                lYearObjectList.Add(pCountYear)
            Next

            'FINALLY assigning the values for year and yeartype to the precord class before editing
            'If there are no years the "nothings" below will not be changed.
            pRecord.YEAR1 = Nothing
            pRecord.YEAR2 = Nothing
            pRecord.YEAR3 = Nothing
            pRecord.YEAR4 = Nothing
            pRecord.YEAR1_TYPE = Nothing
            pRecord.YEAR2_TYPE = Nothing
            pRecord.YEAR3_TYPE = Nothing
            pRecord.YEAR4_TYPE = Nothing

            If listcount > 0 Then
                Dim year1object As CountYear = lYearObjectList(0)
                pRecord.YEAR1 = year1object.Year
                pRecord.YEAR1_TYPE = year1object.YearType
            End If

            If listcount > 1 Then
                Dim year2object As CountYear = lYearObjectList(1)
                pRecord.YEAR2 = year2object.Year
                pRecord.YEAR2_TYPE = year2object.YearType
            End If

            If listcount > 2 Then
                Dim year3object As CountYear = lYearObjectList(2)
                pRecord.YEAR3 = year3object.Year
                pRecord.YEAR3_TYPE = year3object.YearType
            End If

            If listcount > 3 Then
                Dim year4object As CountYear = lYearObjectList(3)
                pRecord.YEAR4 = year4object.Year
                pRecord.YEAR4_TYPE = year4object.YearType
            End If


            'DATE
            'Date will automatically be the most recent year.
            'If there are no years, date will be 0
            If Not pRecord.YEAR1 = Nothing Then
                pRecord.rDATE = pRecord.YEAR1
            Else
                pRecord.rDATE = 0
            End If


            '::::::::::::::::::::::::::::::
            ':::::EDITING THE FEATURE::::::
            '::::::::::::::::::::::::::::::
            'Error code 130: Error when editing the features.  Values sucessfully retreived from the
            'form but there was an error during the workspace editing session or getting the values from
            'the pRecord class.
            Try
                'Here begins the actual editing of the feature with the values prepared above.
                pWorkspaceEdit.StartEditing(True)
                pWorkspaceEdit.StartEditOperation()

                Dim pFeature As IFeature = pFeatureClass.GetFeature(pOID)

                If Not pRecord.GDX_FILE = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("GDX_FILE")) = pRecord.GDX_FILE
                Else
                    pFeature.Value(pFeature.Fields.FindField("GDX_FILE")) = 0
                End If

                If Not pRecord.GDX_NUM Is Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("GDX_NUM")) = pRecord.GDX_NUM
                Else
                    pFeature.Value(pFeature.Fields.FindField("GDX_NUM")) = DBNull.Value
                End If

                If Not pRecord.GDX_SUB = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("GDX_SUB")) = pRecord.GDX_SUB
                Else
                    pFeature.Value(pFeature.Fields.FindField("GDX_SUB")) = DBNull.Value
                End If

                If Not pRecord.RECORD Is Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("RECORD")) = pRecord.RECORD
                Else
                    pFeature.Value(pFeature.Fields.FindField("RECORD")) = DBNull.Value
                End If

                If Not pRecord.LOCATION Is Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("LOCATION")) = pRecord.LOCATION
                Else
                    pFeature.Value(pFeature.Fields.FindField("LOCATION")) = DBNull.Value
                End If

                If Not pRecord.rDATE = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("DATE")) = pRecord.rDATE
                Else
                    pFeature.Value(pFeature.Fields.FindField("DATE")) = 0
                End If

                If Not pRecord.SERIES_TIT Is Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("SERIES_TIT")) = pRecord.SERIES_TIT
                Else
                    pFeature.Value(pFeature.Fields.FindField("SERIES_TIT")) = DBNull.Value
                End If

                If Not pRecord.PUBLISHER Is Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("PUBLISHER")) = pRecord.PUBLISHER
                Else
                    pFeature.Value(pFeature.Fields.FindField("PUBLISHER")) = DBNull.Value
                End If

                If Not pRecord.MAP_TYPE = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("MAP_TYPE")) = pRecord.MAP_TYPE
                Else
                    pFeature.Value(pFeature.Fields.FindField("MAP_TYPE")) = DBNull.Value
                End If

                If Not pRecord.PRODUCTION = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("PRODUCTION")) = pRecord.PRODUCTION
                Else
                    pFeature.Value(pFeature.Fields.FindField("PRODUCTION")) = DBNull.Value
                End If

                If Not pRecord.MAP_FOR = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("MAP_FOR")) = pRecord.MAP_FOR
                Else
                    pFeature.Value(pFeature.Fields.FindField("MAP_FOR")) = DBNull.Value
                End If

                If Not pRecord.PROJECT = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("PROJECT")) = pRecord.PROJECT
                Else
                    pFeature.Value(pFeature.Fields.FindField("PROJECT")) = DBNull.Value
                End If

                If Not pRecord.PRIME_MER = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("PRIME_MER")) = pRecord.PRIME_MER
                Else
                    pFeature.Value(pFeature.Fields.FindField("PRIME_MER")) = DBNull.Value
                End If

                If Not pRecord.SCALE = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("SCALE")) = pRecord.SCALE
                Else
                    pFeature.Value(pFeature.Fields.FindField("SCALE")) = DBNull.Value
                End If

                If Not pRecord.CATLOC Is Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("CATLOC")) = pRecord.CATLOC
                Else
                    pFeature.Value(pFeature.Fields.FindField("CATLOC")) = DBNull.Value
                End If


                pFeature.Value(pFeature.Fields.FindField("HOLD")) = pRecord.HOLD

                'pFeature.Value(pFeature.Fields.FindField("HOLD")) = DBNull.Value


                If Not pRecord.YEAR1 = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("YEAR1")) = pRecord.YEAR1
                Else
                    pFeature.Value(pFeature.Fields.FindField("YEAR1")) = DBNull.Value
                End If

                If Not pRecord.YEAR1_TYPE = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("YEAR1_TYPE")) = pRecord.YEAR1_TYPE
                Else
                    pFeature.Value(pFeature.Fields.FindField("YEAR1_TYPE")) = DBNull.Value
                End If


                If Not pRecord.YEAR2 = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("YEAR2")) = pRecord.YEAR2
                Else
                    pFeature.Value(pFeature.Fields.FindField("YEAR2")) = DBNull.Value
                End If

                If Not pRecord.YEAR2_TYPE = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("YEAR2_TYPE")) = pRecord.YEAR2_TYPE
                Else
                    pFeature.Value(pFeature.Fields.FindField("YEAR2_TYPE")) = DBNull.Value
                End If

                If Not pRecord.YEAR3 = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("YEAR3")) = pRecord.YEAR3
                Else
                    pFeature.Value(pFeature.Fields.FindField("YEAR3")) = DBNull.Value
                End If

                If Not pRecord.YEAR3_TYPE = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("YEAR3_TYPE")) = pRecord.YEAR3_TYPE
                Else
                    pFeature.Value(pFeature.Fields.FindField("YEAR3_TYPE")) = DBNull.Value
                End If

                If Not pRecord.YEAR4 = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("YEAR4")) = pRecord.YEAR4
                Else
                    pFeature.Value(pFeature.Fields.FindField("YEAR4")) = DBNull.Value
                End If

                If Not pRecord.YEAR4_TYPE = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("YEAR4_TYPE")) = pRecord.YEAR4_TYPE
                Else
                    pFeature.Value(pFeature.Fields.FindField("YEAR4_TYPE")) = DBNull.Value
                End If

                If Not pRecord.EDITION_NO = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("EDITION_NO")) = pRecord.EDITION_NO
                Else
                    pFeature.Value(pFeature.Fields.FindField("EDITION_NO")) = DBNull.Value
                End If

                If Not pRecord.ISO_TYPE = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("ISO_TYPE")) = pRecord.ISO_TYPE
                Else
                    pFeature.Value(pFeature.Fields.FindField("ISO_TYPE")) = DBNull.Value
                End If

                If Not pRecord.ISO_VAL = Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("ISO_VAL")) = pRecord.ISO_VAL
                Else
                    pFeature.Value(pFeature.Fields.FindField("ISO_VAL")) = DBNull.Value
                End If

                'Lat and Lon dimensions have not been included yet.  Still not sure what to do with them

                If Not pRecord.LAT_DIMEN Is Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("LAT_DIMEN")) = pRecord.LAT_DIMEN
                Else
                    pFeature.Value(pFeature.Fields.FindField("LAT_DIMEN")) = DBNull.Value
                End If

                'Lat and Lon dimensions have not been included yet.  Still not sure what to do with them
                If Not pRecord.LON_DIMEN Is Nothing Then
                    pFeature.Value(pFeature.Fields.FindField("LON_DIMEN")) = pRecord.LON_DIMEN
                Else
                    pFeature.Value(pFeature.Fields.FindField("LON_DIMEN")) = DBNull.Value
                End If

                pFeature.Value(pFeature.Fields.FindField("RUN_DATE")) = DateAndTime.DateString() & My.User.Name.ToString()

                pFeature.Store()
                pWorkspaceEdit.StopEditOperation()
                pWorkspaceEdit.StopEditing(True)
            Catch ex As Exception
                MsgBox("An error occured during the editing operation." & vbCrLf & ex.ToString())
            End Try


        Catch ex As Exception
            MsgBox("An error occured when editing the feature." & vbCrLf & "ERROR CODE 130" & vbCrLf & ex.ToString(), MsgBoxStyle.OkOnly)
            Exit Sub
        End Try
        lblMSG1.Text = "A new feature has been added with ObjectID: " & My.Settings.CurrentOID.ToString()
        

        My.Settings.CurrentOID = 0
        lblOID.Text = "..."
    End Sub

    ''Copy the feature by the given OID and create a new feature as a copy.  Called on clikcing "Commit and copy"
    'This function runs the entire "Edit Feature" function above, but also copies the feature shape and doesn't clear the form
    Private Sub CopyFeature(ByVal newOID As Integer, ByVal oldOID As Integer)

        'ERROR CODE 131
        Try
            EditFeature(newOID)
        Catch ex As Exception
            MsgBox("there was an error editing the attributes from the old feature to the new feature." & vbCrLf & "Error Code 131")
        End Try

        Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pOldFeature As IFeature = pFeatureClass.GetFeature(oldOID)
        Dim pNewFeature As IFeature = pFeatureClass.GetFeature(newOID)
        Dim pShape As IGeometry = pOldFeature.Shape
        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace
        Dim pWorkspaceEdit As IWorkspaceEdit = pWorkspace
        Dim pRecord As Record = New Record
        pRecord.X1 = pOldFeature.Value(pOldFeature.Fields.FindField("X1"))
        pRecord.X2 = pOldFeature.Value(pOldFeature.Fields.FindField("X2"))
        pRecord.Y1 = pOldFeature.Value(pOldFeature.Fields.FindField("Y1"))
        pRecord.Y2 = pOldFeature.Value(pOldFeature.Fields.FindField("Y2"))

        Try
            pWorkspaceEdit.StartEditing(True)
            pWorkspaceEdit.StartEditOperation()


            pNewFeature.Shape = pShape
            pNewFeature.Value(pNewFeature.Fields.FindField("X1")) = pRecord.X1
            pNewFeature.Value(pNewFeature.Fields.FindField("X2")) = pRecord.X2
            pNewFeature.Value(pNewFeature.Fields.FindField("Y1")) = pRecord.Y1
            pNewFeature.Value(pNewFeature.Fields.FindField("Y2")) = pRecord.Y2

            pNewFeature.Store()
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            If pNewFeature.HasOID = True Then
                Dim pNewOID As Integer = pNewFeature.OID
                lblOID.Text = pNewOID.ToString()
                lblMSG1.Text = "The feature with OID " & oldOID.ToString() & " has been copied and a new feature with OID " & pNewOID.ToString() &
                    " has been created."
                
                My.Settings.CurrentOID = pNewFeature.OID

            Else
                MsgBox("WARNING: The new feature has no Object ID!", MsgBoxStyle.OkOnly)
                lblOID.Text = "..."
                My.Settings.CurrentOID = oldOID
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox("There was an error copying the shape of the feature." & vbCrLf & "ERROR CODE: 132" & vbCrLf & ex.ToString, MsgBoxStyle.OkOnly)
            lblOID.Text = "..."
            Exit Sub
        End Try

    End Sub

    'On clicking the "new" button Makes a new feature in the file with all null values except for OID (19)
    'As one would expect, this creates a new feature.  It doesn't clear the form, so most often
    'this is used if the user loaded the information from an existing feature and just wants to make
    'a new feature with a new OID.
    Private Sub btNew_Click(sender As Object, e As EventArgs) Handles btNew.Click
        Try
            newfeature()
        Catch ex As Exception
            MsgBox("An error occured when creating a new feature." & vbCrLf & ex.ToString())
            Exit Sub
        End Try

        lblNorth.Text = "North"
        lblSouth.Text = "South"
        lblEast.Text = "East"
        lblWest.Text = "West"

        'If you're here it was successful
        'Edited by SRA 12/27/16 to stop from flashing message because I think it was hanging the application
        lblOID.Text = My.Settings.CurrentOID.ToString()
        lblMSG1.Text = "A blank feature has been created with OID " & My.Settings.CurrentOID.ToString() &
            "; if you wish to copy another feature's shape, use the copy button."
        
    End Sub

    'Fill Form Function (FillForm(OID As Integer)) (527)
    'This function takes an object ID and fills the form with the values from the table
    'associated with that object ID.  It is called any time the form needs to be filled
    'with values.
    Private Sub FillForm(ByVal OID As Integer)
        Dim pOID As Integer = OID
        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureclass As IFeatureClass = pLayer.FeatureClass
        Dim pDataset As IDataset = pFeatureclass
        Dim pWorkspace As IWorkspace = pDataset.Workspace
        Dim pWorkspaceEdit As IWorkspaceEdit = pWorkspace
        Dim pFeature As IFeature = pFeatureclass.GetFeature(pOID)

        Dim pRecord As Record = New Record

        Try
            'GDX Series text box:
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("GDX_FILE"))) = False Then
                pRecord.GDX_FILE = pFeature.Value(pFeature.Fields.FindField("GDX_FILE"))
            Else
                pRecord.GDX_FILE = Nothing
            End If

            'Need to make a subtype list and create a series object to select the index based on the string.
            'Convoluted but it works...
            Dim pSubtypes As ISubtypes = pFeatureclass
            If Not pSubtypes Is Nothing Then
                Dim eSubtypes As IEnumSubtype = pSubtypes.Subtypes

                eSubtypes.Reset()

                Dim code As Integer
                Dim sSubtypetext As String

                sSubtypetext = eSubtypes.Next(code)

                'This loop will make new "series" objects based on the subtypes
                'The if loop will check the codes until it finds the one for the loaded oid
                'Once it finds this, it uses the override tostring() output from the series
                'Class to find the index of the selected item in the combo box.
                Do While sSubtypetext <> ""
                    Dim pSeries As Series = New Series
                    pSeries.GdxCode = code
                    pSeries.GdxName = sSubtypetext
                    If pSeries.GdxCode = pRecord.GDX_FILE Then
                        cmbFile.SelectedIndex = cmbFile.FindStringExact(pSeries.ToString())
                        Exit Do
                    End If
                    sSubtypetext = eSubtypes.Next(code)
                Loop
            Else
                cmbFile.SelectedItem = Nothing
            End If

            'GDX_NUM
            'This value is not in the form, not sure if it needs to be here at all
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("GDX_NUM"))) = False Then
                pRecord.GDX_NUM = pFeature.Value(pFeature.Fields.FindField("GDX_NUM"))
            Else
                pRecord.GDX_NUM = Nothing
            End If

            'GDX_SUB
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("GDX_SUB"))) = False Then
                pRecord.GDX_SUB = pFeature.Value(pFeature.Fields.FindField("GDX_SUB"))
            Else
                pRecord.GDX_SUB = Nothing
            End If


            If Not pRecord.GDX_SUB = Nothing Then
                Dim subcode As Integer = 150
                Dim pSUBDomain As IDomain = pSubtypes.Domain(subcode, "GDX_SUB")
                Dim pCodedValueDomain2 As ICodedValueDomain = pSUBDomain

                Dim x As Integer

                For x = 0 To pCodedValueDomain2.CodeCount - 1
                    Dim pSUB As GDX_SUB = New GDX_SUB
                    pSUB.Description = pCodedValueDomain2.Name(x)
                    pSUB.Code = pCodedValueDomain2.Value(x)
                    Dim pSubCode As Short = pSUB.Code
                    If pSubCode = pRecord.GDX_SUB Then
                        cmb_subtype.SelectedIndex = cmb_subtype.FindStringExact(pSUB.Description)
                        Exit For
                    End If
                Next
            Else
                cmb_subtype.SelectedItem = Nothing
            End If

            'RECORD
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("RECORD"))) = False Then
                pRecord.RECORD = pFeature.Value(pFeature.Fields.FindField("RECORD"))
            Else
                pRecord.RECORD = Nothing
            End If

            If Not pRecord.RECORD = Nothing Then
                txtRecord.Text = pRecord.RECORD
            Else
                txtRecord.Clear()
            End If

            'LOCATION
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("LOCATION"))) = False Then
                pRecord.LOCATION = pFeature.Value(pFeature.Fields.FindField("LOCATION"))
            Else
                pRecord.LOCATION = Nothing
            End If

            If Not pRecord.LOCATION Is Nothing Then
                txtLocation.Text = pRecord.LOCATION
            Else
                txtLocation.Clear()
            End If

            'DATE
            'still need to figure out what to do with this text box
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("DATE"))) = False Then
                pRecord.rDATE = pFeature.Value(pFeature.Fields.FindField("DATE"))
            ElseIf IsDBNull(pFeature.Value(pFeature.Fields.FindField("YEAR1"))) = False Then
                pRecord.rDATE = pFeature.Value(pFeature.Fields.FindField("YEAR1"))
            Else
                pRecord.rDATE = Nothing
            End If

            If Not pRecord.rDATE = Nothing Then
                tbP_Date.Text = pRecord.rDATE
            End If

            'Series Title
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("SERIES_TIT"))) = False Then
                pRecord.SERIES_TIT = pFeature.Value(pFeature.Fields.FindField("SERIES_TIT"))
            Else
                pRecord.SERIES_TIT = Nothing
            End If

            If Not pRecord.SERIES_TIT = Nothing Then
                txtSeriesTit.Text = pRecord.SERIES_TIT
            Else
                txtSeriesTit.Clear()
            End If

            'Publisher
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("PUBLISHER"))) = False Then
                pRecord.PUBLISHER = pFeature.Value(pFeature.Fields.FindField("PUBLISHER"))
            Else
                pRecord.PUBLISHER = Nothing
            End If

            If Not pRecord.PUBLISHER = Nothing Then
                txtPublisher.Text = pRecord.PUBLISHER
            Else
                txtPublisher.Clear()
            End If


            'Production
            'DOMAIN
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("PRODUCTION"))) = False Then
                pRecord.PRODUCTION = pFeature.Value(pFeature.Fields.FindField("PRODUCTION"))
            Else
                pRecord.PRODUCTION = Nothing
            End If

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
                        cmbProduction.SelectedIndex = cmbProduction.FindStringExact(pProduction.Description)
                        Exit For
                    End If
                Next
            Else
                cmbProduction.SelectedItem = Nothing
            End If

            'Map_type
            'DOMAIN
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
                        cmbMapType.SelectedIndex = cmbMapType.FindStringExact(pMAP_TYPE.Description)
                        Exit For
                    End If
                Next
            Else
                cmbMapType.SelectedItem = Nothing
            End If

            'Projection
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("PROJECT"))) = False Then
                pRecord.PROJECT = pFeature.Value(pFeature.Fields.FindField("PROJECT"))
            Else
                pRecord.PROJECT = Nothing
            End If

            If Not pRecord.PRODUCTION = Nothing Then
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
                        cmbProjection.SelectedIndex = cmbProjection.FindStringExact(pProjection.Description)
                        Exit For
                    End If
                Next
            Else
                cmbProjection.SelectedItem = Nothing
            End If

            'Prime Merididan
            ' domian
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
                    Dim pProjectCode As Short = pPrimeMer.Code
                    If pProjectCode = pRecord.PRIME_MER Then
                        cmbPrimeMer.SelectedIndex = cmbPrimeMer.FindStringExact(pPrimeMer.Description)
                        Exit For
                    End If
                Next
            Else
                cmbPrimeMer.SelectedItem = Nothing
            End If

            'Map Format
            'DOMAIN
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
                        cmbFormat.SelectedIndex = cmbFormat.FindStringExact(pMapFor.Description)
                        Exit For
                    End If
                Next
            Else
                cmbFormat.SelectedItem = Nothing
            End If

            'SCALE
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("SCALE"))) = False Then
                pRecord.SCALE = pFeature.Value(pFeature.Fields.FindField("SCALE"))
            Else
                pRecord.SCALE = Nothing
            End If

            If Not pRecord.SCALE = Nothing Then
                txtScale.Text = pRecord.SCALE
            Else
                txtScale.Clear()
            End If


            'Catalog Location
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("CATLOC"))) = False Then
                pRecord.CATLOC = pFeature.Value(pFeature.Fields.FindField("CATLOC"))
            Else
                pRecord.CATLOC = Nothing
            End If

            If Not pRecord.CATLOC = Nothing Then
                txtCatalog.Text = pRecord.CATLOC
            Else
                txtCatalog.Clear()
            End If

            'Holding
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("HOLD"))) = False Then
                pRecord.HOLD = pFeature.Value(pFeature.Fields.FindField("HOLD"))
            Else
                pRecord.HOLD = 0
            End If

            If pRecord.HOLD = 0 Then
                dudHold.SelectedIndex = 20
            ElseIf Not pRecord.HOLD = Nothing Then
                dudHold.SelectedIndex = 20 - pRecord.HOLD
            End If

            'Year 1
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("YEAR1"))) = False Then
                pRecord.YEAR1 = pFeature.Value(pFeature.Fields.FindField("YEAR1"))
            Else
                pRecord.YEAR1 = Nothing
            End If

            If Not pRecord.YEAR1 = Nothing Then
                txtYear1.Text = pRecord.YEAR1
            Else
                txtYear1.Clear()
            End If

            'Year 2
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("YEAR2"))) = False Then
                pRecord.YEAR2 = pFeature.Value(pFeature.Fields.FindField("YEAR2"))
            Else
                pRecord.YEAR2 = Nothing
            End If

            If Not pRecord.YEAR2 = Nothing Then
                txtYear2.Text = pRecord.YEAR2
            Else
                txtYear2.Clear()
            End If

            'Year 3
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("YEAR3"))) = False Then
                pRecord.YEAR3 = pFeature.Value(pFeature.Fields.FindField("YEAR3"))
            Else
                pRecord.YEAR3 = Nothing
            End If

            If Not pRecord.YEAR3 = Nothing Then
                txtYear3.Text = pRecord.YEAR3
            Else
                txtYear3.Clear()
            End If

            'Year 4
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("YEAR4"))) = False Then
                pRecord.YEAR4 = pFeature.Value(pFeature.Fields.FindField("YEAR4"))
            Else
                pRecord.YEAR4 = Nothing
            End If

            If Not pRecord.YEAR4 = Nothing Then
                txtYear4.Text = pRecord.YEAR4
            Else
                txtYear4.Clear()
            End If

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
                        cmbYear1.SelectedIndex = cmbYear1.FindStringExact(pYearType.Description)
                        Exit For
                    End If
                Next
            Else
                cmbYear1.Text = ""
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
                        cmbYear2.SelectedIndex = cmbYear2.FindStringExact(pYearType.Description)
                        Exit For
                    End If
                Next
            Else
                cmbYear2.Text = ""
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
                        cmbYear3.SelectedIndex = cmbYear3.FindStringExact(pYearType.Description)
                        Exit For
                    End If
                Next
            Else
                cmbYear3.Text = ""
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
                        cmbYear4.SelectedIndex = cmbYear4.FindStringExact(pYearType.Description)
                        Exit For
                    End If
                Next
            Else
                cmbYear4.Text = ""
            End If

            'Edition No
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("EDITION_NO"))) = False Then
                pRecord.EDITION_NO = pFeature.Value(pFeature.Fields.FindField("EDITION_NO"))
            Else
                pRecord.EDITION_NO = Nothing
            End If

            If Not pRecord.EDITION_NO = Nothing Then
                txtEdition.Text = pRecord.EDITION_NO.ToString()
            Else
                txtEdition.Clear()
            End If

            'ISO_TYPE
            'this is another domain
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("ISO_TYPE"))) = False Then
                pRecord.ISO_TYPE = pFeature.Value(pFeature.Fields.FindField("ISO_TYPE"))
            Else
                pRecord.ISO_TYPE = Nothing
            End If

            If Not pRecord.ISO_TYPE = Nothing Then
                Dim subcode As Integer = 2
                Dim pIsoDomain As IDomain = pSubtypes.Domain(subcode, "ISO_TYPE")
                Dim pCodedValueDomain As ICodedValueDomain = pIsoDomain

                Dim x As Integer
                For x = 0 To pCodedValueDomain.CodeCount - 1
                    Dim pIsoType As IsoType = New IsoType
                    pIsoType.Code = pCodedValueDomain.Value(x)
                    pIsoType.Description = pCodedValueDomain.Name(x)
                    Dim pIsoCode As Short = pIsoType.Code
                    If pIsoCode = pRecord.ISO_TYPE Then
                        cmbIsoType.SelectedIndex = cmbIsoType.FindStringExact(pIsoType.Description)
                        Exit For
                    End If
                Next
            Else
                cmbIsoType.SelectedItem = Nothing
            End If


            'ISO_VAL
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("ISO_VAL"))) = False Then
                pRecord.ISO_VAL = pFeature.Value(pFeature.Fields.FindField("ISO_VAL"))
            Else
                pRecord.ISO_VAL = Nothing
            End If

            If Not pRecord.ISO_VAL = Nothing Then
                txtIsoVal.Text = pRecord.ISO_VAL
            Else
                txtIsoVal.Clear()
            End If

            'Lat and Lon dimensions have not been included yet.  Still not sure what to do with them
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("LAT_DIMEN"))) = False Then
                pRecord.LAT_DIMEN = pFeature.Value(pFeature.Fields.FindField("LAT_DIMEN"))
            Else
                pRecord.LAT_DIMEN = Nothing
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("LON_DIMEN"))) = False Then
                pRecord.LON_DIMEN = pFeature.Value(pFeature.Fields.FindField("LON_DIMEN"))
            Else
                pRecord.LON_DIMEN = Nothing
            End If

            'Setting the values in the Spatial index display window:
            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("X1"))) = False Then
                pRecord.X1 = pFeature.Value(pFeature.Fields.FindField("X1"))
                lblWest.Text = pRecord.X1.ToString()
            Else
                pRecord.X1 = Nothing
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("X2"))) = False Then
                pRecord.X2 = pFeature.Value(pFeature.Fields.FindField("X2"))
                lblEast.Text = pRecord.X2.ToString()
            Else
                pRecord.X2 = Nothing
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("Y1"))) = False Then
                pRecord.Y1 = pFeature.Value(pFeature.Fields.FindField("Y1"))
                lblNorth.Text = pRecord.Y1.ToString()
            Else
                pRecord.Y1 = Nothing
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("Y2"))) = False Then
                pRecord.Y2 = pFeature.Value(pFeature.Fields.FindField("Y2"))
                lblSouth.Text = pRecord.Y2.ToString()
            Else
                pRecord.Y2 = Nothing
            End If


        Catch ex As Exception
            MsgBox("There was an error retreiving the values from the feature: " & vbCrLf & ex.ToString())
            Exit Sub
        End Try
    End Sub

    'This function gets the table from the database as it is passed a name.  The name will be different
    'depending on the type of geodatabases.  For the local database it will be "DEF" and for the 
    'Enterprise GDB it will be Geodex.DBO.DEF
    Function GetTableByName(sTableName As String) As ITable
        Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace
        Dim eEnumDatasets As IEnumDataset = pWorkspace.Datasets(esriDatasetType.esriDTTable)


        eEnumDatasets.Reset()

        Dim pDatasetTable As IDataset = eEnumDatasets.Next()

        'MsgBox("Made it to 1923")
        Do Until pDatasetTable Is Nothing
            'MsgBox("Made it inside the do loop")
            Dim pTableName As String = pDatasetTable.Name
            If pTableName = sTableName Then
                Dim pDatatable As ITable = pDatasetTable
                Return pDatatable
                Exit Do
                'Else
                'MsgBox("The table name did not match, looping to the next table")
            End If
            pDatasetTable = eEnumDatasets.Next()
        Loop

        MsgBox("Table not found")

        Return Nothing
    End Function


    ':::::::::::::::::Major Interactions::::::::::::::::::

    '"COMMIT" commits changes to the record and clears the form for a new record
    'Clicking commit edits the feature with the values in the form and then clears the form
    Public Sub btCommit_Click(sender As Object, e As EventArgs) Handles btCommit.Click
        Dim pOID As Integer = My.Settings.CurrentOID
        If pOID = 0 Then
            MsgBox("The form is not pointing to a feature, either select one or create a new one", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        Try
            EditFeature(pOID)
        Catch ex As Exception
            MsgBox("There was an error in the 'EditFeature(by pOID) function." & vbCrLf & ex.ToString())
            Exit Sub
        End Try
        ClearForm()
        lblMSG1.Text = "The feature with OID: " & pOID.ToString() & " has been edited succesfully!"

        'If you're here the operation was successful.
    End Sub

    '"COMMIT AND COPY" commits changes and creates a new record and fills the form with the values from the old
    Private Sub btCommitCopy_Click(sender As Object, e As EventArgs) Handles btCommitCopy.Click
        Dim pOID As Integer = My.Settings.CurrentOID
        Dim oldOid As Integer = My.Settings.CurrentOID
        My.Settings.oldOID = pOID
        If pOID = 0 Then
            MsgBox("The form is not pointing to a feature, either select one or create a new one", MsgBoxStyle.OkOnly)
            lblOID.Text = "..."
            Exit Sub
        End If
        Try
            EditFeature(pOID)
            'This operation will set the my.settings.currentOID to 0
        Catch ex As Exception
            MsgBox("There was an error in the 'EditFeature(by pOID) function." & vbCrLf & ex.ToString())
            Exit Sub
        End Try
        'If you're here the operation was successful.
        lblMSG1.Text = "The feature with oid " & pOID & " has been created successfuly."
       
        Try
            newfeature()
            'This sets the current OID to the new features OID
        Catch ex As Exception
            MsgBox("There was an error creating a new feature." & vbCrLf & ex.ToString())
            Exit Sub
        End Try

        Try
            CopyFeature(My.Settings.CurrentOID, My.Settings.oldOID)
        Catch ex As Exception
            MsgBox("There was an error editing the new feature with the values of the old feature")
        End Try
        My.Settings.oldOID = 0
        lblOID.Text = My.Settings.CurrentOID.ToString()
        lblMSG1.Text = "The feature with OID " & pOID & " has been created successfuly and a new feature with OID " & My.Settings.CurrentOID.ToString() &
            " has been created."
       
        My.Settings.oldOID = 0

    End Sub

    '"LOAD" Click to Fill the form with values from an existing record as entered in the OID field
    Private Sub btLoad_Click(sender As Object, e As EventArgs) Handles btLoad.Click
        Dim pOID As Integer = 0
        If txtOID.Text = "" Then
            MsgBox("Please enter an ObjectID into the text box to load a feature.")
            Exit Sub
        Else
            pOID = txtOID.Text.Trim()
        End If
        Try
            FillForm(pOID)
        Catch ex As Exception
            MsgBox("There was an error when filling the form:" & vbCrLf & ex.ToString())
            My.Settings.CurrentOID = 0
            Exit Sub
        End Try
        My.Settings.CurrentOID = pOID
        lblOID.Text = pOID.ToString()
        lblMSG1.Text = "The selected features attributes have been loaded but the feature has not been copied nor has the shape been set" &
            vbCrLf & "WARNING if you commit changes now, the feature you loaded will be edited." &
            "  To make a new layer keeping the values in the form, click ""New"" to create a new feature." &
            " Undo is usable when editing records with this tool if you make a mistake."
        
    End Sub

    '"COPY" Click to create a new record that is a copy of another record.
    Private Sub btCopy_Click(sender As Object, e As EventArgs) Handles btCopy.Click
        Dim pOID As Integer = My.Settings.CurrentOID
        If My.Settings.CurrentOID = 0 Then
            MsgBox("Please load a feature to copy by entering an OID or loading a selected feature.")
            Exit Sub
        
        End If

        'Try
        '    'Fills the form with the records form the original, with the exception of spatial fields and shape.
        '    FillForm(pOID)
        'Catch ex As Exception
        '    MsgBox("There was an error when filling the form:" & vbCrLf & ex.ToString())
        '    Exit Sub
        'End Try

        'This part will actually edit the shape and the x and y fields
        Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pFeature As IFeature = pFeatureClass.GetFeature(pOID)
        Dim pShape As IGeometry = pFeature.Shape
        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace
        Dim pWorkspaceEdit As IWorkspaceEdit = pWorkspace
        Dim pRecord As Record = New Record
        pRecord.X1 = pFeature.Value(pFeature.Fields.FindField("X1"))
        pRecord.X2 = pFeature.Value(pFeature.Fields.FindField("X2"))
        pRecord.Y1 = pFeature.Value(pFeature.Fields.FindField("Y1"))
        pRecord.Y2 = pFeature.Value(pFeature.Fields.FindField("Y2"))

        Try
            pWorkspaceEdit.StartEditing(True)
            pWorkspaceEdit.StartEditOperation()

            Dim pNewFeature As IFeature = pFeatureClass.CreateFeature()
            pNewFeature.Shape = pShape
            pNewFeature.Value(pFeature.Fields.FindField("X1")) = pRecord.X1
            pNewFeature.Value(pFeature.Fields.FindField("X2")) = pRecord.X2
            pNewFeature.Value(pFeature.Fields.FindField("Y1")) = pRecord.Y1
            pNewFeature.Value(pFeature.Fields.FindField("Y2")) = pRecord.Y2

            pNewFeature.Store()
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)

            If pNewFeature.HasOID = True Then
                Dim pNewOID As Integer = pNewFeature.OID
                lblOID.Text = pNewOID.ToString()
                lblMSG1.Text = "The feature with OID " & pOID.ToString() & " has been copied and a new feature with OID " & pNewOID.ToString() &
                    " has been created."
               
                My.Settings.CurrentOID = pNewFeature.OID

            Else
                MsgBox("WARNING: The new feature has no Object ID!", MsgBoxStyle.OkOnly)
                lblOID.Text = "..."
                My.Settings.CurrentOID = pOID
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox("There was an error copying the shape of the feature." & vbCrLf & ex.ToString, MsgBoxStyle.OkOnly)
            lblOID.Text = "..."
            Exit Sub
        End Try
    End Sub

    '"EDIT SPATIAL INDEX" opens a new spatial index editor form.
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles btEdit.Click
        Dim pIndexEdit As New frmIndexEdit

        pIndexEdit.ArcMapApplication = _application

        'create the arc map wrapper
        Dim pArcMapApplication As New ArcMapWrapper
        pArcMapApplication.ArcMapApplication = _application

        'Show the form
        pIndexEdit.Show(pArcMapApplication)
    End Sub

    '"DELETE" Delete the currently focused OID
    'There is a message box to confirm that the user wants to delete the feature.
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles btDelete.Click
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
                ClearForm()
            Catch ex As Exception
                MsgBox(ex.ToString(), MsgBoxStyle.OkOnly)
            End Try
            lblMSG1.Text = "The feature with OID " & pOID.ToString & " has been deleted."
            My.Settings.CurrentOID = 0
        ElseIf msg = MsgBoxResult.No Then
            Exit Sub
        Else
            Exit Sub
        End If
    End Sub

    'On clicking "set checked to defaults"
    'the functions in this sub will actually write a new record
    'to the default table under the gdx file number
    'if an entry for the series already exists, it will confirm that you want to replace it
    Private Sub cmdDefault_Click(sender As Object, e As EventArgs) Handles cmdDefault.Click
        Dim pTable As ITable = Nothing
        Dim pSeries As Series = New Series
        pSeries = cmbFile.SelectedItem

        Dim pFLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()
        Dim pFeatureClass As IFeatureClass = pFLayer.FeatureClass
        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace
        Dim pSQLSyntax As ISQLSyntax = pWorkspace


        Dim ql As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierPrefix)
        Dim qr As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierSuffix)

        Try
            pTable = GetTableByName(My.Settings.DefaultTable)
            If pTable Is Nothing Then
                MsgBox("pTable is nothing")
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            Exit Sub
        End Try
        'If you're here the table was successfully retreived

        Try

            'Now I need something to check if that series already has an entry in the table!
            Dim pQueryFilter As IQueryFilter = New QueryFilter

            pQueryFilter.WhereClause = ql & "GDX_FILE" & qr & " = " & pSeries.GdxCode()

            Dim pCursor1 As ICursor = pTable.Search(pQueryFilter, True)

            Dim pRow1 As IRow = pCursor1.NextRow()
            Dim pRow As IRow = Nothing

            If Not pRow1 Is Nothing Then
                Dim pMsgBoxResponse As MsgBoxResult = MsgBox("There is a row with that value, would you like to replace the existing defaults?", MsgBoxStyle.YesNo)
                If pMsgBoxResponse = MsgBoxResult.No Then
                    Exit Sub
                ElseIf pMsgBoxResponse = MsgBoxResult.Yes Then
                    pRow = pRow1
                Else
                    Exit Sub
                End If
            Else
                pRow = pTable.CreateRow()
            End If

            Dim pSelectedSeries As Series = cmbFile.SelectedItem

            pRow.Value(pRow.Fields.FindField("GDX_FILE")) = pSelectedSeries.GdxCode

            If cbCatalog.Checked = True Then
                pRow.Value(pRow.Fields.FindField("CATLOC")) = txtCatalog.Text.Trim()
            End If

            If cbSeries.Checked = True Then
                pRow.Value(pRow.Fields.FindField("SERIES_TIT")) = txtSeriesTit.Text.Trim()
            End If

            If cbPublisher.Checked = True Then
                pRow.Value(pRow.Fields.FindField("PUBLISHER")) = txtPublisher.Text.Trim()
            End If

            If cbScale.Checked = True Then
                pRow.Value(pRow.Fields.FindField("SCALE")) = txtScale.Text.Trim()
            End If

            'DOMAIN!
            If cbMapType.Checked = True Then
                If cmbMapType.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("MAP_TYPE")) = DBNull.Value
                Else
                    Dim pMapType As MapType = New MapType
                    pMapType = cmbMapType.SelectedItem
                    pRow.Value(pRow.Fields.FindField("MAP_TYPE")) = pMapType.Code
                End If
            End If

            'DOMAIN!
            If cbProduction.Checked = True Then
                If cmbProduction.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("PRODUCTION")) = DBNull.Value
                Else
                    Dim pProduction As Production = New Production
                    pProduction = cmbProduction.SelectedItem
                    pRow.Value(pRow.Fields.FindField("PRODUCTION")) = pProduction.Code
                End If
            End If

            'DOMAIN!
            If cbPrimMer.Checked = True Then
                If cmbPrimeMer.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("PRIME_MER")) = DBNull.Value
                Else
                    Dim pPrimeMer As PrimeMer = New PrimeMer
                    pPrimeMer = cmbPrimeMer.SelectedItem
                    pRow.Value(pRow.Fields.FindField("PRIME_MER")) = pPrimeMer.Code
                End If
            End If

            'DOMAIN!
            If cbFormat.Checked = True Then
                If cmbFormat.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("MAP_FOR")) = DBNull.Value
                Else
                    Dim pFormat As Format = New Format
                    pFormat = cmbFormat.SelectedItem
                    pRow.Value(pRow.Fields.FindField("MAP_FOR")) = pFormat.Code
                End If
            End If

            If cbProjection.Checked = True Then
                If cmbProjection.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("PROJECT")) = DBNull.Value
                Else
                    Dim pProject As Projection = New Projection
                    pProject = cmbProjection.SelectedItem
                    pRow.Value(pRow.Fields.FindField("PROJECT")) = pProject.Code
                End If
            End If

            If cb_iso_int.Checked = True Then
                pRow.Value(pRow.Fields.FindField("ISO_VAL")) = txtIsoVal.Text.Trim()
            End If

            'DOMAIN!
            If cb_iso_type.Checked = True Then
                If cmbIsoType.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("ISO_TYPE")) = DBNull.Value
                Else
                    Dim pIsoType As IsoType = New IsoType
                    pIsoType = cmbIsoType.SelectedItem
                    pRow.Value(pRow.Fields.FindField("ISO_TYPE")) = pIsoType.Code
                End If
            End If

            If cbYear1.Checked = True Then
                If cmbYear1.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("YEAR1_TYPE")) = DBNull.Value
                Else
                    Dim pYear1Type As YearType = New YearType
                    pYear1Type = cmbYear1.SelectedItem
                    pRow.Value(pRow.Fields.FindField("YEAR1_TYPE")) = pYear1Type.Code
                End If
            End If

            If cbYear2.Checked = True Then
                If cmbYear2.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("YEAR2_TYPE")) = DBNull.Value
                Else
                    Dim pYear2Type As YearType = New YearType
                    pYear2Type = cmbYear2.SelectedItem
                    pRow.Value(pRow.Fields.FindField("YEAR2_TYPE")) = pYear2Type.Code
                End If
            End If

            If cbYear3.Checked = True Then
                If cmbYear3.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("YEAR3_TYPE")) = DBNull.Value
                Else
                    Dim pYear3Type As YearType = New YearType
                    pYear3Type = cmbYear3.SelectedItem
                    pRow.Value(pRow.Fields.FindField("YEAR3_TYPE")) = pYear3Type.Code
                End If
            End If

            If cbYear4.Checked = True Then
                If cmbYear4.SelectedItem Is Nothing Then
                    pRow.Value(pRow.Fields.FindField("YEAR4_TYPE")) = DBNull.Value
                Else
                    Dim pYear4Type As YearType = New YearType
                    pYear4Type = cmbYear4.SelectedItem
                    pRow.Value(pRow.Fields.FindField("YEAR4_TYPE")) = pYear4Type.Code
                End If
            End If

            If cbyearval1.Checked = True Then
                If txtYear1.Text = "" Then
                    pRow.Value(pRow.Fields.FindField("YEAR1")) = DBNull.Value
                Else
                    pRow.Value(pRow.Fields.FindField("YEAR1")) = txtYear1.Text.Trim()
                End If
            End If

            If cbyearval2.Checked = True Then
                If txtYear2.Text = "" Then
                    pRow.Value(pRow.Fields.FindField("YEAR2")) = DBNull.Value
                Else
                    pRow.Value(pRow.Fields.FindField("YEAR2")) = txtYear2.Text.Trim()
                End If
            End If

            If cbyearval3.Checked = True Then
                If txtYear3.Text = "" Then
                    pRow.Value(pRow.Fields.FindField("YEAR3")) = DBNull.Value
                Else
                    pRow.Value(pRow.Fields.FindField("YEAR3")) = txtYear3.Text.Trim()
                End If
            End If

            If cbyearval4.Checked = True Then
                If txtYear4.Text = "" Then
                    pRow.Value(pRow.Fields.FindField("YEAR4")) = DBNull.Value
                Else
                    pRow.Value(pRow.Fields.FindField("YEAR4")) = txtYear4.Text.Trim()
                End If
            End If
         

            pRow.Store()
        Catch ex As Exception
            MsgBox("There was an error storing the values in the default table.")
            Exit Sub
        End Try

        lblMSG1.Text = "The default values for " & pSeries.GdxName & " has been successfully added to the default table." &
               vbCrLf & "(Series #: " & pSeries.GdxCode & ")"
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = ""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = "The default values for " & pSeries.GdxName & " has been successfully added to the default table." &
               vbCrLf & "(Series #: " & pSeries.GdxCode & ")"
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = ""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = "The default values for " & pSeries.GdxName & " has been successfully added to the default table." &
               vbCrLf & "(Series #: " & pSeries.GdxCode & ")"
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = ""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = "The default values for " & pSeries.GdxName & " has been successfully added to the default table." &
               vbCrLf & "(Series #: " & pSeries.GdxCode & ")"
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = ""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = "The default values for " & pSeries.GdxName & " has been successfully added to the default table." &
               vbCrLf & "(Series #: " & pSeries.GdxCode & ")"
    End Sub

    'On clicking "Load from selected"
    'Load values from the table selected in the table
    'Return error if more or less than 1 record are selected
    Private Sub btLoadSelected_Click(sender As Object, e As EventArgs) Handles btLoadSelected.Click
        
        Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()
        Dim pMXDocument As IMxDocument = _application.Document
        Dim pMap As IMap = pMXDocument.FocusMap
        Dim count As Integer = pMap.SelectionCount
        Dim pFeature As IFeature = Nothing
        Dim pOID As Integer = 0
        Dim pFeatureSelection As IFeatureSelection = pFeatureLayer
        Dim pSelectionSet As ISelectionSet = pFeatureSelection.SelectionSet
        Dim eIDS As IEnumIDs = pSelectionSet.IDs
        If count <> 1 Then
            MsgBox("this function will only work if one feature is selected on the map")
            Exit Sub
        Else
            pOID = eIDS.Next
            My.Settings.CurrentOID = pOID
            Try
                FillForm(pOID)
            Catch ex As Exception
                MsgBox("There was an error when filling the form:" & vbCrLf & ex.ToString())
                Exit Sub
            End Try
            lblOID.Text = pOID.ToString()
            lblMSG1.Text = "The selected features attributes have been loaded but the feature has not been copied nor has the shape been set" &
                "WARNING if you commit changes now, the feature you loaded will be edited." &
                "  To make a new layer keeping the values in the form, click ""New"" to create a new feature." &
                " Undo is usable when editing records with this tool if you make a mistake."
           
        End If

    End Sub

    'When selecting a series (geodex file)
    'ask if the user wants to import the defaults for that file 
    Private Sub cmbFile_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFile.SelectionChangeCommitted
        Dim pSelectedSeries As Series = cmbFile.SelectedItem

        Dim pLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureclass As IFeatureClass = pLayer.FeatureClass
        Dim pSubtypes As ISubtypes = pFeatureclass

        Dim pTable As ITable = Nothing

        Dim pDataset As IDataset = pFeatureclass
        Dim pWorkspace As IWorkspace = pDataset.Workspace

        Dim pSQLSyntax As ISQLSyntax = pWorkspace


        Dim ql As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierPrefix)
        Dim qr As Char = pSQLSyntax.GetSpecialCharacter(esriSQLSpecialCharacters.esriSQL_DelimitedIdentifierSuffix)

        Try
            pTable = GetTableByName(My.Settings.DefaultTable)
            If pTable Is Nothing Then
                MsgBox("pTable is nothing")
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            Exit Sub
        End Try
        'If you're here the table was successfully retreived

        Dim pQueryFilter As IQueryFilter = New QueryFilter
        pQueryFilter.WhereClause = ql & "GDX_FILE" & qr & " = " & pSelectedSeries.GdxCode()

        Dim pCursor As ICursor = pTable.Search(pQueryFilter, True)
        Dim pRow As IRow = pCursor.NextRow()



        If Not pRow Is Nothing Then
            'there is a default entry for the selected series.
            Dim msg As MsgBoxResult = MsgBox("Do you want to import the default values for the series selected?", MsgBoxStyle.YesNo)
            If msg = MsgBoxResult.Yes Then
                'Load the default values
                '1. Make a new record type object
                '2. Populate the object properties
                '3. Populate the Form with values from the object

                Dim pRecord As Record = New Record

                Try
                    'Record
                    If Not pRow.Value(pRow.Fields.FindField("CATLOC")) Is DBNull.Value Then
                        pRecord.CATLOC = pRow.Value(pRow.Fields.FindField("CATLOC"))
                    End If

                    If Not pRecord.CATLOC Is Nothing Then
                        txtCatalog.Text = pRecord.CATLOC
                    End If

                    'Series Title
                    If Not pRow.Value(pRow.Fields.FindField("SERIES_TIT")) Is DBNull.Value Then
                        pRecord.SERIES_TIT = pRow.Value(pRow.Fields.FindField("SERIES_TIT"))
                    End If

                    If Not pRecord.SERIES_TIT Is Nothing Then
                        txtSeriesTit.Text = pRecord.SERIES_TIT
                    End If

                    'Publisher
                    If Not pRow.Value(pRow.Fields.FindField("PUBLISHER")) Is DBNull.Value Then
                        pRecord.PUBLISHER = pRow.Value(pRow.Fields.FindField("PUBLISHER"))
                    End If

                    If Not pRecord.PUBLISHER Is Nothing Then
                        txtPublisher.Text = pRecord.PUBLISHER
                    End If

                    'Scale
                    If Not pRow.Value(pRow.Fields.FindField("SCALE")) Is DBNull.Value Then
                        pRecord.SCALE = pRow.Value(pRow.Fields.FindField("SCALE"))
                    End If

                    If Not pRecord.SCALE = Nothing Then
                        txtScale.Text = pRecord.SCALE
                    End If

                    '####################
                    '#### attributes ####
                    '### with domains ###
                    '####################

                    'Production
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("PRODUCTION"))) = False Then
                        pRecord.PRODUCTION = pRow.Value(pRow.Fields.FindField("PRODUCTION"))
                    Else
                        pRecord.PRODUCTION = Nothing
                    End If

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
                                cmbProduction.SelectedIndex = cmbProduction.FindStringExact(pProduction.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbProduction.SelectedItem = Nothing
                    End If

                    'Map_type
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("MAP_TYPE"))) = False Then
                        pRecord.MAP_TYPE = pRow.Value(pRow.Fields.FindField("MAP_TYPE"))
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
                                cmbMapType.SelectedIndex = cmbMapType.FindStringExact(pMAP_TYPE.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbMapType.SelectedItem = Nothing
                    End If

                    'Projection
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("PROJECT"))) = False Then
                        pRecord.PROJECT = pRow.Value(pRow.Fields.FindField("PROJECT"))
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
                                cmbProjection.SelectedIndex = cmbProjection.FindStringExact(pProjection.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbProjection.SelectedItem = Nothing
                    End If

                    'Prime Merididan
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("PRIME_MER"))) = False Then
                        pRecord.PRIME_MER = pRow.Value(pRow.Fields.FindField("PRIME_MER"))
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
                            Dim pProjectCode As Short = pPrimeMer.Code
                            If pProjectCode = pRecord.PRIME_MER Then
                                cmbPrimeMer.SelectedIndex = cmbPrimeMer.FindStringExact(pPrimeMer.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbPrimeMer.SelectedItem = Nothing
                    End If

                    'Map Format
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("MAP_FOR"))) = False Then
                        pRecord.MAP_FOR = pRow.Value(pRow.Fields.FindField("MAP_FOR"))
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
                                cmbFormat.SelectedIndex = cmbFormat.FindStringExact(pMapFor.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbFormat.SelectedItem = Nothing
                    End If

                    'Year1_Type
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("YEAR1_TYPE"))) = False Then
                        pRecord.YEAR1_TYPE = pRow.Value(pRow.Fields.FindField("YEAR1_TYPE"))
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
                                cmbYear1.SelectedIndex = cmbYear1.FindStringExact(pYearType.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbYear1.Text = ""
                    End If

                    'Year2_Type
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("YEAR2_TYPE"))) = False Then
                        pRecord.YEAR2_TYPE = pRow.Value(pRow.Fields.FindField("YEAR2_TYPE"))
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
                                cmbYear2.SelectedIndex = cmbYear2.FindStringExact(pYearType.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbYear2.Text = ""
                    End If

                    'Year3_Type
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("YEAR3_TYPE"))) = False Then
                        pRecord.YEAR3_TYPE = pRow.Value(pRow.Fields.FindField("YEAR3_TYPE"))
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
                                cmbYear3.SelectedIndex = cmbYear3.FindStringExact(pYearType.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbYear3.Text = ""
                    End If

                    'Year4_Type
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("YEAR4_TYPE"))) = False Then
                        pRecord.YEAR4_TYPE = pRow.Value(pRow.Fields.FindField("YEAR4_TYPE"))
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
                                cmbYear4.SelectedIndex = cmbYear4.FindStringExact(pYearType.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbYear4.Text = ""
                    End If


                    'ISO_TYPE
                    'this is another domain
                    If IsDBNull(pRow.Value(pRow.Fields.FindField("ISO_TYPE"))) = False Then
                        pRecord.ISO_TYPE = pRow.Value(pRow.Fields.FindField("ISO_TYPE"))
                    Else
                        pRecord.ISO_TYPE = Nothing
                    End If

                    If Not pRecord.ISO_TYPE = Nothing Then
                        Dim subcode As Integer = 2
                        Dim pIsoDomain As IDomain = pSubtypes.Domain(subcode, "ISO_TYPE")
                        Dim pCodedValueDomain As ICodedValueDomain = pIsoDomain

                        Dim x As Integer
                        For x = 0 To pCodedValueDomain.CodeCount - 1
                            Dim pIsoType As IsoType = New IsoType
                            pIsoType.Code = x
                            pIsoType.Description = pCodedValueDomain.Name(x)
                            Dim pIsoCode As Short = pIsoType.Code
                            If pIsoCode = pRecord.ISO_TYPE Then
                                cmbIsoType.SelectedIndex = cmbIsoType.FindStringExact(pIsoType.Description)
                                Exit For
                            End If
                        Next
                    Else
                        cmbIsoType.SelectedItem = Nothing
                    End If

                    'Iso_Val
                    If Not pRow.Value(pRow.Fields.FindField("ISO_VAL")) Is DBNull.Value Then
                        pRecord.ISO_VAL = pRow.Value(pRow.Fields.FindField("ISO_VAL"))
                    End If

                    If Not pRecord.ISO_VAL = Nothing Then
                        txtIsoVal.Text = pRecord.ISO_VAL
                    End If


                    If Not pRow.Value(pRow.Fields.FindField("YEAR1")) Is DBNull.Value Then
                        pRecord.YEAR1 = pRow.Value(pRow.Fields.FindField("YEAR1"))
                    End If
                    If Not pRecord.YEAR1 = Nothing Then
                        txtYear1.Text = pRecord.YEAR1
                    End If



                    If Not pRow.Value(pRow.Fields.FindField("YEAR2")) Is DBNull.Value Then
                        pRecord.YEAR2 = pRow.Value(pRow.Fields.FindField("YEAR2"))
                    End If
                    If Not pRecord.YEAR2 = Nothing Then
                        txtYear2.Text = pRecord.YEAR2
                    End If


                    If Not pRow.Value(pRow.Fields.FindField("YEAR3")) Is DBNull.Value Then
                        pRecord.YEAR3 = pRow.Value(pRow.Fields.FindField("YEAR3"))
                    End If
                    If Not pRecord.YEAR3 = Nothing Then
                        txtYear3.Text = pRecord.YEAR3
                    End If

                    If Not pRow.Value(pRow.Fields.FindField("YEAR4")) Is DBNull.Value Then
                        pRecord.YEAR4 = pRow.Value(pRow.Fields.FindField("YEAR4"))
                    End If
                    If Not pRecord.YEAR4 = Nothing Then
                        txtYear4.Text = pRecord.YEAR4
                    End If


                    If Not pRow.Value(pRow.Fields.FindField("EDITION_NO")) Is DBNull.Value Then
                        pRecord.EDITION_NO = pRow.Value(pRow.Fields.FindField("EDITION_NO"))
                    End If

                Catch ex As Exception
                    MsgBox("Error retreiving values fromt he default table." & ex.ToString())
                    Exit Sub
                End Try
                MsgBox("Default values have been loaded to the form!", MsgBoxStyle.OkOnly)

            ElseIf msg = MsgBoxResult.No Then
                'Just keep the fields as they are
                Exit Sub
            Else
                'Do nothing
                Exit Sub
            End If
        Else
            'there is no default
            Exit Sub
        End If
    End Sub


    ':::::::::::::::::Lesser Interactions::::::::::::::::::

    'on clicking select all
    Private Sub cmdSelAll_Click(sender As Object, e As EventArgs) Handles cmdSelAll.Click

        cbCatalog.Checked = True
        cbFormat.Checked = True
        cbMapType.Checked = True
        cbPrimMer.Checked = True
        cbProduction.Checked = True
        cbProjection.Checked = True
        cbPublisher.Checked = True
        cbScale.Checked = True
        cbSeries.Checked = True

        cb_iso_int.Checked = True
        cb_iso_type.Checked = True
        cbYear1.Checked = True
        cbYear2.Checked = True
        cbYear3.Checked = True
        cbYear4.Checked = True
        cbyearval1.Checked = True
        cbyearval2.Checked = True
        cbyearval3.Checked = True
        cbyearval4.Checked = True

    End Sub

    'On clicking select none
    Private Sub cmdSelNone_Click(sender As Object, e As EventArgs) Handles cmdSelNone.Click

        cbCatalog.Checked = False
        cbFormat.Checked = False
        cbMapType.Checked = False
        cbPrimMer.Checked = False
        cbProduction.Checked = False
        cbProjection.Checked = False
        cbPublisher.Checked = False
        cbScale.Checked = False
        cbSeries.Checked = False

        cb_iso_int.Checked = False
        cb_iso_type.Checked = False
        cbYear1.Checked = False
        cbYear2.Checked = False
        cbYear3.Checked = False
        cbYear4.Checked = False
        cbyearval1.Checked = False
        cbyearval2.Checked = False
        cbyearval3.Checked = False
        cbyearval4.Checked = False
    End Sub


    'Clear Form upon clicking "Clear"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btClear.Click
        Try
            ClearForm()
        Catch ex As Exception
            MsgBox(ex.ToString())
            Exit Sub
        End Try
    End Sub

    'Abandon.  Clears the form and clears the current OID.  Doesn't delete anything.
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btAbandon.Click
        Try
            ClearForm()
        Catch ex As Exception
            MsgBox(ex.Message.ToString & vbCrLf & ex.ToString)
        End Try
        My.Settings.CurrentOID = 0
        lblMSG1.Text = "The records have been abandoned, but not deleted.  If you created a new record with ""new"" or ""copy"" you will need to delete it manually or by loading it again and pressing ""Delete"""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = ""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = "The records have been abandoned, but not deleted.  If you created a new record with ""new"" or ""copy"" you will need to delete it manually or by loading it again and pressing ""Delete"""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = ""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = "The records have been abandoned, but not deleted.  If you created a new record with ""new"" or ""copy"" you will need to delete it manually or by loading it again and pressing ""Delete"""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = ""
        System.Threading.Thread.Sleep(200)
        lblMSG1.Text = "The records have been abandoned, but not deleted.  If you created a new record with ""new"" or ""copy"" you will need to delete it manually or by loading it again and pressing ""Delete"""
    End Sub

    'On closing the form
    Private Sub frmRecordViewer_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        My.Settings.northstring = "North"
        My.Settings.southstring = "South"
        My.Settings.eaststring = "East"
        My.Settings.weststring = "West"
        pMxDoc.UpdateContents()
        pMxDoc.ActivatedView.Refresh()
        My.Settings.CurrentOID = 0
    End Sub

    'When clicking the "zoom to" button, the map will zoom in to the feature.  If
    'you want the envelope bigger or smaller, make the penvelope.expand values bigger or smaller
    Private Sub btZoom_Click(sender As Object, e As EventArgs) Handles btZoom.Click
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

    'When clickign the select button: creates a new selection of the current OID
    Private Sub btSelect_Click(sender As Object, e As EventArgs) Handles btSelect.Click
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

    'On clicking "FLASH"
    'Make the focused OID flash on the map.
    Private Sub btFlash_Click(sender As Object, e As EventArgs) Handles btFlash.Click
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

    Private Sub askcommit_Click(sender As Object, e As EventArgs) Handles askcommit.Click
        Dim msgbox_result As MsgBoxResult = MsgBox("Are you sure you want to Commit these changes?", MsgBoxStyle.YesNo)

        If msgbox_result = MsgBoxResult.Yes Then
            btCommit_Click(sender, e)
        Else
            Exit Sub
        End If

    End Sub

    Private Sub frmRecordViewer_Enter(sender As Object, e As EventArgs) Handles MyBase.Enter
        lblNorth.Text = My.Settings.northstring
        lblWest.Text = My.Settings.weststring
        lblEast.Text = My.Settings.eaststring
        lblSouth.Text = My.Settings.southstring
    End Sub

End Class