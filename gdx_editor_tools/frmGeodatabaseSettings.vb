Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.Geometry
Imports System.Windows.Forms



Public Class frmGeodatabaseSettings

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

    'On clicking try
    'This basically tries to connect to the geodatabase by creating an iWorkspace instance
    'if pworspace as iWorspace doesn't exist, it fails.  If it exists, that means we can use the database.
    Private Sub btTry_Click(sender As Object, e As EventArgs) Handles btTry.Click
        If rbSDE.Checked = True Then
            Try
                Dim server As String
                Dim instance As String
                Dim user As String
                Dim password As String
                Dim database As String
                Dim version As String
                Dim authentication As String

                If Not txtServer.Text = "" Then
                    server = txtServer.Text
                Else
                    MsgBox("Please enter a value for ""Server""", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                If Not txtInstance.Text = "" Then
                    instance = txtInstance.Text
                Else
                    MsgBox("Please enter a value for ""Instance""", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                If txtUser.Enabled = True Then
                    If Not txtUser.Text = "" Then
                        user = txtUser.Text
                    Else
                        MsgBox("Please enter a value for ""Instance""", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                Else
                    user = ""
                End If

                If txtUser.Enabled = True Then
                    If Not txtPassword.Text = "" Then
                        password = txtUser.Text
                    Else
                        MsgBox("Please enter a value for ""Password""", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                Else
                    password = ""
                End If

                If Not txtDatabase.Text = "" Then
                    database = txtDatabase.Text
                Else
                    MsgBox("Please enter a value for ""Database""", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                If Not txtVersion.Text = "" Then
                    version = txtVersion.Text
                Else
                    MsgBox("Please enter a value for ""Version""", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                Dim pPropertySet As IPropertySet = New PropertySetClass()
                pPropertySet.SetProperty("SERVER", server)
                pPropertySet.SetProperty("INSTANCE", instance)
                pPropertySet.SetProperty("DATABASE", database)
                pPropertySet.SetProperty("VERSION", version)

                If cbAuth.SelectedIndex = 0 Then
                    authentication = "OSA"
                Else
                    authentication = "DBA"

                    If Not password = "" Then
                        pPropertySet.SetProperty("PASSWORD", password)
                    Else
                        MsgBox("Please enter a value for ""Pasword""", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                    If Not user = "" Then
                        pPropertySet.SetProperty("USER", user)
                    Else
                        MsgBox("Please enter a value for ""User""", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If

                End If
                pPropertySet.SetProperty("AUTHENTICATION_MODE", authentication)

                Dim pWorkspaceFactory As IWorkspaceFactory = New SdeWorkspaceFactory
                Dim pWorkspace As IWorkspace = pWorkspaceFactory.Open(pPropertySet, _application.hWnd)
                Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace

                If pWorkspace.Exists = True Then
                    MsgBox("Successfully connected to the database", MsgBoxStyle.OkOnly)
                Else
                    MsgBox("There was a problem connecting to the database! (before line 118)")
                End If

            Catch ex As Exception
                MsgBox("There was a problem connecting to the database!" & vbCrLf & ex.ToString(), MsgBoxStyle.OkOnly)
                Exit Sub
            End Try
        Else
            Try
                Dim pPath As String = txtPath.Text.ToString().Trim()
                Dim pWorkspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
                Dim pWorkspace As IWorkspace = pWorkspaceFactory.OpenFromFile(pPath, _application.hWnd)
                Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace

                If pWorkspace.Exists = True Then
                    MsgBox("Successfully connected to the database", MsgBoxStyle.OkOnly)
                Else
                    MsgBox("There was a problem connecting to the database!")
                End If
            Catch ex As Exception
                MsgBox("There was a problem connecting to the database!" & vbCrLf & ex.ToString())
                Exit Sub
            End Try
        End If
    End Sub

    'On load
    'make sure that user and password are disabled
    'Check the SDE version by default.
    Private Sub frmGeodatabaseSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtUser.Enabled = False
        txtPassword.Enabled = False
        cbAuth.SelectedIndex = 0
        rbSDE.Checked = True
        rbFGDB.Checked = False
        txtPath.Enabled = False

        txtServer.Text = My.Settings.server
        txtInstance.Text = My.Settings.instance
        cbAuth.SelectedIndex = 0
        txtUser.Text = Nothing
        txtPassword.Text = Nothing
        txtDatabase.Text = My.Settings.database
        txtVersion.Text = My.Settings.version

        txtPath.Text = My.Settings.path

    End Sub

    'On changing the type of authentication.
    'The SDE database in SARUP uses operating system authentication
    'so this is the default and is unlikley to change.
    Private Sub cbAuth_SelectionChangeCommitted(sender As Object, e As EventArgs)
        IIf(cbAuth.SelectedIndex = 1, txtUser.Enabled = True, txtUser.Enabled = False)
        IIf(cbAuth.SelectedIndex = 1, txtPassword.Enabled = True, txtPassword.Enabled = False)

        If cbAuth.SelectedIndex = 1 Then
            txtUser.Enabled = True
            txtPassword.Enabled = True
        ElseIf cbAuth.SelectedIndex = 0 Then
            txtUser.Enabled = False
            txtPassword.Enabled = False
        Else
            MsgBox("!!!")
            Exit Sub
        End If
    End Sub

    'Does the same thing as Try, but loads the data as well.
    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click
        Dim pFeatureWorkspace As IFeatureWorkspace = Nothing
        If rbSDE.Checked = True Then
            Dim sDatabase As String = My.Settings.database
            Dim sDefaultTable As String = sDatabase & ".DBO.DEF"
            My.Settings.DBLocal = False
            Try
                If Not txtServer.Text = "" Then
                    My.Settings.server = txtServer.Text
                Else
                    MsgBox("Please enter a value for ""Server""", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                If Not txtInstance.Text = "" Then
                    My.Settings.instance = txtInstance.Text
                Else
                    MsgBox("Please enter a value for ""Instance""", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                If txtUser.Enabled = True Then
                    If Not txtUser.Text = "" Then
                        My.Settings.user = txtUser.Text
                    Else
                        MsgBox("Please enter a value for ""Instance""", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                Else
                    My.Settings.user = "user"
                End If

                If txtUser.Enabled = True Then
                    If Not txtPassword.Text = "" Then
                        My.Settings.password = txtUser.Text
                    Else
                        MsgBox("Please enter a value for ""Password""", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                Else
                    My.Settings.password = ""
                End If

                If Not txtDatabase.Text = "" Then
                    My.Settings.database = txtDatabase.Text
                Else
                    MsgBox("Please enter a value for ""Database""", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                If Not txtVersion.Text = "" Then
                    My.Settings.version = txtVersion.Text
                Else
                    MsgBox("Please enter a value for ""Version""", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                'The following code will grab the values from the settings and apply them
                'as a Property Set to the server connection.

                Dim server As String = My.Settings.server
                Dim instance As String = My.Settings.instance
                Dim user As String = My.Settings.user
                Dim password As String = My.Settings.password
                Dim database As String = My.Settings.database
                Dim version As String = My.Settings.version
                Dim authentication As String = My.Settings.auth

                Dim pPropertySet As IPropertySet = New PropertySetClass()
                pPropertySet.SetProperty("SERVER", server)
                pPropertySet.SetProperty("INSTANCE", instance)
                pPropertySet.SetProperty("DATABASE", database)
                pPropertySet.SetProperty("VERSION", version)

                If cbAuth.SelectedIndex = 0 Then
                    authentication = "OSA"
                Else
                    authentication = "DBA"

                    If Not password = "" Then
                        pPropertySet.SetProperty("PASSWORD", password)
                    Else
                        MsgBox("Please enter a value for ""Pasword""", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                    If Not user = "" Then
                        pPropertySet.SetProperty("USER", user)
                    Else
                        MsgBox("Please enter a value for ""User""", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If

                End If
                pPropertySet.SetProperty("AUTHENTICATION_MODE", authentication)

                Dim pWorkspaceFactory As IWorkspaceFactory = New SdeWorkspaceFactory
                Dim pWorkspace As IWorkspace = pWorkspaceFactory.Open(pPropertySet, _application.hWnd)
                pFeatureWorkspace = pWorkspace

                If pWorkspace.Exists = True Then
                    MsgBox("Successfully connected to the database", MsgBoxStyle.OkOnly)
                Else
                    MsgBox("There was a problem connecting to the database!")
                    Exit Sub
                End If

            Catch ex As Exception
                MsgBox("There was a problem connecting to the database!" & vbCrLf & ex.ToString(), MsgBoxStyle.OkOnly)
                Exit Sub

            End Try
        Else
            My.Settings.DBLocal = True
            My.Settings.DefaultTable = "DEF"
            Try
                Dim pPath As String = txtPath.Text.ToString().Trim()
                My.Settings.path = pPath

                Dim pWorkspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
                Dim pWorkspace As IWorkspace = pWorkspaceFactory.OpenFromFile(pPath, _application.hWnd)
                pFeatureWorkspace = pWorkspace

                If Not pWorkspace.Exists = True Then
                    MsgBox("There was a problem connecting to the database!")
                End If

            Catch ex As Exception
                MsgBox("There was a problem connecting to the database!" & vbCrLf & ex.ToString())
                Exit Sub
            End Try

        End If



        Dim pMXdoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMXdoc.FocusMap
        Dim FeatureClassName As String = My.Settings.FeatureClassName
        pMap.ClearLayers()


        For i As Integer = 0 To pMap.LayerCount - 1
            Dim pLayer As ILayer = pMap.Layer(i)

            If pLayer.Name = FeatureClassName Then
                MsgBox("The layer """ & FeatureClassName & """ is already present in the document", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
        Next

        Dim pFeatureLayer As IFeatureLayer = New FeatureLayer

        If Not pFeatureWorkspace Is Nothing Then
            Dim pGeodexFeatureClass As IFeatureClass = pFeatureWorkspace.OpenFeatureClass(FeatureClassName)
            pFeatureLayer.Name = FeatureClassName
            pFeatureLayer.FeatureClass = pGeodexFeatureClass
            pMXdoc.FocusMap.AddLayer(pFeatureLayer)

            Me.Close()

            pMXdoc.UpdateContents()
            pMXdoc.ActivatedView.Refresh()
        Else
            MsgBox("There was an error retreiving the feature workspace.")
            Exit Sub
        End If
    End Sub

    'Closes the form wtihout doing anything.
    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
        Me.Close()

    End Sub

    'This makes sure that the ArcMap window stays in focus on closing the form
    Private Sub frmGeodatabaseSettings_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Dim pMxDoc As IMxDocument = _application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        pMxDoc.UpdateContents()
        pMxDoc.ActivatedView.Refresh()
    End Sub

    'This stuff is purely esthetic, just makes the controls for SDE diabled if
    'local is selected.  and Vice Versa
    Private Sub rbSDE_CheckedChanged(sender As Object, e As EventArgs) Handles rbSDE.CheckedChanged
        If rbSDE.Checked = True Then
            txtPath.Enabled = False
            txtDatabase.Enabled = True
            txtInstance.Enabled = True
            txtServer.Enabled = True
            txtVersion.Enabled = True
            cbAuth.Enabled = True
        ElseIf rbSDE.Checked = False Then
            txtPath.Enabled = True
            txtDatabase.Enabled = False
            txtInstance.Enabled = False
            txtServer.Enabled = False
            txtVersion.Enabled = False
            cbAuth.Enabled = False
        End If
    End Sub
    Private Sub rbFGDB_CheckedChanged(sender As Object, e As EventArgs) Handles rbFGDB.CheckedChanged
        If rbFGDB.Checked = True Then
            txtPath.Enabled = True
            txtDatabase.Enabled = False
            txtInstance.Enabled = False
            txtServer.Enabled = False
            txtVersion.Enabled = False
            cbAuth.Enabled = False
        ElseIf rbFGDB.Checked = False Then
            txtPath.Enabled = False
            txtDatabase.Enabled = True
            txtInstance.Enabled = True
            txtServer.Enabled = True
            txtVersion.Enabled = True
            cbAuth.Enabled = True
        End If
    End Sub

    'This happens when you click the big button under the carpet.  If you expand the carpet straight down
    'there is a button that creates a new Geodex feature class.  Took me forever to make, but it was just
    'too buggy and not that useful.  
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim pGXDialog As IGxDialog = New GxDialog
        Dim pGXObjects As IEnumGxObject = Nothing
        Dim bObjectSelected As Boolean
        'Dim pGXWorkspace As IGxDatabase


        'Dim p As igx

        Try
            With pGXDialog
                .AllowMultiSelect = False
                .ButtonCaption = "Add"
                .Title = "Define a Feature Class in a File Geodatabase"

                bObjectSelected = .DoModalSave(_application.hWnd)

            End With

            If bObjectSelected = Nothing Then
                Exit Sub
            End If

            Dim sFeatureClass As String = pGXDialog.Name
            Dim sDir As String = pGXDialog.FinalLocation.FullName
            If Not sDir.EndsWith("\") Then
                sDir = sDir & "\"
            End If

            Dim pPath As String = sDir.Trim("\").ToString()

            MsgBox(pPath)


            Dim pWorkspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
            Dim pWorkspace As IWorkspace = pWorkspaceFactory.OpenFromFile(pPath, _application.hWnd)
            Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
            Dim pWorkspaceDomains As IWorkspaceDomains = pWorkspace
            Dim pWorkspaceSpatial As IWorkspaceSpatialReferenceInfo = pWorkspace
            Dim pSpatialRef As ISpatialReference = pWorkspaceSpatial.SpatialReferenceInfo.Next(3857)

            Dim pFields As IFields = New FieldsClass()
            Dim pFieldsEdit As IFieldsEdit = pFields
            pFieldsEdit.FieldCount_2 = 35


            Dim pOIDfield As IField = New FieldClass()
            Dim pOIDfieldedit As IFieldEdit = pOIDfield
            pOIDfieldedit.Name_2 = "OBJETID"
            pOIDfieldedit.Type_2 = esriFieldType.esriFieldTypeOID
            pOIDfieldedit.AliasName_2 = "Object ID"
            pFieldsEdit.Field_2(0) = pOIDfield

            Dim pShapeField As IField = New FieldClass()
            Dim pShapefieldEdit As IFieldEdit = pShapeField

            pShapefieldEdit.Name_2 = "Shape"
            pShapefieldEdit.AliasName_2 = "Shape"
            pShapefieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry

            Dim pGeomDef As IGeometryDef = New GeometryDef
            Dim pGeomDefEdit As IGeometryDefEdit = pGeomDef

            With pGeomDefEdit
                .GeometryType_2 = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolygon
                .SpatialReference_2 = pSpatialRef
            End With

            pShapefieldEdit.GeometryDef_2 = pGeomDef

            Dim strShapeFieldName As String = pShapeField.Name.ToString()
            pFieldsEdit.Field_2(1) = pShapeField


            Dim pGDX_FILEField As IField = New FieldClass()
            Dim pGDX_FILEFieldEdit As IFieldEdit = pGDX_FILEField
            pGDX_FILEFieldEdit.Name_2 = "GDX_FILE"
            pGDX_FILEFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pGDX_FILEFieldEdit.AliasName_2 = "Geodex Series"
            pFieldsEdit.Field_2(2) = pGDX_FILEField

            Dim pGDX_SUBField As IField = New FieldClass()
            Dim pGDX_SUBFieldEdit As IFieldEdit = pGDX_SUBField
            pGDX_SUBFieldEdit.Name_2 = "GDX_SUB"
            pGDX_SUBFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pGDX_SUBFieldEdit.AliasName_2 = "Geodex Subtype"
            pGDX_SUBFieldEdit.Domain_2 = pWorkspaceDomains.DomainByName("SUBS")
            pFieldsEdit.Field_2(3) = pGDX_SUBField

            Dim pRECORDField As IField = New FieldClass()
            Dim pRECORDFieldEdit As IFieldEdit = pRECORDField
            pRECORDFieldEdit.Name_2 = "RECORD"
            pRECORDFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pRECORDFieldEdit.AliasName_2 = "Record"
            pFieldsEdit.Field_2(4) = pRECORDField

            Dim pLOCATIONField As IField = New FieldClass()
            Dim pLOCATIONFieldEdit As IFieldEdit = pLOCATIONField
            pLOCATIONFieldEdit.Name_2 = "LOCATION"
            pLOCATIONFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pLOCATIONFieldEdit.AliasName_2 = "Location"
            pFieldsEdit.Field_2(5) = pLOCATIONField

            Dim pDATEField As IField = New FieldClass()
            Dim pDATEFieldEdit As IFieldEdit = pDATEField
            pDATEFieldEdit.Name_2 = "DATE"
            pDATEFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pDATEFieldEdit.AliasName_2 = "Date"
            pFieldsEdit.Field_2(6) = pDATEField

            Dim pSERIES_TITField As IField = New FieldClass()
            Dim pSERIES_TITFieldEdit As IFieldEdit = pSERIES_TITField
            pSERIES_TITFieldEdit.Name_2 = "SERIES_TIT"
            pSERIES_TITFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pSERIES_TITFieldEdit.AliasName_2 = "Series Title"
            pFieldsEdit.Field_2(7) = pSERIES_TITField

            Dim pPUBLISHERField As IField = New FieldClass()
            Dim pPUBLISHERFieldEdit As IFieldEdit = pPUBLISHERField
            pPUBLISHERFieldEdit.Name_2 = "PUBLISHER"
            pPUBLISHERFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pPUBLISHERFieldEdit.AliasName_2 = "Publisher"
            pFieldsEdit.Field_2(8) = pPUBLISHERField

            Dim pSCALEField As IField = New FieldClass()
            Dim pSCALEFieldEdit As IFieldEdit = pSCALEField
            pSCALEFieldEdit.Name_2 = "SCALE"
            pSCALEFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
            pSCALEFieldEdit.AliasName_2 = "scale"
            pFieldsEdit.Field_2(9) = pSCALEField

            Dim pMAP_TYPEField As IField = New FieldClass()
            Dim pMAP_TYPEFieldEdit As IFieldEdit = pMAP_TYPEField
            pMAP_TYPEFieldEdit.Name_2 = "MAP_TYPE"
            pMAP_TYPEFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pMAP_TYPEFieldEdit.AliasName_2 = "Map Type"
            pFieldsEdit.Field_2(10) = pMAP_TYPEField

            Dim pPRODUCTIONField As IField = New FieldClass()
            Dim pPRODUCTIONFieldEdit As IFieldEdit = pPRODUCTIONField
            pPRODUCTIONFieldEdit.Name_2 = "PRODUCTION"
            pPRODUCTIONFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pPRODUCTIONFieldEdit.AliasName_2 = "Production"
            pFieldsEdit.Field_2(11) = pPRODUCTIONField

            Dim pMAP_FORField As IField = New FieldClass()
            Dim pMAP_FORFieldEdit As IFieldEdit = pMAP_FORField
            pMAP_FORFieldEdit.Name_2 = "MAP_FOR"
            pMAP_FORFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pMAP_FORFieldEdit.AliasName_2 = "Map Format"
            pFieldsEdit.Field_2(12) = pMAP_FORField

            Dim pPROJECTField As IField = New FieldClass()
            Dim pPROJECTFieldEdit As IFieldEdit = pPROJECTField
            pPROJECTFieldEdit.Name_2 = "PROJECT"
            pPROJECTFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pPROJECTFieldEdit.AliasName_2 = "Projection"
            pFieldsEdit.Field_2(13) = pPROJECTField

            Dim pPRIME_MERField As IField = New FieldClass()
            Dim pPRIME_MERFieldEdit As IFieldEdit = pPRIME_MERField
            pPRIME_MERFieldEdit.Name_2 = "PRIME_MER"
            pPRIME_MERFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pPRIME_MERFieldEdit.AliasName_2 = "Prime Meridian"
            pFieldsEdit.Field_2(14) = pPRIME_MERField

            Dim pCATLOCField As IField = New FieldClass()
            Dim pCATLOCFieldEdit As IFieldEdit = pCATLOCField
            pCATLOCFieldEdit.Name_2 = "CATLOC"
            pCATLOCFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pCATLOCFieldEdit.AliasName_2 = "Catalog Location"
            pFieldsEdit.Field_2(15) = pCATLOCField

            Dim pHOLDField As IField = New FieldClass()
            Dim pHOLDFieldEdit As IFieldEdit = pHOLDField
            pHOLDFieldEdit.Name_2 = "HOLD"
            pHOLDFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pHOLDFieldEdit.AliasName_2 = "Holdings"
            pFieldsEdit.Field_2(16) = pHOLDField


            Dim pYEAR1Field As IField = New FieldClass()
            Dim pYEAR1FieldEdit As IFieldEdit = pYEAR1Field
            pYEAR1FieldEdit.Name_2 = "YEAR1"
            pYEAR1FieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pYEAR1FieldEdit.AliasName_2 = "Year 1"
            pFieldsEdit.Field_2(17) = pYEAR1Field

            Dim pYEAR2Field As IField = New FieldClass()
            Dim pYEAR2FieldEdit As IFieldEdit = pYEAR2Field
            pYEAR2FieldEdit.Name_2 = "YEAR2"
            pYEAR2FieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pYEAR2FieldEdit.AliasName_2 = "Year 2"
            pFieldsEdit.Field_2(18) = pYEAR2Field

            Dim pYEAR3Field As IField = New FieldClass()
            Dim pYEAR3FieldEdit As IFieldEdit = pYEAR3Field
            pYEAR3FieldEdit.Name_2 = "YEAR3"
            pYEAR3FieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pYEAR3FieldEdit.AliasName_2 = "Year 3"
            pFieldsEdit.Field_2(19) = pYEAR3Field

            Dim pYEAR4Field As IField = New FieldClass()
            Dim pYEAR4FieldEdit As IFieldEdit = pYEAR4Field
            pYEAR4FieldEdit.Name_2 = "YEAR4"
            pYEAR4FieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pYEAR4FieldEdit.AliasName_2 = "Year 4"
            pFieldsEdit.Field_2(20) = pYEAR4Field

            Dim pYEAR1_TYPEField As IField = New FieldClass()
            Dim pYEAR1_TYPEFieldEdit As IFieldEdit = pYEAR1_TYPEField
            pYEAR1_TYPEFieldEdit.Name_2 = "YEAR1_TYPE"
            pYEAR1_TYPEFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pYEAR1_TYPEFieldEdit.AliasName_2 = "Year 1 Type"
            pFieldsEdit.Field_2(21) = pYEAR1_TYPEField

            Dim pYEAR2_TYPEField As IField = New FieldClass()
            Dim pYEAR2_TYPEFieldEdit As IFieldEdit = pYEAR2_TYPEField
            pYEAR2_TYPEFieldEdit.Name_2 = "YEAR2_TYPE"
            pYEAR2_TYPEFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pYEAR2_TYPEFieldEdit.AliasName_2 = "Year 2 Type"
            pFieldsEdit.Field_2(22) = pYEAR2_TYPEField

            Dim pYEAR3_TYPEField As IField = New FieldClass()
            Dim pYEAR3_TYPEFieldEdit As IFieldEdit = pYEAR3_TYPEField
            pYEAR3_TYPEFieldEdit.Name_2 = "YEAR3_TYPE"
            pYEAR3_TYPEFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pYEAR3_TYPEFieldEdit.AliasName_2 = "Year 3 Type"
            pFieldsEdit.Field_2(23) = pYEAR3_TYPEField

            Dim pYEAR4_TYPEField As IField = New FieldClass()
            Dim pYEAR4_TYPEFieldEdit As IFieldEdit = pYEAR4_TYPEField
            pYEAR4_TYPEFieldEdit.Name_2 = "YEAR4_TYPE"
            pYEAR4_TYPEFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pYEAR4_TYPEFieldEdit.AliasName_2 = "Year 4 Type"
            pFieldsEdit.Field_2(24) = pYEAR4_TYPEField

            Dim pEDITION_NOField As IField = New FieldClass()
            Dim pEDITION_NOFieldEdit As IFieldEdit = pEDITION_NOField
            pEDITION_NOFieldEdit.Name_2 = "EDITION_NO"
            pEDITION_NOFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pEDITION_NOFieldEdit.AliasName_2 = "Edition No"
            pFieldsEdit.Field_2(25) = pEDITION_NOField

            Dim pISO_TYPEField As IField = New FieldClass()
            Dim pISO_TYPEFieldEdit As IFieldEdit = pISO_TYPEField
            pISO_TYPEFieldEdit.Name_2 = "ISO_TYPE"
            pISO_TYPEFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pISO_TYPEFieldEdit.AliasName_2 = "Isobar Type"
            pFieldsEdit.Field_2(26) = pISO_TYPEField

            Dim pISO_VALField As IField = New FieldClass()
            Dim pISO_VALFieldEdit As IFieldEdit = pISO_VALField
            pISO_VALFieldEdit.Name_2 = "ISO_VAL"
            pISO_VALFieldEdit.Type_2 = esriFieldType.esriFieldTypeSmallInteger
            pISO_VALFieldEdit.AliasName_2 = "Isobar Value"
            pFieldsEdit.Field_2(27) = pISO_VALField

            Dim pLAT_DIMENField As IField = New FieldClass()
            Dim pLAT_DIMENFieldEdit As IFieldEdit = pLAT_DIMENField
            pLAT_DIMENFieldEdit.Name_2 = "LAT_DIMEN"
            pLAT_DIMENFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pLAT_DIMENFieldEdit.AliasName_2 = "Latitude Dimension"
            pFieldsEdit.Field_2(28) = pLAT_DIMENField

            Dim pLON_DIMENField As IField = New FieldClass()
            Dim pLON_DIMENFieldEdit As IFieldEdit = pLON_DIMENField
            pLON_DIMENFieldEdit.Name_2 = "LON_DIMEN"
            pLON_DIMENFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pLON_DIMENFieldEdit.AliasName_2 = "Longitude Dimension"
            pFieldsEdit.Field_2(29) = pLON_DIMENField


            Dim pX1Field As IField = New FieldClass()
            Dim pX1FieldEdit As IFieldEdit = pX1Field
            pX1FieldEdit.Name_2 = "X1"
            pX1FieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
            pX1FieldEdit.AliasName_2 = "West"
            pFieldsEdit.Field_2(30) = pX1Field

            Dim pX2Field As IField = New FieldClass()
            Dim pX2FieldEdit As IFieldEdit = pX2Field
            pX2FieldEdit.Name_2 = "X2"
            pX2FieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
            pX2FieldEdit.AliasName_2 = "East"
            pFieldsEdit.Field_2(31) = pX2Field

            Dim pY1Field As IField = New FieldClass()
            Dim pY1FieldEdit As IFieldEdit = pY1Field
            pY1FieldEdit.Name_2 = "Y1"
            pY1FieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
            pY1FieldEdit.AliasName_2 = "North"
            pFieldsEdit.Field_2(32) = pY1Field

            Dim pY2Field As IField = New FieldClass()
            Dim pY2FieldEdit As IFieldEdit = pY2Field
            pY2FieldEdit.Name_2 = "Y2"
            pY2FieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble
            pY2FieldEdit.AliasName_2 = "South"
            pFieldsEdit.Field_2(33) = pY2Field

            Dim pRUN_DATEField As IField = New FieldClass()
            Dim pRUN_DATEFieldEdit As IFieldEdit = pRUN_DATEField
            pRUN_DATEFieldEdit.Name_2 = "RUN_DATE"
            pRUN_DATEFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
            pRUN_DATEFieldEdit.AliasName_2 = "Run Date"
            pFieldsEdit.Field_2(34) = pRUN_DATEField

            Dim pFeatureClass As IFeatureClass = pFeatureWorkspace.CreateFeatureClass(sFeatureClass, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, strShapeFieldName, Nothing)

            Dim pFeature As IFeature = pFeatureClass.CreateFeature()





        Catch ex As Exception
            MsgBox(ex.ToString())
            Exit Sub
        End Try

    End Sub
End Class