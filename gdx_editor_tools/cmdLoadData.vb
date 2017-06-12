Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.esriSystem

<ComClass(cmdLoadData.ClassId, cmdLoadData.InterfaceId, cmdLoadData.EventsId), _
 ProgId("gdx_editor_tools.cmdLoadData")> _
Public NotInheritable Class cmdLoadData
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3597a720-21d4-47f5-ab3e-673e03f34e69"
    Public Const InterfaceId As String = "b8e1a9aa-a7c0-41b4-865a-54e6d234b291"
    Public Const EventsId As String = "e1d7ad56-cd39-4807-b116-ce30579d6b4e"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region


    Private _application As IApplication

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "Geodex_Tools"  'localizable text 
        MyBase.m_caption = "Load Data"   'localizable text 
        MyBase.m_message = "Load Data into ArcMap from Geodex Database"   'localizable text 
        MyBase.m_toolTip = "Load Data" 'localizable text 
        MyBase.m_name = "Geodex_Tools_LoadData"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            'TODO: change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try
    End Sub

    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            _application = CType(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If
        GDXTools.createInstance(_application)
        ' TODO:  Add other initialization code
    End Sub

    Public Overrides Sub OnClick()
        'TODO: Add cmdLoadData.OnClick implementation
        'if my.settings.dblocal is false then we're connecting to a SDE
        'if my.settings.dblocal is true we are connecting to a FGDB
        'MsgBox("The button is working", MsgBoxStyle.OkOnly)

        If My.Settings.DBLocal = False Then
            Try
                'MsgBox("This may take some time.  Please be patient.", MsgBoxStyle.OkOnly)
                Try
                    Dim pPropertySet As IPropertySet = New PropertySetClass()
                    pPropertySet.SetProperty("SERVER", My.Settings.server)
                    pPropertySet.SetProperty("AUTHENTICATION_MODE", My.Settings.auth)
                    pPropertySet.SetProperty("INSTANCE", My.Settings.instance)
                    pPropertySet.SetProperty("DATABASE", My.Settings.database)
                    pPropertySet.SetProperty("VERSION", My.Settings.version)

                    Dim pWorkspaceFactory As IWorkspaceFactory = New SdeWorkspaceFactory
                    Dim pWorkspace As IWorkspace = pWorkspaceFactory.Open(pPropertySet, _application.hWnd)
                    ' OpenFromFile("Database Connections\Connection to WEBGIS.sde", _application.hWnd)
                    Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
                    Dim FeatureClassName As String = My.Settings.FeatureClassName
                    Dim pGeodexFeatureClass As IFeatureClass = pFeatureWorkspace.OpenFeatureClass(FeatureClassName)
                    Dim pMXDoc As IMxDocument = _application.Document
                    Dim pMap As IMap = pMXDoc.FocusMap

                    For i As Integer = 0 To pMap.LayerCount - 1
                        Dim pLayer As ILayer = pMap.Layer(i)

                        If pLayer.Name = FeatureClassName Then
                            MsgBox("The layer """ & FeatureClassName & """ is already present in the document", MsgBoxStyle.OkOnly)
                            Exit Sub
                        End If
                    Next
                    Dim pFeatureLayer As IFeatureLayer = New FeatureLayer
                    pFeatureLayer.Name = FeatureClassName
                    pFeatureLayer.FeatureClass = pGeodexFeatureClass
                    pMXDoc.FocusMap.AddLayer(pFeatureLayer)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                    Exit Sub
                End Try
            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.OkOnly)
                Exit Sub
            End Try
        ElseIf My.Settings.DBLocal = True Then
            Try
                Dim pWorkspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
                If My.Settings.path = Nothing Then
                    Exit Sub
                End If
                Dim pWorkspace As IWorkspace = pWorkspaceFactory.OpenFromFile(My.Settings.path, _application.hWnd)
                ' OpenFromFile("Database Connections\Connection to WEBGIS.sde", _application.hWnd)
                Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspace
                Dim FeatureClassName As String = My.Settings.FeatureClassName
                Dim pGeodexFeatureClass As IFeatureClass = pFeatureWorkspace.OpenFeatureClass(FeatureClassName)
                Dim pMXDoc As IMxDocument = _application.Document
                Dim pMap As IMap = pMXDoc.FocusMap

                For i As Integer = 0 To pMap.LayerCount - 1
                    Dim pLayer As ILayer = pMap.Layer(i)

                    If pLayer.Name = FeatureClassName Then
                        MsgBox("The layer """ & FeatureClassName & """ is already present in the document", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                Next
                Dim pFeatureLayer As IFeatureLayer = New FeatureLayer
                pFeatureLayer.Name = FeatureClassName
                pFeatureLayer.FeatureClass = pGeodexFeatureClass
                pMXDoc.FocusMap.AddLayer(pFeatureLayer)
            Catch ex As Exception
                MsgBox(ex.ToString)
                Exit Sub
            End Try
        End If
        
    End Sub
End Class