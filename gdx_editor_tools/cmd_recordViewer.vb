Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms

<ComClass(cmd_recordViewer.ClassId, cmd_recordViewer.InterfaceId, cmd_recordViewer.EventsId), _
 ProgId("gdx_editor_tools.cmd_recordViewer")> _
Public NotInheritable Class cmd_recordViewer
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "035be670-36ef-4260-bb7d-31a53ea3e0ea"
    Public Const InterfaceId As String = "3928ad1a-7506-4594-be09-1111cc2ccfb7"
    Public Const EventsId As String = "14e10558-271d-4e3d-b597-3f84003205bb"
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
        MyBase.m_caption = "Record Editor"   'localizable text 
        MyBase.m_message = "Create a new record in the Geodex feature class"   'localizable text 
        MyBase.m_toolTip = "Record Editor" 'localizable text 
        MyBase.m_name = "Geodex_Tools_RecordEditor"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")



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
End Class

