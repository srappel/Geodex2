Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase

<ComClass(cmdViewer.ClassId, cmdViewer.InterfaceId, cmdViewer.EventsId), _
 ProgId("gdx_editor_tools.cmdViewer")> _
Public NotInheritable Class cmdViewer
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "1b62fe69-d4db-4af4-8897-985374142626"
    Public Const InterfaceId As String = "d6faf16e-6f6d-4987-a7ca-622084ec8fd4"
    Public Const EventsId As String = "a205ed23-5e03-4489-aea7-b25454911181"
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
        MyBase.m_caption = "Viewer"   'localizable text 
        MyBase.m_message = "Manipulate the Geodex Layer in Arc Map by selecting a series or searching for records."   'localizable text 
        MyBase.m_toolTip = "Viwer" 'localizable text 
        MyBase.m_name = "Geodex_Tools_Viewer"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmdViewer.OnClick implementation
        Try
            Dim pViewer As New frmViewer

            pViewer.ArcMapApplication = _application

            'create the arc map wrapper
            Dim pArcMapApplication As New ArcMapWrapper

            pArcMapApplication.ArcMapApplication = _application

            'Show the form
            pViewer.Show(pArcMapApplication)


            Dim pMxDoc As IMxDocument = _application.Document

            'get the map from the document
            Dim pMap As IMap = pMxDoc.FocusMap
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OkOnly)
        End Try

    End Sub
End Class



