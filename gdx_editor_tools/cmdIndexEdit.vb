Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Carto

<ComClass(cmdIndexEdit.ClassId, cmdIndexEdit.InterfaceId, cmdIndexEdit.EventsId), _
 ProgId("gdx_editor_tools.cmdIndexEdit")> _
Public NotInheritable Class cmdIndexEdit
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "76ed37e6-58f7-44d7-b04c-f3b46f4712d4"
    Public Const InterfaceId As String = "d0d18769-fb5a-4b2f-87b0-72f742aa6bc8"
    Public Const EventsId As String = "7f6adb73-f4cd-43e1-9435-01667c9e383c"
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
        MyBase.m_caption = "Index Editor"   'localizable text 
        MyBase.m_message = "Allows for the editing of a spatial index of a Geodex record by inputting spatial extents in Decimal Degrees or Degrees, Minutes, Seconds and setting the Prime Meridian"   'localizable text 
        MyBase.m_toolTip = "Index Editor" 'localizable text 
        MyBase.m_name = "Geodex_Tools_IndexEditor"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmdIndexEdit.OnClick implementation

        Dim pIndexEdit As New frmIndexEdit

        pIndexEdit.ArcMapApplication = _application

        'create the arc map wrapper
        Dim pArcMapApplication As New ArcMapWrapper
        pArcMapApplication.ArcMapApplication = _application

        'Show the form
        pIndexEdit.Show(pArcMapApplication)

        Dim pMxDoc As IMxDocument = _application.Document

        'get the map from the document
        Dim pMap As IMap = pMxDoc.FocusMap

        'First clear the list:
        'pIndexEdit.gdxfilebox.Items.Clear()

        'loop through all layers and add them to the list.
        For i As Integer = 0 To pMap.LayerCount - 1
            Dim pLayer As ILayer = pMap.Layer(i)
            'pIndexEdit.gdxfilebox.Items.Add(pLayer.Name)
        Next


    End Sub
End Class