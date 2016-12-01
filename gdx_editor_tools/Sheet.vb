Imports System.Windows.Forms
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI


Public Class Sheet

    Private _record As String
    Public Property Record() As String
        Get
            Return _record
        End Get
        Set(ByVal value As String)
            _record = value
        End Set
    End Property

    Private _location As String
    Public Property Location() As String
        Get
            Return _location
        End Get
        Set(ByVal value As String)
            _location = value
        End Set
    End Property

    Private _date As Short
    Public Property SheetDate() As Short
        Get
            Return _date
        End Get
        Set(ByVal value As Short)
            _date = value
        End Set
    End Property

    Private _objectid As String
    Public Property ObjectID() As String
        Get
            Return _objectid
        End Get
        Set(ByVal value As String)
            _objectid = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Dim sSheet As String
        If My.Settings.sheet = "record" Then
            sSheet = String.Format("{0} - {1} ({2})", _record.Trim(), _location.Trim(), _date.ToString().Trim())
        ElseIf My.Settings.sheet = "location" Then
            sSheet = String.Format("{1} - {0} ({2})", _record.Trim(), _location.Trim(), _date.ToString().Trim())
        Else
            sSheet = String.Format("{0} - {1} ({2})", _record.Trim(), _location.Trim(), _date.ToString().Trim())
        End If

        Return sSheet
    End Function

End Class

Public Class Series
    Private _gdxname As String
    Public Property GdxName() As String
        Get
            Return _gdxname
        End Get
        Set(ByVal value As String)
            _gdxname = value
        End Set
    End Property

    Private _gdxcode As Short
    Public Property GdxCode() As Short
        Get
            Return _gdxcode
        End Get
        Set(ByVal value As Short)
            _gdxcode = value
        End Set
    End Property


    Public Overrides Function ToString() As String
        Dim code As String
        If Len(_gdxcode.ToString()) = 1 Then
            code = "F000" & _gdxcode.ToString().Trim()
        ElseIf Len(_gdxcode.ToString()) = 2 Then
            code = "F00" & _gdxcode.ToString().Trim()
        ElseIf Len(_gdxcode.ToString() = 3) Then
            code = "F0" & _gdxcode.ToString().Trim()
        Else
            code = "F" & _gdxcode.ToString().Trim()
        End If
        Dim stringprint As String = _gdxname & " (" & code & ")"
        Return stringprint
    End Function

End Class

Public Class GDX_SUB
    Private _code As Short
    Public Property Code() As Short
        Get
            Return _code
        End Get
        Set(ByVal value As Short)
            _code = value
        End Set
    End Property

    Private _description As String
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _description
    End Function

End Class


Public Class ArcMapWrapper
    Implements IWin32Window

    Private _application As IApplication

    Public Property ArcMapApplication() As IApplication
        Get
            Return _application
        End Get
        Set(ByVal value As IApplication)
            _application = value
        End Set
    End Property


    Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
        Get
            Return New IntPtr(_application.hWnd)
        End Get
    End Property
End Class

Public Class Tools
    Private _application As IApplication
    Public Property ArcMapApplication() As IApplication
        Get
            Return _application
        End Get
        Set(ByVal value As IApplication)
            _application = value
        End Set
    End Property

    
End Class

Public Class YearType

    Private _description As String
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Private _code As Integer
    Public Property Code() As Integer
        Get
            Return _code
        End Get
        Set(ByVal value As Integer)
            _code = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _description

    End Function
End Class

Public Class IsoType
    Private _description As String
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Private _code As Integer
    Public Property Code() As Integer
        Get
            Return _code
        End Get
        Set(ByVal value As Integer)
            _code = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _description
    End Function

End Class

Public Class MapType
    Private _description As String
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Private _code As Integer
    Public Property Code() As Integer
        Get
            Return _code
        End Get
        Set(ByVal value As Integer)
            _code = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _description
    End Function
End Class

Public Class Production
    Private _description As String
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Private _code As Integer
    Public Property Code() As Integer
        Get
            Return _code
        End Get
        Set(ByVal value As Integer)
            _code = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _description
    End Function
End Class

Public Class Projection
    Private _description As String
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Private _code As Integer
    Public Property Code() As Integer
        Get
            Return _code
        End Get
        Set(ByVal value As Integer)
            _code = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _description
    End Function
End Class

Public Class PrimeMer
    Private _description As String
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Private _code As Integer
    Public Property Code() As Integer
        Get
            Return _code
        End Get
        Set(ByVal value As Integer)
            _code = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _description
    End Function
End Class

Public Class Format
    Private _description As String
    Public Property Description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Private _code As Integer
    Public Property Code() As Integer
        Get
            Return _code
        End Get
        Set(ByVal value As Integer)
            _code = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _description
    End Function
End Class

Public Class CountYear
    Private _year As Short
    Public Property Year() As Short
        Get
            Return _year
        End Get
        Set(ByVal value As Short)
            _year = value
        End Set
    End Property

    Private _yeartype As Short
    Public Property YearType() As Short
        Get
            Return _yeartype
        End Get
        Set(ByVal value As Short)
            _yeartype = value
        End Set
    End Property

End Class
