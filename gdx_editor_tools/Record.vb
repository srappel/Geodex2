Public Class Record

    Private _GDX_FILE As Short
    Public Property GDX_FILE() As Short
        Get
            Return _GDX_FILE
        End Get
        Set(ByVal value As Short)
            _GDX_FILE = value
        End Set

    End Property

    Private _GDX_NUM As String
    Public Property GDX_NUM() As String
        Get
            Return _GDX_NUM
        End Get
        Set(ByVal value As String)
            _GDX_NUM = value
        End Set
    End Property

    Private _GDX_SUB As Short
    Public Property GDX_SUB() As Short
        Get
            Return _GDX_SUB
        End Get
        Set(ByVal value As Short)
            _GDX_SUB = value
        End Set
    End Property


    Private _record As String
    Public Property RECORD() As String
        Get
            Return _record
        End Get
        Set(ByVal value As String)
            _record = value
        End Set
    End Property

    Private _location As String
    Public Property LOCATION() As String
        Get
            Return _location
        End Get
        Set(ByVal value As String)
            _location = value
        End Set
    End Property


    Private _date As Short
    Public Property rDATE() As Short
        Get
            Return _date
        End Get
        Set(ByVal value As Short)
            _date = value
        End Set
    End Property

    Private _series_tit As String
    Public Property SERIES_TIT() As String
        Get
            Return _series_tit
        End Get
        Set(ByVal value As String)
            _series_tit = value
        End Set
    End Property

    Private _publisher As String
    Public Property PUBLISHER() As String
        Get
            Return _publisher
        End Get
        Set(ByVal value As String)
            _publisher = value
        End Set
    End Property

    Private _map_type As Short
    Public Property MAP_TYPE() As Short
        Get
            Return _map_type
        End Get
        Set(ByVal value As Short)
            _map_type = value
        End Set
    End Property

    Private _production As Short
    Public Property PRODUCTION() As Short
        Get
            Return _production
        End Get
        Set(ByVal value As Short)
            _production = value
        End Set
    End Property

    Private _format As Short
    Public Property MAP_FOR() As Short
        Get
            Return _format
        End Get
        Set(ByVal value As Short)
            _format = value
        End Set
    End Property

    Private _projection As Short
    Public Property PROJECT() As Short
        Get
            Return _projection
        End Get
        Set(ByVal value As Short)
            _projection = value
        End Set
    End Property

    Private _prime_mer As Short
    Public Property PRIME_MER() As Short
        Get
            Return _prime_mer
        End Get
        Set(ByVal value As Short)
            _prime_mer = value
        End Set
    End Property

    Private _scale As Long
    Public Property SCALE() As Long
        Get
            Return _scale
        End Get
        Set(ByVal value As Long)
            _scale = value
        End Set
    End Property

    Private _catloc As String
    Public Property CATLOC() As String
        Get
            Return _catloc
        End Get
        Set(ByVal value As String)
            _catloc = value
        End Set
    End Property

    Private _hold As Short
    Public Property HOLD() As Short
        Get
            Return _hold
        End Get
        Set(ByVal value As Short)
            _hold = value
        End Set
    End Property

    Private _year1 As Short
    Public Property YEAR1() As Short
        Get
            Return _year1
        End Get
        Set(ByVal value As Short)
            _year1 = value
        End Set
    End Property

    Private _year1_type As Short
    Public Property YEAR1_TYPE() As Short
        Get
            Return _year1_type
        End Get
        Set(ByVal value As Short)
            _year1_type = value
        End Set
    End Property

    Private _year2 As Short
    Public Property YEAR2() As Short
        Get
            Return _year2
        End Get
        Set(ByVal value As Short)
            _year2 = value
        End Set
    End Property

    Private _year2_type As Short
    Public Property YEAR2_TYPE() As Short
        Get
            Return _YEAR2_TYPE
        End Get
        Set(ByVal value As Short)
            _year2_type = value
        End Set
    End Property

    Private _year3 As Short
    Public Property YEAR3() As Short
        Get
            Return _year3
        End Get
        Set(ByVal value As Short)
            _year3 = value
        End Set
    End Property

    Private _year3_type As Short
    Public Property YEAR3_TYPE() As Short
        Get
            Return _YEAR3_TYPE
        End Get
        Set(ByVal value As Short)
            _year3_type = value
        End Set
    End Property

    Private _year4 As Short
    Public Property YEAR4() As Short
        Get
            Return _YEAR4
        End Get
        Set(ByVal value As Short)
            _year4 = value
        End Set
    End Property

    Private _year4_type As String
    Public Property YEAR4_TYPE() As String
        Get
            Return _year4_type
        End Get
        Set(ByVal value As String)
            _year4_type = value
        End Set
    End Property

    Private _edition As Short
    Public Property EDITION_NO() As Short
        Get
            Return _edition
        End Get
        Set(ByVal value As Short)
            _edition = value
        End Set
    End Property

    Private _iso_type As Short
    Public Property ISO_TYPE() As Short
        Get
            Return _iso_type
        End Get
        Set(ByVal value As Short)
            _iso_type = value
        End Set
    End Property

    Private _iso_val As Short
    Public Property ISO_VAL() As Short
        Get
            Return _iso_val
        End Get
        Set(ByVal value As Short)
            _iso_val = value
        End Set
    End Property

    Private _lat_dimen As String
    Public Property LAT_DIMEN()As String
        Get
            Return _lat_dimen
        End Get
        Set(ByVal value As String)
            _lat_dimen = value
        End Set
    End Property

    Private _lon_dimen As String
    Public Property LON_DIMEN() As String
        Get
            Return _lon_dimen
        End Get
        Set(ByVal value As String)
            _lon_dimen = value
        End Set
    End Property

    Private _x1 As Double
    Public Property X1() As Double
        Get
            Return _x1
        End Get
        Set(ByVal value As Double)
            _x1 = value
        End Set
    End Property

    Private _x2 As Double
    Public Property X2() As Double
        Get
            Return _x2
        End Get
        Set(ByVal value As Double)
            _x2 = value
        End Set
    End Property

    Private _y1 As Double
    Public Property Y1() As Double
        Get
            Return _y1
        End Get
        Set(ByVal value As Double)
            _y1 = value
        End Set
    End Property

    Private _y2 As Double
    Public Property Y2() As Double
        Get
            Return _y2
        End Get
        Set(ByVal value As Double)
            _y2 = value
        End Set
    End Property

    Private _run_date As String
    Public Property RUN_DATE() As String
        Get
            Return _run_date
        End Get
        Set(ByVal value As String)
            _run_date = value
        End Set
    End Property

End Class
