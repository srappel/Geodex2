Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem

Public Class GDXTools
    ' Dim GeodexLayer As String = My.Settings.GeodexLayer
    Dim FeatureClassName As String = My.Settings.FeatureClassName
    'Public Const GeodexLayer As String = FeatureClassName

    Private Shared _gdxtools As GDXTools

    Private _application As IApplication

    Public Property ArcMapApplication() As IApplication
        Get
            Return _application
        End Get
        Set(ByVal value As IApplication)
            _application = value
        End Set
    End Property

    Private Sub New()

    End Sub

    Public Shared Sub createInstance(pApplication As IApplication)
        _gdxtools = New GDXTools
        _gdxtools.ArcMapApplication = pApplication
    End Sub

    Public Shared Function getInstance() As GDXTools
        Return _gdxtools
    End Function

    Public Function GetGeodexLayer() As ILayer
        Return getLayerByName(FeatureClassName)
    End Function

    Private Function getLayerByName(sLayerName As String) As ILayer

        Try
            Dim pMxDoc As IMxDocument = _application.Document
            Dim pMap As IMap = pMxDoc.FocusMap
            For i As Integer = 0 To pMap.LayerCount - 1
                Dim pLayer As ILayer = pMap.Layer(i)

                If pLayer.Name = sLayerName Then
                    Return pLayer
                End If
            Next
            Return Nothing
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try
    End Function

    Public Function FmtDegMinSec(ByVal DD As Single)
        Dim deg As Double
        Dim minDouble As Double
        Dim min As Double
        Dim sec As Double
        Dim dms As String
        Dim absDD As Double = Math.Abs(DD)

        'how to account for negative values indicating south and east?
        'Here, I will just use absolute values to determine, then we can handle the E and West in the fillform function
        'This is working well now.  Still need to account for E and W

        deg = Math.Truncate(absDD)
        minDouble = (absDD - deg) * 60
        min = Math.Truncate(minDouble)
        sec = Math.Round((minDouble - min) * 60, 3)

        'This If statement will evaluate numbers that should be clearly rounded to the nearest minute (seconds > 59.95 or < 0.05) where this error is clearly
        'due to the repeating decimal.  The round error here will be less than the degrees of freedom.  0.05 seconds translates to less than 10 feet on most latitutes 
        If Math.Abs(sec) < 0.05 And Math.Abs(sec) > 0 Then
            'very close to the current minute
            sec = 0
        ElseIf Math.Abs(sec) > 59.95 Then
            'Very close to the next higher minute
            sec = 0
            min += 1
            'This will make sure that the minutes don't exceed 59
            If min > 59 Then
                deg += 1
                min = 0
            End If
        End If
        dms = Math.Abs(deg) & "° " & Math.Abs(min) & "' " & Math.Round(sec, 2) & "''"
        Return dms

    End Function

End Class


