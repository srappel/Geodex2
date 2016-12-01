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
End Class


