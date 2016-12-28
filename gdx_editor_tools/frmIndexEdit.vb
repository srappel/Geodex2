Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geometry
Imports gdx_editor_tools.frmRecordViewer
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms


Public Class frmIndexEdit

    Public _application As IApplication
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

    Public Sub gdxfilebox_SelectedIndexChanged(sender As Object, e As EventArgs)
        'When the gdxfilebox is changed, we want to load that layer as the pSelectedLayer
        'This box is populated with layers in the map (this code is in cmbIndexEdit.vb)



    End Sub


    Public Sub gidbox_TextChanged(sender As Object, e As EventArgs)
        'This text box will contain the object ID of the feature to be edited
        'maybe it wont need to be a .TextChanged method but
        'I just wanted to document this box.



    End Sub


    Public Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        'When this button is selected, the input method will be DMS
        'RadioButton2.Checked = False
        'RadioButton1.Checked = True
        If RadioButton1.Checked = True Then

   

            TabControl.SelectTab(0)
            northddbox.Visible = False
            southddbox.Visible = False
            eastddbox.Visible = False
            westddbox.Visible = False

            northcarbox.Visible = True
            northdegbox.Visible = True
            northminbox.Visible = True
            northsecbox.Visible = True

            southcarbox.Visible = True
            southdegbox.Visible = True
            southminbox.Visible = True
            southsecbox.Visible = True

            eastcarbox.Visible = True
            eastdegbox.Visible = True
            eastminbox.Visible = True
            eastsecbox.Visible = True

            westcarbox.Visible = True
            westdegbox.Visible = True
            westminbox.Visible = True
            westsecbox.Visible = True


        ElseIf RadioButton1.Checked = False Then
       
            TabControl.SelectTab(1)
            northddbox.Visible = True
            southddbox.Visible = True
            eastddbox.Visible = True
            westddbox.Visible = True

            northcarbox.Visible = False
            northdegbox.Visible = False
            northminbox.Visible = False
            northsecbox.Visible = False

            southcarbox.Visible = False
            southdegbox.Visible = False
            southminbox.Visible = False
            southsecbox.Visible = False

            eastcarbox.Visible = False
            eastdegbox.Visible = False
            eastminbox.Visible = False
            eastsecbox.Visible = False

            westcarbox.Visible = False
            westdegbox.Visible = False
            westminbox.Visible = False
            westsecbox.Visible = False

        End If


    End Sub


    Public Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        'When this button is selected, the input method will be DD
        'RadioButton1.Checked = False
        'RadioButton2.Checked = True
        If RadioButton2.Checked = False Then
            northddbox.Visible = False
            southddbox.Visible = False
            eastddbox.Visible = False
            westddbox.Visible = False
        ElseIf RadioButton2.Checked = True Then
            northddbox.Visible = True
            southddbox.Visible = True
            eastddbox.Visible = True
            westddbox.Visible = True
        End If
    End Sub


    Public Sub primerbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles primerbox.SelectedIndexChanged
        'This box will select the prime meridian, it will apply the prime meridian corrections
        'to the entered values and change them to Greenwich for storage in the geodatabase.
    End Sub


    Public Sub sibox_Click(sender As Object, e As EventArgs) Handles sibox.Click

        Dim northcar As String
        Dim northdeg As Integer
        Dim northmin As Integer
        Dim northsec As Integer

        Dim southcar As String
        Dim southdeg As Integer
        Dim southmin As Integer
        Dim southsec As Integer

        Dim eastcar As String
        Dim eastdeg As Integer
        Dim eastmin As Integer
        Dim eastsec As Integer

        Dim westcar As String
        Dim westdeg As Integer
        Dim westmin As Integer
        Dim westsec As Integer

        Dim northext As Decimal
        Dim southext As Decimal
        Dim eastext As Decimal
        Dim westext As Decimal

        'adjust for different PM
        Dim eastextp As Decimal
        Dim westextp As Decimal

        'adjust for values > 180 or < -180 
        Dim eastextpm As Decimal
        Dim westextpm As Decimal

        Dim northstring As String
        Dim southstring As String
        Dim eaststring As String
        Dim weststring As String
        'Dim gdxfile As String


        'Choose between DMS or DecDeg
        If RadioButton1.Checked = True Then
            'DMS
            Try
                northcar = northcarbox.Text
                northdeg = northdegbox.Text
                northmin = northminbox.Text
                northsec = northsecbox.Text

                southcar = southcarbox.Text
                southdeg = southdegbox.Text
                southmin = southminbox.Text
                southsec = southsecbox.Text

                eastcar = eastcarbox.Text
                eastdeg = eastdegbox.Text
                eastmin = eastminbox.Text
                eastsec = eastsecbox.Text

                westcar = westcarbox.Text
                westdeg = westdegbox.Text
                westmin = westminbox.Text
                westsec = westsecbox.Text
            Catch ex As Exception
                MsgBox("Please make sure you have values entered for each extent." &
                       "  Minutes and Seconds should not be null rather filled with 0s")
                Exit Sub
            End Try
            northext = northdeg + (northmin / 60) + (northsec / 3600)
            southext = southdeg + (southmin / 60) + (southsec / 3600)
            eastext = eastdeg + (eastmin / 60) + (eastsec / 3600)
            westext = westdeg + (westmin / 60) + (westsec / 3600)
        ElseIf RadioButton2.Checked = True Then
            'DEC DEG
            northext = northddbox.Text
            southext = southddbox.Text
            eastext = eastddbox.Text
            westext = westddbox.Text
        End If

        'To handle gdx file chooser and gid box

        Try

            'to handle prime meridians (convert to Greenwich)
            If primerbox.Text.Trim() = "Greenwich" Then
                eastextpm = eastext
                westextpm = westext


            ElseIf primerbox.Text.Trim() = "Athens" Then
                MsgBox("This PM has not been configured. Using Greenwich PM.", MsgBoxStyle.OkOnly)
                eastextpm = eastext
                westextpm = westext


            ElseIf primerbox.Text.Trim() = "Brussels" Then
                '4.367975 (-)
                eastextp = eastext - 4.367975
                westextp = westext - 4.367975

                If eastextp < -180 Then
                    eastextpm = 180 - (180 + eastext)
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp < -180 Then
                    westextpm = 180 - (180 + westext)
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If

            ElseIf primerbox.Text.Trim() = "Copenhagen" Then
                '12.575625 (-)
                eastextp = eastext - 12.575625
                westextp = westext - 12.575625

                If eastextp < -180 Then
                    eastextpm = 180 - (180 + eastext)
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp < -180 Then
                    westextpm = 180 - (180 + westext)
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If

            ElseIf primerbox.Text.Trim() = "Ferro" Then
                '17.662778 (+)
                eastextp = eastext + 17.662778
                westextp = westext + 17.662778

                If eastextp > 180 Then
                    eastextpm = (180 - (eastext - 180)) * -1
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp > 180 Then
                    westextpm = (180 - (westext - 180)) * -1
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If

            ElseIf primerbox.Text.Trim() = "Lisbon" Then
                '9.131906 (+)
                eastextp = eastext + 9.131906
                westextp = westext + 9.131906

                If eastextp > 180 Then
                    eastextpm = (180 - (eastext - 180)) * -1
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp > 180 Then
                    westextpm = (180 - (westext - 180)) * -1
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If

            ElseIf primerbox.Text.Trim() = "Madrid" Then
                '3.687939 (+)
                eastextp = eastext + 3.687939
                westextp = westext + 3.687939

                If eastextp > 180 Then
                    eastextpm = (180 - (eastext - 180)) * -1
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp > 180 Then
                    westextpm = (180 - (westext - 180)) * -1
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If


            ElseIf primerbox.Text.Trim() = "Munich" Then
                'NEEDS TO BE CONFIGURED
                MsgBox("This PM has not been configured. Using Greenwich PM.", MsgBoxStyle.OkOnly)

                eastextpm = eastext
                westextpm = westext

            ElseIf primerbox.Text.Trim() = "Oslo" Then
                '10.722917 (-)
                eastextp = eastext - 10.722917
                westextp = westext - 10.722917

                If eastextp < -180 Then
                    eastextpm = 180 - (180 + eastext)
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp < -180 Then
                    westextpm = 180 - (180 + westext)
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If

            ElseIf primerbox.Text.Trim() = "Paris" Then
                '2.337229 (-)
                eastextp = eastext - 2.337229
                westextp = westext - 2.337229

                If eastextp < -180 Then
                    eastextpm = 180 - (180 + eastext)
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp < -180 Then
                    westextpm = 180 - (180 + westext)
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If

            ElseIf primerbox.Text.Trim() = "Pulkovo" Then
                'NEEDS TO BE CONFIGURED
                MsgBox("This PM has not been configured. Using Greenwich PM.", MsgBoxStyle.OkOnly)
                eastextpm = eastext
                westextpm = westext
            ElseIf primerbox.Text.Trim() = "Rio de Janeriro" Then
                '-73.171944 (+)
                eastextp = eastext + 73.171944
                westextp = westext + 73.171944

                If eastextp > 180 Then
                    eastextpm = (180 - (eastext - 180)) * -1
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp > 180 Then
                    westextpm = (180 - (westext - 180)) * -1
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If

            ElseIf primerbox.Text.Trim() = "Rome" Then
                '12.452333 (-)
                eastextp = eastext - 12.452333
                westextp = westext - 12.452333

                If eastextp < -180 Then
                    eastextpm = 180 - (180 + eastext)
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp < -180 Then
                    westextpm = 180 - (180 + westext)
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If


            ElseIf primerbox.Text.Trim() = "Quito" Then
                'NEEDS TO BE CONFIGURED
                MsgBox("This PM has not been configured. Using Greenwich PM.", MsgBoxStyle.OkOnly)
                eastextpm = eastext
                westextpm = westext


            ElseIf primerbox.Text.Trim() = "Stockholm" Then
                '18.058278 (-)
                eastextp = eastext - 18.058278
                westextp = westext - 18.058278

                If eastextp < -180 Then
                    eastextpm = 180 - (180 + eastext)
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp < -180 Then
                    westextpm = 180 - (180 + westext)
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If

            ElseIf primerbox.Text.Trim() = "Tirane" Then
                '19.779167 (-)
                eastextp = eastext - 19.779167
                westextp = westext - 19.779167

                If eastextp < -180 Then
                    eastextpm = 180 - (180 + eastext)
                ElseIf eastextp = -180 Then
                    eastextpm = 180.0
                Else
                    eastextpm = eastextp
                End If

                If westextp < -180 Then
                    westextpm = 180 - (180 + westext)
                ElseIf westextp = -180 Then
                    westextpm = 180
                Else
                    westextpm = westextp
                End If

            Else
                MsgBox("Prime meridian """ & primerbox.Text.ToString & """ is not a valid Prime Meridian.  Please select from the options or edit IndexEdit.vb to convert PM to Greenwich.", MsgBoxStyle.OkOnly)
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox("error adjusting PM", MsgBoxStyle.OkOnly)
            Exit Sub
        End Try



        If RadioButton1.Checked = True Then

            Try
                If northcarbox.Text.ToUpper.Trim() = "S" Then
                    northstring = "-" & northext.ToString.Trim()
                ElseIf northcarbox.Text.ToUpper.Trim() = "N" Then
                    northstring = northext.ToString.Trim()
                Else
                    MsgBox("Please enter either ""S"" or ""N"" for the northern extent coordinate!", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                If southcarbox.Text.ToUpper.Trim() = "S" Then
                    southstring = "-" & southext.ToString.Trim()
                ElseIf southcarbox.Text.ToUpper.Trim() = "N" Then
                    southstring = southext.ToString.Trim()
                Else
                    MsgBox("Please enter either ""S"" or ""N"" for the southern extent coordinate!", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                If eastcarbox.Text.ToUpper.Trim() = "W" Then
                    eaststring = "-" & eastextpm.ToString.Trim()
                ElseIf eastcarbox.Text.ToUpper.Trim() = "E" Then
                    eaststring = eastextpm.ToString.Trim()
                Else
                    MsgBox("Please enter either ""E"" or ""W"" for the eastern extent coordinate!", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If

                If westcarbox.Text.ToUpper.Trim() = "W" Then
                    weststring = "-" & westextpm.ToString.Trim()
                ElseIf westcarbox.Text.ToUpper.Trim() = "E" Then
                    weststring = westextpm.ToString.Trim()
                Else
                    MsgBox("Please enter either ""E"" or ""W"" for the western extent coordinate!", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
            Catch ex As Exception
                MsgBox("error with cardinal directions")
                Exit Sub
            End Try
        ElseIf RadioButton2.Checked = True Then
            northstring = northddbox.Text.Trim()
            southstring = southddbox.Text.Trim()
            eaststring = eastextpm.ToString.Trim()
            weststring = westextpm.ToString.Trim()

        Else
            MsgBox("Select either decimal degree input or DMS input", MsgBoxStyle.OkOnly)
            Exit Sub
        End If






        'northstring, southstring, eaststring, and westring need to be entered as the
        'spatial extents for the record(s) in DD
        'They are the final outputs of all the above code!
        'Need to watch Hussein Nasser's videos on editing features first!
        Dim west As Double = weststring
        Dim east As Double = eaststring
        Dim south As Double = southstring
        Dim north As Double = northstring

        Dim pPolygonPoints As IPointCollection = New Polygon

        Dim pPolygon As IPolygon = pPolygonPoints


        Dim pPoint(0 To 3) As Point
        For i = 0 To 3
            pPoint(i) = New Point
        Next i
        pPoint(0).X = west
        pPoint(0).Y = north
        pPoint(1).X = east
        pPoint(1).Y = north
        pPoint(2).X = east
        pPoint(2).Y = south
        pPoint(3).X = west
        pPoint(3).Y = south

        pPolygonPoints.AddPoint(pPoint(0))
        pPolygonPoints.AddPoint(pPoint(1))
        pPolygonPoints.AddPoint(pPoint(2))
        pPolygonPoints.AddPoint(pPoint(3))
        pPolygonPoints.AddPoint(pPoint(0))
        'Just added the second point 0 to make a full circle of polygons

        Dim pGeogRef As ISpatialReference = GetcoordinateSystem()
        pPolygon.SpatialReference = pGeogRef

        Dim pOID As Integer = My.Settings.CurrentOID
        If pOID = 0 Then
            MsgBox("No feature is selected, chose a feature by OID and load it before clicking ""Create/Update Spatial Reference""", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pDataset As IDataset = pFeatureClass
        Dim pWorkspace As IWorkspace = pDataset.Workspace
        Dim pWorkspaceEdit As IWorkspaceEdit = pWorkspace

        'MsgBox("The new spatial index will be:" & vbCrLf & "North Extent: " & north.ToString() & vbCrLf &
        '"South Extent: " & south.ToString() & vbCrLf & "West Extent: " & west.ToString() & vbCrLf & "East Extent :" &
        'east.ToString())

        Try
            pWorkspaceEdit.StartEditing(True)
            pWorkspaceEdit.StartEditOperation()

            Dim pFeature As IFeature = pFeatureClass.GetFeature(pOID)

            pFeature.Value(pFeature.Fields.FindField("X1")) = west
            pFeature.Value(pFeature.Fields.FindField("X2")) = east
            pFeature.Value(pFeature.Fields.FindField("Y1")) = north
            pFeature.Value(pFeature.Fields.FindField("Y2")) = south
            pFeature.Value(pFeature.Fields.FindField("RUN_DATE")) = DateAndTime.DateString() & My.User.Name.ToString()

            Dim pSpatialRef As ISpatialReference = GetcoordinateSystem()
            pPolygon.Project(pSpatialRef)
            pFeature.Shape = pPolygon

            pFeature.Store()
            pWorkspaceEdit.StopEditOperation()
            pWorkspaceEdit.StopEditing(True)
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OkOnly)

        End Try

        My.Settings.northstring = north.ToString()
        My.Settings.southstring = south.ToString()
        My.Settings.eaststring = east.ToString()
        My.Settings.weststring = west.ToString()


        Dim pMxdoc As IMxDocument = _application.Document
        pMxdoc.ActiveView.Refresh()
        Me.Close()

    End Sub

    Public Sub cmdClear1_Click(sender As Object, e As EventArgs) Handles cmdClear1.Click
        northcarbox.Text = "N"
        northdegbox.Text = "00"
        northminbox.Text = "00"
        northsecbox.Text = "00"
        southcarbox.Text = "N"
        southdegbox.Text = "00"
        southminbox.Text = "00"
        southsecbox.Text = "00"
        eastcarbox.Text = "W"
        eastdegbox.Text = "00"
        eastminbox.Text = "00"
        eastsecbox.Text = "00"
        westcarbox.Text = "W"
        westdegbox.Text = "00"
        westminbox.Text = "00"
        westsecbox.Text = "00"
        northddbox.Text = "0.0"
        southddbox.Text = "0.0"
        eastddbox.Text = "0.0"
        westddbox.Text = "0.0"
    End Sub

    Public Sub northcarbox_TextChanged(sender As Object, e As EventArgs)
        RadioButton1.Checked = True
    End Sub

    Public Sub northddbox_TextChanged(sender As Object, e As EventArgs)
        RadioButton2.Checked = True
    End Sub

    'This hasn't been debugged yet
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Public Sub frmIndexEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
        northddbox.Visible = False
        southddbox.Visible = False
        eastddbox.Visible = False
        westddbox.Visible = False
        primerbox.SelectedIndex = 0

        Dim pLayer As ILayer = GDXTools.getInstance.GetGeodexLayer
        If pLayer Is Nothing Then
            MsgBox("Layer Geodex is not loaded, use the load data button to load the layer.")
            Me.Close()
            Exit Sub
        End If


        If Not My.Settings.CurrentOID = 0 Then
            lblOID.Text = My.Settings.CurrentOID.ToString()
            Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
            Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
            Dim pFeature As IFeature = pFeatureClass.GetFeature(My.Settings.CurrentOID)

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("RECORD"))) = False Then
                lblRecord.Text = pFeature.Value(pFeature.Fields.FindField("RECORD"))
            Else
                lblRecord.Text = "..."
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("LOCATION"))) = False Then
                lblLocation.Text = pFeature.Value(pFeature.Fields.FindField("LOCATION"))
            Else
                lblLocation.Text = "..."
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("GDX_NUM"))) = False Then
                lblSeries.Text = pFeature.Value(pFeature.Fields.FindField("GDX_NUM"))
            Else
                lblSeries.Text = "..."
            End If

            If IsDBNull(pFeature.Value(pFeature.Fields.FindField("DATE"))) = False Then
                lblDate.Text = pFeature.Value(pFeature.Fields.FindField("DATE")).ToString()
            Else
                lblDate.Text = "..."
            End If
        Else
            lblOID.Text = "..."
        End If

    End Sub

    'This grabs the projection from the install folder on the C drive
    'The commented text below is something I'm working on to try to grab the projection from a feature sot hat the projection
    'file doesn't need to be stored on the machine.

    Public Function GetcoordinateSystem() As IGeographicCoordinateSystem
        Dim pSpatialReferenceFactory As ISpatialReferenceFactory = New SpatialReferenceEnvironmentClass()
        Dim pGeographicCoordinateSystem As IGeographicCoordinateSystem = pSpatialReferenceFactory.CreateESRISpatialReferenceFromPRJFile("C:\GEODEX\wgs84.prj")

        'Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        'Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        'Dim pFeature As IFeature = pFeatureClass.GetFeature(My.Settings.CurrentOID)
        'Dim pGeoDataset As IGeoDataset = pFeatureLayer
        'Dim sr As IGeographicCoordinateSystem = pGeoDataset.SpatialReference

        Return pGeographicCoordinateSystem
    End Function

    Private Sub btLoad_Click(sender As Object, e As EventArgs) Handles btLoad.Click
        If txtOID.Text Is Nothing Then
            Exit Sub
        ElseIf txtOID.Text = "" Then
            Exit Sub
        End If

        Dim pOID As Integer = txtOID.Text.Trim()
        My.Settings.CurrentOID = pOID
        Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer
        Dim pFeatureClass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pFeature As IFeature = pFeatureClass.GetFeature(My.Settings.CurrentOID)

        If IsDBNull(pFeature.Value(pFeature.Fields.FindField("RECORD"))) = False Then
            lblRecord.Text = pFeature.Value(pFeature.Fields.FindField("RECORD"))
        Else
            lblRecord.Text = "..."
        End If

        If IsDBNull(pFeature.Value(pFeature.Fields.FindField("LOCATION"))) = False Then
            lblLocation.Text = pFeature.Value(pFeature.Fields.FindField("LOCATION"))
        Else
            lblLocation.Text = "..."
        End If

        If IsDBNull(pFeature.Value(pFeature.Fields.FindField("GDX_NUM"))) = False Then
            lblSeries.Text = pFeature.Value(pFeature.Fields.FindField("GDX_NUM"))
        Else
            lblSeries.Text = "..."
        End If

        If IsDBNull(pFeature.Value(pFeature.Fields.FindField("DATE"))) = False Then
            lblDate.Text = pFeature.Value(pFeature.Fields.FindField("DATE")).ToString()
        Else
            lblDate.Text = "..."
        End If
        lblOID.Text = My.Settings.CurrentOID.ToString()
    End Sub

    Private Sub btLoadExt_Click(sender As Object, e As EventArgs) Handles btLoadExt.Click
        Dim pFeatureLayer As IFeatureLayer = GDXTools.getInstance.GetGeodexLayer()
        Dim pMXDocument As IMxDocument = _application.Document
        Dim pMap As IMap = pMXDocument.FocusMap
        Dim count As Integer = pMap.SelectionCount
        Dim pOID As Integer = 0
        Dim pFeatureSelection As IFeatureSelection = pFeatureLayer
        Dim pSelectionSet As ISelectionSet = pFeatureSelection.SelectionSet
        Dim eIDS As IEnumIDs = pSelectionSet.IDs
        Dim pFeatureclass As IFeatureClass = pFeatureLayer.FeatureClass
        Dim pFeature As IFeature = Nothing

        If count <> 1 Then
            MsgBox("this function will only work if one feature is selected on the map")
            Exit Sub
        Else
            pOID = eIDS.Next
            pFeature = pFeatureclass.GetFeature(pOID)
            Try
                RadioButton2.Checked = True
                northddbox.Text = pFeature.Value(pFeature.Fields.FindField("Y1"))
                southddbox.Text = pFeature.Value(pFeature.Fields.FindField("Y2"))
                westddbox.Text = pFeature.Value(pFeature.Fields.FindField("X1"))
                eastddbox.Text = pFeature.Value(pFeature.Fields.FindField("X2"))

            Catch ex As Exception
                MsgBox("There was an error when filling the form:" & vbCrLf & ex.ToString())
                Exit Sub
            End Try

        End If
    End Sub

    Private Sub TabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl.SelectedIndexChanged

        If TabControl.SelectedTab Is TabPage1 Then
            RadioButton1.Checked = True
            RadioButton2.Checked = False
        Else
            RadioButton1.Checked = False
            RadioButton2.Checked = True
        End If

        'If TabControl.SelectedTab.Text = "DMS" Then
        'RadioButton1.Checked = True
        'RadioButton2.Checked = False
        'Else
        'RadioButton1.Checked = False
        'RadioButton2.Checked = True
        'End If
    End Sub
End Class