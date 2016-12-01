<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReconcile
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReconcile))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblHold = New System.Windows.Forms.Label()
        Me.cmbSelect = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmdGetSelect = New System.Windows.Forms.Button()
        Me.cmbClear = New System.Windows.Forms.Button()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.cmbFile = New System.Windows.Forms.ComboBox()
        Me.cmbSortOrder = New System.Windows.Forms.ComboBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.cmbJumpType = New System.Windows.Forms.ComboBox()
        Me.cmdSearch = New System.Windows.Forms.Button()
        Me.txtQuery = New System.Windows.Forms.TextBox()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.txtHold = New System.Windows.Forms.TextBox()
        Me.cmbFlash = New System.Windows.Forms.Button()
        Me.cmbZoom = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.picYes = New System.Windows.Forms.PictureBox()
        Me.cmbNext = New System.Windows.Forms.Button()
        Me.cmbPrev = New System.Windows.Forms.Button()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cmdNO = New System.Windows.Forms.Button()
        Me.cmdYES = New System.Windows.Forms.Button()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.lblE = New System.Windows.Forms.Label()
        Me.lblW = New System.Windows.Forms.Label()
        Me.lblS = New System.Windows.Forms.Label()
        Me.lblN = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtPublisher = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtFormat = New System.Windows.Forms.TextBox()
        Me.lblformat = New System.Windows.Forms.Label()
        Me.txtPrimeMer = New System.Windows.Forms.TextBox()
        Me.lblprimer = New System.Windows.Forms.Label()
        Me.txtProjection = New System.Windows.Forms.TextBox()
        Me.lblprojec = New System.Windows.Forms.Label()
        Me.txtProduction = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtMapType = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtScale = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtOID = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtDates = New System.Windows.Forms.TextBox()
        Me.txtFile = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtSeriesTit = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtCataloc = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtEdition = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtDate = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtLocation = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtRecord = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.picNo = New System.Windows.Forms.PictureBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.picYes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.picNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblHold)
        Me.GroupBox1.Controls.Add(Me.cmbSelect)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.cmdDelete)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.txtHold)
        Me.GroupBox1.Controls.Add(Me.cmbFlash)
        Me.GroupBox1.Controls.Add(Me.cmbZoom)
        Me.GroupBox1.Controls.Add(Me.cmdEdit)
        Me.GroupBox1.Controls.Add(Me.picYes)
        Me.GroupBox1.Controls.Add(Me.cmbNext)
        Me.GroupBox1.Controls.Add(Me.cmbPrev)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.cmdNO)
        Me.GroupBox1.Controls.Add(Me.cmdYES)
        Me.GroupBox1.Controls.Add(Me.GroupBox6)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.picNo)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(738, 460)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'lblHold
        '
        Me.lblHold.AutoSize = True
        Me.lblHold.Location = New System.Drawing.Point(390, 114)
        Me.lblHold.Name = "lblHold"
        Me.lblHold.Size = New System.Drawing.Size(25, 13)
        Me.lblHold.TabIndex = 13
        Me.lblHold.Text = "Yes"
        '
        'cmbSelect
        '
        Me.cmbSelect.Location = New System.Drawing.Point(450, 117)
        Me.cmbSelect.Name = "cmbSelect"
        Me.cmbSelect.Size = New System.Drawing.Size(75, 23)
        Me.cmbSelect.TabIndex = 14
        Me.cmbSelect.Text = "Select"
        Me.cmbSelect.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label12)
        Me.GroupBox4.Controls.Add(Me.cmdGetSelect)
        Me.GroupBox4.Controls.Add(Me.cmbClear)
        Me.GroupBox4.Controls.Add(Me.Label20)
        Me.GroupBox4.Controls.Add(Me.cmbFile)
        Me.GroupBox4.Controls.Add(Me.cmbSortOrder)
        Me.GroupBox4.Controls.Add(Me.Label21)
        Me.GroupBox4.Controls.Add(Me.cmbJumpType)
        Me.GroupBox4.Controls.Add(Me.cmdSearch)
        Me.GroupBox4.Controls.Add(Me.txtQuery)
        Me.GroupBox4.Location = New System.Drawing.Point(6, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(358, 132)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(7, 48)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(41, 13)
        Me.Label12.TabIndex = 2
        Me.Label12.Text = "Search"
        '
        'cmdGetSelect
        '
        Me.cmdGetSelect.Location = New System.Drawing.Point(216, 97)
        Me.cmdGetSelect.Name = "cmdGetSelect"
        Me.cmdGetSelect.Size = New System.Drawing.Size(136, 23)
        Me.cmdGetSelect.TabIndex = 9
        Me.cmdGetSelect.Text = "Get Selected"
        Me.cmdGetSelect.UseVisualStyleBackColor = True
        '
        'cmbClear
        '
        Me.cmbClear.Location = New System.Drawing.Point(6, 97)
        Me.cmbClear.Name = "cmbClear"
        Me.cmbClear.Size = New System.Drawing.Size(74, 23)
        Me.cmbClear.TabIndex = 8
        Me.cmbClear.Text = "Clear Form"
        Me.cmbClear.UseVisualStyleBackColor = True
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(7, 74)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(55, 13)
        Me.Label20.TabIndex = 5
        Me.Label20.Text = "Sort Order"
        '
        'cmbFile
        '
        Me.cmbFile.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbFile.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbFile.FormattingEnabled = True
        Me.cmbFile.Location = New System.Drawing.Point(51, 18)
        Me.cmbFile.Name = "cmbFile"
        Me.cmbFile.Size = New System.Drawing.Size(301, 21)
        Me.cmbFile.Sorted = True
        Me.cmbFile.TabIndex = 1
        '
        'cmbSortOrder
        '
        Me.cmbSortOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSortOrder.FormattingEnabled = True
        Me.cmbSortOrder.Items.AddRange(New Object() {"Object ID", "Record > Date (D)", "Record > Date (A)", "Location > Date (D)", "Location > Date (A)", "Date (D)", "Date (A)", "Series Title > Date (D)", "Series Title > Date (A)"})
        Me.cmbSortOrder.Location = New System.Drawing.Point(68, 70)
        Me.cmbSortOrder.Name = "cmbSortOrder"
        Me.cmbSortOrder.Size = New System.Drawing.Size(203, 21)
        Me.cmbSortOrder.TabIndex = 6
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(7, 22)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(36, 13)
        Me.Label21.TabIndex = 0
        Me.Label21.Text = "Series"
        '
        'cmbJumpType
        '
        Me.cmbJumpType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbJumpType.FormattingEnabled = True
        Me.cmbJumpType.Items.AddRange(New Object() {"Object ID", "File", "Record", "Location"})
        Me.cmbJumpType.Location = New System.Drawing.Point(51, 44)
        Me.cmbJumpType.Name = "cmbJumpType"
        Me.cmbJumpType.Size = New System.Drawing.Size(89, 21)
        Me.cmbJumpType.TabIndex = 3
        '
        'cmdSearch
        '
        Me.cmdSearch.Location = New System.Drawing.Point(277, 69)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(75, 23)
        Me.cmdSearch.TabIndex = 7
        Me.cmdSearch.Text = "Search"
        Me.cmdSearch.UseVisualStyleBackColor = True
        '
        'txtQuery
        '
        Me.txtQuery.Location = New System.Drawing.Point(149, 44)
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(203, 20)
        Me.txtQuery.TabIndex = 4
        '
        'cmdDelete
        '
        Me.cmdDelete.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.cmdDelete.Location = New System.Drawing.Point(612, 17)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(108, 23)
        Me.cmdDelete.TabIndex = 7
        Me.cmdDelete.Text = "Delete Record"
        Me.cmdDelete.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(482, 48)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(39, 13)
        Me.Label17.TabIndex = 9
        Me.Label17.Text = "Copies"
        '
        'txtHold
        '
        Me.txtHold.Location = New System.Drawing.Point(450, 45)
        Me.txtHold.Name = "txtHold"
        Me.txtHold.Size = New System.Drawing.Size(26, 20)
        Me.txtHold.TabIndex = 8
        Me.txtHold.Text = "1"
        '
        'cmbFlash
        '
        Me.cmbFlash.Location = New System.Drawing.Point(531, 117)
        Me.cmbFlash.Name = "cmbFlash"
        Me.cmbFlash.Size = New System.Drawing.Size(75, 23)
        Me.cmbFlash.TabIndex = 15
        Me.cmbFlash.Text = "Flash"
        Me.cmbFlash.UseVisualStyleBackColor = True
        '
        'cmbZoom
        '
        Me.cmbZoom.Location = New System.Drawing.Point(612, 117)
        Me.cmbZoom.Name = "cmbZoom"
        Me.cmbZoom.Size = New System.Drawing.Size(75, 23)
        Me.cmbZoom.TabIndex = 16
        Me.cmbZoom.Text = "Zoom To"
        Me.cmbZoom.UseVisualStyleBackColor = True
        '
        'cmdEdit
        '
        Me.cmdEdit.Location = New System.Drawing.Point(612, 46)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(108, 23)
        Me.cmdEdit.TabIndex = 10
        Me.cmdEdit.Text = "Edit Record"
        Me.cmdEdit.UseVisualStyleBackColor = True
        '
        'picYes
        '
        Me.picYes.Image = CType(resources.GetObject("picYes.Image"), System.Drawing.Image)
        Me.picYes.Location = New System.Drawing.Point(370, 45)
        Me.picYes.Name = "picYes"
        Me.picYes.Size = New System.Drawing.Size(64, 64)
        Me.picYes.TabIndex = 16
        Me.picYes.TabStop = False
        '
        'cmbNext
        '
        Me.cmbNext.Location = New System.Drawing.Point(530, 70)
        Me.cmbNext.Name = "cmbNext"
        Me.cmbNext.Size = New System.Drawing.Size(75, 36)
        Me.cmbNext.TabIndex = 12
        Me.cmbNext.Text = "Next Record"
        Me.cmbNext.UseVisualStyleBackColor = True
        '
        'cmbPrev
        '
        Me.cmbPrev.Location = New System.Drawing.Point(449, 70)
        Me.cmbPrev.Name = "cmbPrev"
        Me.cmbPrev.Size = New System.Drawing.Size(75, 36)
        Me.cmbPrev.TabIndex = 11
        Me.cmbPrev.Text = "Previous Record"
        Me.cmbPrev.UseVisualStyleBackColor = True
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(369, 21)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(71, 13)
        Me.Label16.TabIndex = 2
        Me.Label16.Text = "AGSL Holds?"
        '
        'cmdNO
        '
        Me.cmdNO.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdNO.Location = New System.Drawing.Point(530, 16)
        Me.cmdNO.Name = "cmdNO"
        Me.cmdNO.Size = New System.Drawing.Size(75, 23)
        Me.cmdNO.TabIndex = 6
        Me.cmdNO.Text = "No"
        Me.cmdNO.UseVisualStyleBackColor = False
        '
        'cmdYES
        '
        Me.cmdYES.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.cmdYES.Location = New System.Drawing.Point(449, 16)
        Me.cmdYES.Name = "cmdYES"
        Me.cmdYES.Size = New System.Drawing.Size(75, 23)
        Me.cmdYES.TabIndex = 5
        Me.cmdYES.Text = "Yes"
        Me.cmdYES.UseVisualStyleBackColor = False
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.lblE)
        Me.GroupBox6.Controls.Add(Me.lblW)
        Me.GroupBox6.Controls.Add(Me.lblS)
        Me.GroupBox6.Controls.Add(Me.lblN)
        Me.GroupBox6.Location = New System.Drawing.Point(372, 358)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(356, 96)
        Me.GroupBox6.TabIndex = 4
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Spatial Index"
        '
        'lblE
        '
        Me.lblE.AutoSize = True
        Me.lblE.Location = New System.Drawing.Point(224, 49)
        Me.lblE.Name = "lblE"
        Me.lblE.Size = New System.Drawing.Size(61, 13)
        Me.lblE.TabIndex = 2
        Me.lblE.Text = "East Extent"
        '
        'lblW
        '
        Me.lblW.AutoSize = True
        Me.lblW.Location = New System.Drawing.Point(69, 49)
        Me.lblW.Name = "lblW"
        Me.lblW.Size = New System.Drawing.Size(65, 13)
        Me.lblW.TabIndex = 0
        Me.lblW.Text = "West Extent"
        '
        'lblS
        '
        Me.lblS.AutoSize = True
        Me.lblS.Location = New System.Drawing.Point(146, 75)
        Me.lblS.Name = "lblS"
        Me.lblS.Size = New System.Drawing.Size(68, 13)
        Me.lblS.TabIndex = 3
        Me.lblS.Text = "South Extent"
        '
        'lblN
        '
        Me.lblN.AutoSize = True
        Me.lblN.Location = New System.Drawing.Point(148, 22)
        Me.lblN.Name = "lblN"
        Me.lblN.Size = New System.Drawing.Size(66, 13)
        Me.lblN.TabIndex = 1
        Me.lblN.Text = "North Extent"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtPublisher)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.txtFormat)
        Me.GroupBox3.Controls.Add(Me.lblformat)
        Me.GroupBox3.Controls.Add(Me.txtPrimeMer)
        Me.GroupBox3.Controls.Add(Me.lblprimer)
        Me.GroupBox3.Controls.Add(Me.txtProjection)
        Me.GroupBox3.Controls.Add(Me.lblprojec)
        Me.GroupBox3.Controls.Add(Me.txtProduction)
        Me.GroupBox3.Controls.Add(Me.Label11)
        Me.GroupBox3.Controls.Add(Me.txtMapType)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Controls.Add(Me.txtScale)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Location = New System.Drawing.Point(370, 141)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(358, 211)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        '
        'txtPublisher
        '
        Me.txtPublisher.Location = New System.Drawing.Point(90, 24)
        Me.txtPublisher.Name = "txtPublisher"
        Me.txtPublisher.ReadOnly = True
        Me.txtPublisher.Size = New System.Drawing.Size(249, 20)
        Me.txtPublisher.TabIndex = 1
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(34, 28)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(50, 13)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Publisher"
        '
        'txtFormat
        '
        Me.txtFormat.Location = New System.Drawing.Point(90, 175)
        Me.txtFormat.Name = "txtFormat"
        Me.txtFormat.ReadOnly = True
        Me.txtFormat.Size = New System.Drawing.Size(249, 20)
        Me.txtFormat.TabIndex = 13
        '
        'lblformat
        '
        Me.lblformat.AutoSize = True
        Me.lblformat.Location = New System.Drawing.Point(45, 178)
        Me.lblformat.Name = "lblformat"
        Me.lblformat.Size = New System.Drawing.Size(39, 13)
        Me.lblformat.TabIndex = 12
        Me.lblformat.Text = "Format"
        '
        'txtPrimeMer
        '
        Me.txtPrimeMer.Location = New System.Drawing.Point(90, 150)
        Me.txtPrimeMer.Name = "txtPrimeMer"
        Me.txtPrimeMer.ReadOnly = True
        Me.txtPrimeMer.Size = New System.Drawing.Size(249, 20)
        Me.txtPrimeMer.TabIndex = 11
        '
        'lblprimer
        '
        Me.lblprimer.AutoSize = True
        Me.lblprimer.Location = New System.Drawing.Point(8, 153)
        Me.lblprimer.Name = "lblprimer"
        Me.lblprimer.Size = New System.Drawing.Size(76, 13)
        Me.lblprimer.TabIndex = 10
        Me.lblprimer.Text = "Prime Meridian"
        '
        'txtProjection
        '
        Me.txtProjection.Location = New System.Drawing.Point(90, 125)
        Me.txtProjection.Name = "txtProjection"
        Me.txtProjection.ReadOnly = True
        Me.txtProjection.Size = New System.Drawing.Size(249, 20)
        Me.txtProjection.TabIndex = 9
        '
        'lblprojec
        '
        Me.lblprojec.AutoSize = True
        Me.lblprojec.Location = New System.Drawing.Point(30, 128)
        Me.lblprojec.Name = "lblprojec"
        Me.lblprojec.Size = New System.Drawing.Size(54, 13)
        Me.lblprojec.TabIndex = 8
        Me.lblprojec.Text = "Projection"
        '
        'txtProduction
        '
        Me.txtProduction.Location = New System.Drawing.Point(90, 100)
        Me.txtProduction.Name = "txtProduction"
        Me.txtProduction.ReadOnly = True
        Me.txtProduction.Size = New System.Drawing.Size(249, 20)
        Me.txtProduction.TabIndex = 7
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(26, 103)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(58, 13)
        Me.Label11.TabIndex = 6
        Me.Label11.Text = "Production"
        '
        'txtMapType
        '
        Me.txtMapType.Location = New System.Drawing.Point(90, 75)
        Me.txtMapType.Name = "txtMapType"
        Me.txtMapType.ReadOnly = True
        Me.txtMapType.Size = New System.Drawing.Size(249, 20)
        Me.txtMapType.TabIndex = 5
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(29, 78)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(55, 13)
        Me.Label10.TabIndex = 4
        Me.Label10.Text = "Map Type"
        '
        'txtScale
        '
        Me.txtScale.Location = New System.Drawing.Point(90, 50)
        Me.txtScale.Name = "txtScale"
        Me.txtScale.ReadOnly = True
        Me.txtScale.Size = New System.Drawing.Size(249, 20)
        Me.txtScale.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(50, 53)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(34, 13)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "Scale"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtOID)
        Me.GroupBox2.Controls.Add(Me.Label19)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.txtDates)
        Me.GroupBox2.Controls.Add(Me.txtFile)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.txtSeriesTit)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.txtCataloc)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.txtEdition)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.txtDate)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.txtLocation)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.txtRecord)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 141)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(358, 313)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'txtOID
        '
        Me.txtOID.Location = New System.Drawing.Point(77, 19)
        Me.txtOID.Name = "txtOID"
        Me.txtOID.ReadOnly = True
        Me.txtOID.Size = New System.Drawing.Size(249, 20)
        Me.txtOID.TabIndex = 1
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(17, 23)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(52, 13)
        Me.Label19.TabIndex = 0
        Me.Label19.Text = "Object ID"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(31, 233)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(38, 13)
        Me.Label15.TabIndex = 16
        Me.Label15.Text = "Dates:"
        '
        'txtDates
        '
        Me.txtDates.Location = New System.Drawing.Point(77, 230)
        Me.txtDates.Multiline = True
        Me.txtDates.Name = "txtDates"
        Me.txtDates.ReadOnly = True
        Me.txtDates.Size = New System.Drawing.Size(249, 76)
        Me.txtDates.TabIndex = 17
        '
        'txtFile
        '
        Me.txtFile.Location = New System.Drawing.Point(77, 49)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.ReadOnly = True
        Me.txtFile.Size = New System.Drawing.Size(249, 20)
        Me.txtFile.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(33, 53)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(36, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Series"
        '
        'txtSeriesTit
        '
        Me.txtSeriesTit.Location = New System.Drawing.Point(77, 203)
        Me.txtSeriesTit.Name = "txtSeriesTit"
        Me.txtSeriesTit.ReadOnly = True
        Me.txtSeriesTit.Size = New System.Drawing.Size(249, 20)
        Me.txtSeriesTit.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(10, 206)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(59, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Series Title"
        '
        'txtCataloc
        '
        Me.txtCataloc.Location = New System.Drawing.Point(77, 178)
        Me.txtCataloc.Name = "txtCataloc"
        Me.txtCataloc.ReadOnly = True
        Me.txtCataloc.Size = New System.Drawing.Size(249, 20)
        Me.txtCataloc.TabIndex = 13
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(26, 181)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Catalog"
        '
        'txtEdition
        '
        Me.txtEdition.Location = New System.Drawing.Point(77, 153)
        Me.txtEdition.Name = "txtEdition"
        Me.txtEdition.ReadOnly = True
        Me.txtEdition.Size = New System.Drawing.Size(249, 20)
        Me.txtEdition.TabIndex = 11
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(30, 156)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Edition"
        '
        'txtDate
        '
        Me.txtDate.Location = New System.Drawing.Point(77, 128)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.ReadOnly = True
        Me.txtDate.Size = New System.Drawing.Size(249, 20)
        Me.txtDate.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(39, 131)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Date"
        '
        'txtLocation
        '
        Me.txtLocation.Location = New System.Drawing.Point(77, 103)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.ReadOnly = True
        Me.txtLocation.Size = New System.Drawing.Size(249, 20)
        Me.txtLocation.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 106)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Location"
        '
        'txtRecord
        '
        Me.txtRecord.Location = New System.Drawing.Point(77, 75)
        Me.txtRecord.Name = "txtRecord"
        Me.txtRecord.ReadOnly = True
        Me.txtRecord.Size = New System.Drawing.Size(249, 20)
        Me.txtRecord.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Record"
        '
        'picNo
        '
        Me.picNo.Image = CType(resources.GetObject("picNo.Image"), System.Drawing.Image)
        Me.picNo.Location = New System.Drawing.Point(370, 45)
        Me.picNo.Name = "picNo"
        Me.picNo.Size = New System.Drawing.Size(64, 64)
        Me.picNo.TabIndex = 31
        Me.picNo.TabStop = False
        Me.picNo.Visible = False
        '
        'frmReconcile
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 460)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmReconcile"
        Me.Text = "GEODEX | Reconcile Holdings"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.picYes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.picNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtSeriesTit As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtCataloc As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtEdition As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtDate As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtRecord As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtPublisher As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtFormat As System.Windows.Forms.TextBox
    Friend WithEvents lblformat As System.Windows.Forms.Label
    Friend WithEvents txtPrimeMer As System.Windows.Forms.TextBox
    Friend WithEvents lblprimer As System.Windows.Forms.Label
    Friend WithEvents txtProjection As System.Windows.Forms.TextBox
    Friend WithEvents lblprojec As System.Windows.Forms.Label
    Friend WithEvents txtProduction As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtMapType As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtScale As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtDates As System.Windows.Forms.TextBox
    Friend WithEvents cmbFlash As System.Windows.Forms.Button
    Friend WithEvents cmbZoom As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents picYes As System.Windows.Forms.PictureBox
    Friend WithEvents cmbNext As System.Windows.Forms.Button
    Friend WithEvents cmbPrev As System.Windows.Forms.Button
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cmdNO As System.Windows.Forms.Button
    Friend WithEvents cmdYES As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents lblE As System.Windows.Forms.Label
    Friend WithEvents lblW As System.Windows.Forms.Label
    Friend WithEvents lblS As System.Windows.Forms.Label
    Friend WithEvents lblN As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtHold As System.Windows.Forms.TextBox
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbJumpType As System.Windows.Forms.ComboBox
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents txtQuery As System.Windows.Forms.TextBox
    Friend WithEvents txtOID As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents cmbSortOrder As System.Windows.Forms.ComboBox
    Friend WithEvents picNo As System.Windows.Forms.PictureBox
    Friend WithEvents cmbFile As System.Windows.Forms.ComboBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents cmdGetSelect As System.Windows.Forms.Button
    Friend WithEvents cmbClear As System.Windows.Forms.Button
    Friend WithEvents cmbSelect As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents lblHold As System.Windows.Forms.Label
End Class
