<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewer
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmViewer))
        Me.cmbFile = New System.Windows.Forms.ComboBox()
        Me.cmbSheet = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btRefresh = New System.Windows.Forms.Button()
        Me.rbRecord = New System.Windows.Forms.RadioButton()
        Me.rbLocation = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbFields = New System.Windows.Forms.ComboBox()
        Me.tbFilter = New System.Windows.Forms.TextBox()
        Me.cbOp = New System.Windows.Forms.ComboBox()
        Me.btFilter = New System.Windows.Forms.Button()
        Me.btSelect = New System.Windows.Forms.Button()
        Me.btFlash = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmbFile
        '
        Me.cmbFile.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbFile.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbFile.FormattingEnabled = True
        Me.cmbFile.Location = New System.Drawing.Point(54, 5)
        Me.cmbFile.Name = "cmbFile"
        Me.cmbFile.Size = New System.Drawing.Size(455, 21)
        Me.cmbFile.Sorted = True
        Me.cmbFile.TabIndex = 1
        '
        'cmbSheet
        '
        Me.cmbSheet.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbSheet.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbSheet.FormattingEnabled = True
        Me.cmbSheet.Location = New System.Drawing.Point(53, 73)
        Me.cmbSheet.Name = "cmbSheet"
        Me.cmbSheet.Size = New System.Drawing.Size(456, 21)
        Me.cmbSheet.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Series"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 77)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Sheet"
        '
        'btRefresh
        '
        Me.btRefresh.Location = New System.Drawing.Point(434, 102)
        Me.btRefresh.Name = "btRefresh"
        Me.btRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btRefresh.TabIndex = 13
        Me.btRefresh.Text = "Refresh"
        Me.btRefresh.UseVisualStyleBackColor = True
        '
        'rbRecord
        '
        Me.rbRecord.AutoSize = True
        Me.rbRecord.Location = New System.Drawing.Point(165, 105)
        Me.rbRecord.Name = "rbRecord"
        Me.rbRecord.Size = New System.Drawing.Size(119, 17)
        Me.rbRecord.TabIndex = 11
        Me.rbRecord.TabStop = True
        Me.rbRecord.Text = "Show ""Record"" first"
        Me.rbRecord.UseVisualStyleBackColor = True
        '
        'rbLocation
        '
        Me.rbLocation.AutoSize = True
        Me.rbLocation.Location = New System.Drawing.Point(298, 105)
        Me.rbLocation.Name = "rbLocation"
        Me.rbLocation.Size = New System.Drawing.Size(125, 17)
        Me.rbLocation.TabIndex = 12
        Me.rbLocation.TabStop = True
        Me.rbLocation.Text = "Show ""Location"" first"
        Me.rbLocation.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Filter (optional)"
        '
        'cbFields
        '
        Me.cbFields.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbFields.FormattingEnabled = True
        Me.cbFields.Location = New System.Drawing.Point(88, 39)
        Me.cbFields.Name = "cbFields"
        Me.cbFields.Size = New System.Drawing.Size(121, 21)
        Me.cbFields.TabIndex = 3
        '
        'tbFilter
        '
        Me.tbFilter.Location = New System.Drawing.Point(261, 39)
        Me.tbFilter.Name = "tbFilter"
        Me.tbFilter.Size = New System.Drawing.Size(167, 20)
        Me.tbFilter.TabIndex = 5
        '
        'cbOp
        '
        Me.cbOp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbOp.FormattingEnabled = True
        Me.cbOp.Items.AddRange(New Object() {"=", "<>", ">", "<", "<=", ">=", "Like", "Is Not", "Is", "In"})
        Me.cbOp.Location = New System.Drawing.Point(216, 39)
        Me.cbOp.Name = "cbOp"
        Me.cbOp.Size = New System.Drawing.Size(39, 21)
        Me.cbOp.TabIndex = 4
        '
        'btFilter
        '
        Me.btFilter.Location = New System.Drawing.Point(434, 38)
        Me.btFilter.Name = "btFilter"
        Me.btFilter.Size = New System.Drawing.Size(75, 23)
        Me.btFilter.TabIndex = 6
        Me.btFilter.Text = "Filter"
        Me.btFilter.UseVisualStyleBackColor = True
        '
        'btSelect
        '
        Me.btSelect.Location = New System.Drawing.Point(15, 102)
        Me.btSelect.Name = "btSelect"
        Me.btSelect.Size = New System.Drawing.Size(61, 23)
        Me.btSelect.TabIndex = 9
        Me.btSelect.Text = "Select"
        Me.btSelect.UseVisualStyleBackColor = True
        '
        'btFlash
        '
        Me.btFlash.Location = New System.Drawing.Point(90, 102)
        Me.btFlash.Name = "btFlash"
        Me.btFlash.Size = New System.Drawing.Size(61, 23)
        Me.btFlash.TabIndex = 10
        Me.btFlash.Text = "Flash"
        Me.btFlash.UseVisualStyleBackColor = True
        '
        'frmViewer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(520, 133)
        Me.Controls.Add(Me.btFlash)
        Me.Controls.Add(Me.btSelect)
        Me.Controls.Add(Me.btFilter)
        Me.Controls.Add(Me.cbOp)
        Me.Controls.Add(Me.tbFilter)
        Me.Controls.Add(Me.cbFields)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.rbLocation)
        Me.Controls.Add(Me.rbRecord)
        Me.Controls.Add(Me.btRefresh)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbSheet)
        Me.Controls.Add(Me.cmbFile)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmViewer"
        Me.Text = "Record Viewer"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbFile As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSheet As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btRefresh As System.Windows.Forms.Button
    Friend WithEvents rbRecord As System.Windows.Forms.RadioButton
    Friend WithEvents rbLocation As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbFields As System.Windows.Forms.ComboBox
    Friend WithEvents tbFilter As System.Windows.Forms.TextBox
    Friend WithEvents cbOp As System.Windows.Forms.ComboBox
    Friend WithEvents btFilter As System.Windows.Forms.Button
    Friend WithEvents btSelect As System.Windows.Forms.Button
    Friend WithEvents btFlash As System.Windows.Forms.Button
End Class
