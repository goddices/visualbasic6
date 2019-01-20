<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form31
#Region "Windows 窗体设计器生成的代码 "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'此调用是 Windows 窗体设计器所必需的。
		InitializeComponent()
		'此窗体是 MDI 子窗体。
		'此代码模拟 VB6 
		' 的自动加载和显示
		' MDI 子级的父级
		' 的功能。
		Me.MDIParent = 工程1.MDIForm1
		工程1.MDIForm1.Show
	End Sub
	'窗体重写释放，以清理组件列表。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows 窗体设计器所必需的
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents _Command1_7 As System.Windows.Forms.Button
	Public WithEvents Data1 As System.Windows.Forms.Label
	Public WithEvents _Command1_6 As System.Windows.Forms.Button
	Public WithEvents _Command1_5 As System.Windows.Forms.Button
	Public WithEvents _Command1_4 As System.Windows.Forms.Button
	Public WithEvents _Command1_3 As System.Windows.Forms.Button
	Public WithEvents _Command1_2 As System.Windows.Forms.Button
	Public WithEvents _Command1_1 As System.Windows.Forms.Button
	Public WithEvents _Command1_0 As System.Windows.Forms.Button
	Public WithEvents _Text1_2 As System.Windows.Forms.TextBox
	Public WithEvents _Text1_1 As System.Windows.Forms.TextBox
	Public WithEvents _Text1_0 As System.Windows.Forms.TextBox
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents Command1 As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Text1 As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'注意: 以下过程是 Windows 窗体设计器所必需的
	'可以使用 Windows 窗体设计器来修改它。
	'不要使用代码编辑器修改它。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form31))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me._Command1_7 = New System.Windows.Forms.Button
		Me.Data1 = New System.Windows.Forms.Label
		Me._Command1_6 = New System.Windows.Forms.Button
		Me._Command1_5 = New System.Windows.Forms.Button
		Me._Command1_4 = New System.Windows.Forms.Button
		Me._Command1_3 = New System.Windows.Forms.Button
		Me._Command1_2 = New System.Windows.Forms.Button
		Me._Command1_1 = New System.Windows.Forms.Button
		Me._Command1_0 = New System.Windows.Forms.Button
		Me._Text1_2 = New System.Windows.Forms.TextBox
		Me._Text1_1 = New System.Windows.Forms.TextBox
		Me._Text1_0 = New System.Windows.Forms.TextBox
		Me._Label1_0 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_2 = New System.Windows.Forms.Label
		Me.Command1 = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Text1 = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Command1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Text1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.FromARGB(0, 64, 64)
		Me.Text = "编辑用户"
		Me.ClientSize = New System.Drawing.Size(465, 322)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "Form31"
		Me._Command1_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Command1_7.Text = "刷  新"
		Me._Command1_7.Size = New System.Drawing.Size(67, 25)
		Me._Command1_7.Location = New System.Drawing.Point(288, 192)
		Me._Command1_7.TabIndex = 13
		Me._Command1_7.BackColor = System.Drawing.SystemColors.Control
		Me._Command1_7.CausesValidation = True
		Me._Command1_7.Enabled = True
		Me._Command1_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Command1_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._Command1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Command1_7.TabStop = True
		Me._Command1_7.Name = "_Command1_7"
		Me.Data1.Text = "Data1"
		Me.Data1.Size = New System.Drawing.Size(81, 33)
		Me.Data1.Location = New System.Drawing.Point(352, 264)
		Me.Data1.Visible = False
		Me.Data1.BackColor = System.Drawing.Color.Red
		Me.Data1.ForeColor = System.Drawing.Color.Black
		Me.Data1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Data1.Text = "Data1"
		Me.Data1.Name = "Data1"
		Me._Command1_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Command1_6.Text = "修  改"
		Me._Command1_6.Size = New System.Drawing.Size(67, 25)
		Me._Command1_6.Location = New System.Drawing.Point(224, 192)
		Me._Command1_6.TabIndex = 12
		Me._Command1_6.BackColor = System.Drawing.SystemColors.Control
		Me._Command1_6.CausesValidation = True
		Me._Command1_6.Enabled = True
		Me._Command1_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Command1_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._Command1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Command1_6.TabStop = True
		Me._Command1_6.Name = "_Command1_6"
		Me._Command1_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Command1_5.Text = "删  除"
		Me._Command1_5.Size = New System.Drawing.Size(67, 25)
		Me._Command1_5.Location = New System.Drawing.Point(160, 192)
		Me._Command1_5.TabIndex = 11
		Me._Command1_5.BackColor = System.Drawing.SystemColors.Control
		Me._Command1_5.CausesValidation = True
		Me._Command1_5.Enabled = True
		Me._Command1_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Command1_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Command1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Command1_5.TabStop = True
		Me._Command1_5.Name = "_Command1_5"
		Me._Command1_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Command1_4.Text = "增  加"
		Me._Command1_4.Size = New System.Drawing.Size(67, 25)
		Me._Command1_4.Location = New System.Drawing.Point(96, 192)
		Me._Command1_4.TabIndex = 10
		Me._Command1_4.BackColor = System.Drawing.SystemColors.Control
		Me._Command1_4.CausesValidation = True
		Me._Command1_4.Enabled = True
		Me._Command1_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Command1_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Command1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Command1_4.TabStop = True
		Me._Command1_4.Name = "_Command1_4"
		Me._Command1_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Command1_3.Text = "末  条"
		Me._Command1_3.Size = New System.Drawing.Size(67, 25)
		Me._Command1_3.Location = New System.Drawing.Point(288, 248)
		Me._Command1_3.TabIndex = 9
		Me._Command1_3.BackColor = System.Drawing.SystemColors.Control
		Me._Command1_3.CausesValidation = True
		Me._Command1_3.Enabled = True
		Me._Command1_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Command1_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Command1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Command1_3.TabStop = True
		Me._Command1_3.Name = "_Command1_3"
		Me._Command1_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Command1_2.Text = "前一条"
		Me._Command1_2.Enabled = False
		Me._Command1_2.Size = New System.Drawing.Size(67, 25)
		Me._Command1_2.Location = New System.Drawing.Point(224, 248)
		Me._Command1_2.TabIndex = 8
		Me._Command1_2.BackColor = System.Drawing.SystemColors.Control
		Me._Command1_2.CausesValidation = True
		Me._Command1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Command1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Command1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Command1_2.TabStop = True
		Me._Command1_2.Name = "_Command1_2"
		Me._Command1_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Command1_1.Text = "下一条"
		Me._Command1_1.Size = New System.Drawing.Size(67, 25)
		Me._Command1_1.Location = New System.Drawing.Point(160, 248)
		Me._Command1_1.TabIndex = 7
		Me._Command1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Command1_1.CausesValidation = True
		Me._Command1_1.Enabled = True
		Me._Command1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Command1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Command1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Command1_1.TabStop = True
		Me._Command1_1.Name = "_Command1_1"
		Me._Command1_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._Command1_0.Text = "首  条"
		Me._Command1_0.Size = New System.Drawing.Size(67, 25)
		Me._Command1_0.Location = New System.Drawing.Point(96, 248)
		Me._Command1_0.TabIndex = 6
		Me._Command1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Command1_0.CausesValidation = True
		Me._Command1_0.Enabled = True
		Me._Command1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Command1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Command1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Command1_0.TabStop = True
		Me._Command1_0.Name = "_Command1_0"
		Me._Text1_2.AutoSize = False
		Me._Text1_2.Font = New System.Drawing.Font("宋体", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me._Text1_2.Size = New System.Drawing.Size(225, 25)
		Me._Text1_2.Location = New System.Drawing.Point(176, 120)
		Me._Text1_2.TabIndex = 5
		Me._Text1_2.AcceptsReturn = True
		Me._Text1_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Text1_2.BackColor = System.Drawing.SystemColors.Window
		Me._Text1_2.CausesValidation = True
		Me._Text1_2.Enabled = True
		Me._Text1_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Text1_2.HideSelection = True
		Me._Text1_2.ReadOnly = False
		Me._Text1_2.Maxlength = 0
		Me._Text1_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Text1_2.MultiLine = False
		Me._Text1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Text1_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Text1_2.TabStop = True
		Me._Text1_2.Visible = True
		Me._Text1_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Text1_2.Name = "_Text1_2"
		Me._Text1_1.AutoSize = False
		Me._Text1_1.Font = New System.Drawing.Font("宋体", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me._Text1_1.Size = New System.Drawing.Size(89, 25)
		Me._Text1_1.Location = New System.Drawing.Point(176, 80)
		Me._Text1_1.TabIndex = 4
		Me._Text1_1.AcceptsReturn = True
		Me._Text1_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Text1_1.BackColor = System.Drawing.SystemColors.Window
		Me._Text1_1.CausesValidation = True
		Me._Text1_1.Enabled = True
		Me._Text1_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Text1_1.HideSelection = True
		Me._Text1_1.ReadOnly = False
		Me._Text1_1.Maxlength = 0
		Me._Text1_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Text1_1.MultiLine = False
		Me._Text1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Text1_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Text1_1.TabStop = True
		Me._Text1_1.Visible = True
		Me._Text1_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Text1_1.Name = "_Text1_1"
		Me._Text1_0.AutoSize = False
		Me._Text1_0.Font = New System.Drawing.Font("宋体", 12!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me._Text1_0.Size = New System.Drawing.Size(89, 25)
		Me._Text1_0.Location = New System.Drawing.Point(176, 40)
		Me._Text1_0.TabIndex = 3
		Me._Text1_0.AcceptsReturn = True
		Me._Text1_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Text1_0.BackColor = System.Drawing.SystemColors.Window
		Me._Text1_0.CausesValidation = True
		Me._Text1_0.Enabled = True
		Me._Text1_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Text1_0.HideSelection = True
		Me._Text1_0.ReadOnly = False
		Me._Text1_0.Maxlength = 0
		Me._Text1_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Text1_0.MultiLine = False
		Me._Text1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Text1_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Text1_0.TabStop = True
		Me._Text1_0.Visible = True
		Me._Text1_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Text1_0.Name = "_Text1_0"
		Me._Label1_0.Text = "总户号："
		Me._Label1_0.Font = New System.Drawing.Font("宋体", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me._Label1_0.ForeColor = System.Drawing.Color.Yellow
		Me._Label1_0.Size = New System.Drawing.Size(68, 16)
		Me._Label1_0.Location = New System.Drawing.Point(105, 40)
		Me._Label1_0.TabIndex = 2
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.BackColor = System.Drawing.Color.Transparent
		Me._Label1_0.Enabled = True
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = True
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		Me._Label1_1.Text = "户  名："
		Me._Label1_1.Font = New System.Drawing.Font("宋体", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me._Label1_1.ForeColor = System.Drawing.Color.Yellow
		Me._Label1_1.Size = New System.Drawing.Size(69, 16)
		Me._Label1_1.Location = New System.Drawing.Point(104, 82)
		Me._Label1_1.TabIndex = 1
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.BackColor = System.Drawing.Color.Transparent
		Me._Label1_1.Enabled = True
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = True
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_2.Text = "地  址："
		Me._Label1_2.Font = New System.Drawing.Font("宋体", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me._Label1_2.ForeColor = System.Drawing.Color.Yellow
		Me._Label1_2.Size = New System.Drawing.Size(69, 16)
		Me._Label1_2.Location = New System.Drawing.Point(104, 125)
		Me._Label1_2.TabIndex = 0
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_2.BackColor = System.Drawing.Color.Transparent
		Me._Label1_2.Enabled = True
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = True
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		Me.Command1.SetIndex(_Command1_7, CType(7, Short))
		Me.Command1.SetIndex(_Command1_6, CType(6, Short))
		Me.Command1.SetIndex(_Command1_5, CType(5, Short))
		Me.Command1.SetIndex(_Command1_4, CType(4, Short))
		Me.Command1.SetIndex(_Command1_3, CType(3, Short))
		Me.Command1.SetIndex(_Command1_2, CType(2, Short))
		Me.Command1.SetIndex(_Command1_1, CType(1, Short))
		Me.Command1.SetIndex(_Command1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Text1.SetIndex(_Text1_2, CType(2, Short))
		Me.Text1.SetIndex(_Text1_1, CType(1, Short))
		Me.Text1.SetIndex(_Text1_0, CType(0, Short))
		CType(Me.Text1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Command1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(_Command1_7)
		Me.Controls.Add(Data1)
		Me.Controls.Add(_Command1_6)
		Me.Controls.Add(_Command1_5)
		Me.Controls.Add(_Command1_4)
		Me.Controls.Add(_Command1_3)
		Me.Controls.Add(_Command1_2)
		Me.Controls.Add(_Command1_1)
		Me.Controls.Add(_Command1_0)
		Me.Controls.Add(_Text1_2)
		Me.Controls.Add(_Text1_1)
		Me.Controls.Add(_Text1_0)
		Me.Controls.Add(_Label1_0)
		Me.Controls.Add(_Label1_1)
		Me.Controls.Add(_Label1_2)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class