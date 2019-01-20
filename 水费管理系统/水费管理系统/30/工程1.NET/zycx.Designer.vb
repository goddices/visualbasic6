<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class Form22
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
	Public WithEvents _Option1_0 As System.Windows.Forms.RadioButton
	Public WithEvents _Option1_1 As System.Windows.Forms.RadioButton
	Public WithEvents _Option1_2 As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents DTPicker1 As AxMSComCtl2.AxDTPicker
	Public WithEvents DBGrid1 As AxMSDBGrid.AxDBGrid
	Public WithEvents Data1 As System.Windows.Forms.Label
	Public WithEvents Option1 As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'注意: 以下过程是 Windows 窗体设计器所必需的
	'可以使用 Windows 窗体设计器来修改它。
	'不要使用代码编辑器修改它。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form22))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me._Option1_0 = New System.Windows.Forms.RadioButton
		Me._Option1_1 = New System.Windows.Forms.RadioButton
		Me._Option1_2 = New System.Windows.Forms.RadioButton
		Me.Command1 = New System.Windows.Forms.Button
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.DTPicker1 = New AxMSComCtl2.AxDTPicker
		Me.DBGrid1 = New AxMSDBGrid.AxDBGrid
		Me.Data1 = New System.Windows.Forms.Label
		Me.Option1 = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.DTPicker1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.DBGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Option1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.Color.FromARGB(0, 64, 64)
		Me.Text = "查询缴费情况"
		Me.ClientSize = New System.Drawing.Size(592, 401)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.ForeColor = System.Drawing.Color.FromARGB(0, 64, 64)
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
		Me.Name = "Form22"
		Me.Frame1.BackColor = System.Drawing.Color.FromARGB(0, 64, 64)
		Me.Frame1.Text = "选择查询项"
		Me.Frame1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me.Frame1.ForeColor = System.Drawing.Color.Yellow
		Me.Frame1.Size = New System.Drawing.Size(121, 137)
		Me.Frame1.Location = New System.Drawing.Point(64, 8)
		Me.Frame1.TabIndex = 4
		Me.Frame1.Enabled = True
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Name = "Frame1"
		Me._Option1_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_0.BackColor = System.Drawing.Color.FromARGB(0, 64, 64)
		Me._Option1_0.Text = "总户号"
		Me._Option1_0.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me._Option1_0.ForeColor = System.Drawing.Color.Yellow
		Me._Option1_0.Size = New System.Drawing.Size(65, 17)
		Me._Option1_0.Location = New System.Drawing.Point(24, 24)
		Me._Option1_0.TabIndex = 7
		Me._Option1_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_0.CausesValidation = True
		Me._Option1_0.Enabled = True
		Me._Option1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Option1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Option1_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Option1_0.TabStop = True
		Me._Option1_0.Checked = False
		Me._Option1_0.Visible = True
		Me._Option1_0.Name = "_Option1_0"
		Me._Option1_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_1.BackColor = System.Drawing.Color.FromARGB(0, 64, 64)
		Me._Option1_1.Text = "户  名"
		Me._Option1_1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me._Option1_1.ForeColor = System.Drawing.Color.Yellow
		Me._Option1_1.Size = New System.Drawing.Size(73, 17)
		Me._Option1_1.Location = New System.Drawing.Point(24, 64)
		Me._Option1_1.TabIndex = 6
		Me._Option1_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_1.CausesValidation = True
		Me._Option1_1.Enabled = True
		Me._Option1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Option1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Option1_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Option1_1.TabStop = True
		Me._Option1_1.Checked = False
		Me._Option1_1.Visible = True
		Me._Option1_1.Name = "_Option1_1"
		Me._Option1_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_2.BackColor = System.Drawing.Color.FromARGB(0, 64, 64)
		Me._Option1_2.Text = "缴费日期"
		Me._Option1_2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
		Me._Option1_2.ForeColor = System.Drawing.Color.Yellow
		Me._Option1_2.Size = New System.Drawing.Size(81, 17)
		Me._Option1_2.Location = New System.Drawing.Point(24, 104)
		Me._Option1_2.TabIndex = 5
		Me._Option1_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_2.CausesValidation = True
		Me._Option1_2.Enabled = True
		Me._Option1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Option1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Option1_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Option1_2.TabStop = True
		Me._Option1_2.Checked = False
		Me._Option1_2.Visible = True
		Me._Option1_2.Name = "_Option1_2"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.Command1.Text = "查  询"
		Me.Command1.Size = New System.Drawing.Size(73, 33)
		Me.Command1.Location = New System.Drawing.Point(480, 56)
		Me.Command1.TabIndex = 3
		Me.Command1.Visible = False
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(97, 25)
		Me.Text1.Location = New System.Drawing.Point(200, 40)
		Me.Text1.TabIndex = 2
		Me.Text1.Visible = False
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.BackColor = System.Drawing.SystemColors.Window
		Me.Text1.CausesValidation = True
		Me.Text1.Enabled = True
		Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text1.HideSelection = True
		Me.Text1.ReadOnly = False
		Me.Text1.Maxlength = 0
		Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text1.MultiLine = False
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.TabStop = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		DTPicker1.OcxState = CType(resources.GetObject("DTPicker1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DTPicker1.Size = New System.Drawing.Size(89, 25)
		Me.DTPicker1.Location = New System.Drawing.Point(200, 104)
		Me.DTPicker1.TabIndex = 1
		Me.DTPicker1.Visible = False
		Me.DTPicker1.Name = "DTPicker1"
		DBGrid1.OcxState = CType(resources.GetObject("DBGrid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.DBGrid1.Size = New System.Drawing.Size(641, 233)
		Me.DBGrid1.Location = New System.Drawing.Point(24, 152)
		Me.DBGrid1.TabIndex = 0
		Me.DBGrid1.Name = "DBGrid1"
		Me.Data1.Text = "Data1"
		Me.Data1.Size = New System.Drawing.Size(76, 33)
		Me.Data1.Location = New System.Drawing.Point(72, 360)
		Me.Data1.Visible = False
		Me.Data1.BackColor = System.Drawing.Color.Red
		Me.Data1.ForeColor = System.Drawing.Color.Black
		Me.Data1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.Data1.Text = "Data1"
		Me.Data1.Name = "Data1"
		Me.Option1.SetIndex(_Option1_0, CType(0, Short))
		Me.Option1.SetIndex(_Option1_1, CType(1, Short))
		Me.Option1.SetIndex(_Option1_2, CType(2, Short))
		CType(Me.Option1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DBGrid1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.DTPicker1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(Frame1)
		Me.Controls.Add(Command1)
		Me.Controls.Add(Text1)
		Me.Controls.Add(DTPicker1)
		Me.Controls.Add(DBGrid1)
		Me.Controls.Add(Data1)
		Me.Frame1.Controls.Add(_Option1_0)
		Me.Frame1.Controls.Add(_Option1_1)
		Me.Frame1.Controls.Add(_Option1_2)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class