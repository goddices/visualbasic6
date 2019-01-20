<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class MDIForm1
#Region "Windows 窗体设计器生成的代码 "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'此调用是 Windows 窗体设计器所必需的。
		InitializeComponent()
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
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents StatusBar1 As AxComctlLib.AxStatusBar
	Public WithEvents DL As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mmxg As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents aaa As System.Windows.Forms.ToolStripSeparator
	Public WithEvents TC As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents XT As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents jnsf As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents cxsf As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents sfgl As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents bjyf As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents llyf As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents yfgl As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents DRjfqh As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents DYFW As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents dqjg As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents hjsz As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	'注意: 以下过程是 Windows 窗体设计器所必需的
	'可以使用 Windows 窗体设计器来修改它。
	'不要使用代码编辑器修改它。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MDIForm1))
		Me.IsMDIContainer = True
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.StatusBar1 = New AxComctlLib.AxStatusBar
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.XT = New System.Windows.Forms.ToolStripMenuItem
		Me.DL = New System.Windows.Forms.ToolStripMenuItem
		Me.mmxg = New System.Windows.Forms.ToolStripMenuItem
		Me.aaa = New System.Windows.Forms.ToolStripSeparator
		Me.TC = New System.Windows.Forms.ToolStripMenuItem
		Me.sfgl = New System.Windows.Forms.ToolStripMenuItem
		Me.jnsf = New System.Windows.Forms.ToolStripMenuItem
		Me.cxsf = New System.Windows.Forms.ToolStripMenuItem
		Me.yfgl = New System.Windows.Forms.ToolStripMenuItem
		Me.bjyf = New System.Windows.Forms.ToolStripMenuItem
		Me.llyf = New System.Windows.Forms.ToolStripMenuItem
		Me.DYFW = New System.Windows.Forms.ToolStripMenuItem
		Me.DRjfqh = New System.Windows.Forms.ToolStripMenuItem
		Me.hjsz = New System.Windows.Forms.ToolStripMenuItem
		Me.dqjg = New System.Windows.Forms.ToolStripMenuItem
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.StatusBar1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.BackColor = System.Drawing.SystemColors.AppWorkspace
		Me.Text = "苏州水费管理系统"
		Me.ClientSize = New System.Drawing.Size(649, 528)
		Me.Location = New System.Drawing.Point(11, 37)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Name = "MDIForm1"
		Me.Timer1.Interval = 1000
		Me.Timer1.Enabled = True
		StatusBar1.OcxState = CType(resources.GetObject("StatusBar1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.StatusBar1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.StatusBar1.Size = New System.Drawing.Size(649, 25)
		Me.StatusBar1.Location = New System.Drawing.Point(0, 503)
		Me.StatusBar1.TabIndex = 0
		Me.StatusBar1.Name = "StatusBar1"
		Me.XT.Name = "XT"
		Me.XT.Text = "系统"
		Me.XT.Checked = False
		Me.XT.Enabled = True
		Me.XT.Visible = True
		Me.DL.Name = "DL"
		Me.DL.Text = "登录"
		Me.DL.Checked = False
		Me.DL.Enabled = True
		Me.DL.Visible = True
		Me.mmxg.Name = "mmxg"
		Me.mmxg.Text = "帐户管理"
		Me.mmxg.Enabled = False
		Me.mmxg.Checked = False
		Me.mmxg.Visible = True
		Me.aaa.Enabled = True
		Me.aaa.Visible = True
		Me.aaa.Name = "aaa"
		Me.TC.Name = "TC"
		Me.TC.Text = "退出"
		Me.TC.Checked = False
		Me.TC.Enabled = True
		Me.TC.Visible = True
		Me.sfgl.Name = "sfgl"
		Me.sfgl.Text = "水费管理"
		Me.sfgl.Enabled = False
		Me.sfgl.Checked = False
		Me.sfgl.Visible = True
		Me.jnsf.Name = "jnsf"
		Me.jnsf.Text = "缴纳水费"
		Me.jnsf.Checked = False
		Me.jnsf.Enabled = True
		Me.jnsf.Visible = True
		Me.cxsf.Name = "cxsf"
		Me.cxsf.Text = "查询缴费情况"
		Me.cxsf.Checked = False
		Me.cxsf.Enabled = True
		Me.cxsf.Visible = True
		Me.yfgl.Name = "yfgl"
		Me.yfgl.Text = "用户管理"
		Me.yfgl.Enabled = False
		Me.yfgl.Checked = False
		Me.yfgl.Visible = True
		Me.bjyf.Name = "bjyf"
		Me.bjyf.Text = "编辑用户"
		Me.bjyf.Checked = False
		Me.bjyf.Enabled = True
		Me.bjyf.Visible = True
		Me.llyf.Name = "llyf"
		Me.llyf.Text = "浏览用户"
		Me.llyf.Checked = False
		Me.llyf.Enabled = True
		Me.llyf.Visible = True
		Me.DYFW.Name = "DYFW"
		Me.DYFW.Text = "打印服务"
		Me.DYFW.Enabled = False
		Me.DYFW.Checked = False
		Me.DYFW.Visible = True
		Me.DRjfqh.Name = "DRjfqh"
		Me.DRjfqh.Text = "当日缴费情况"
		Me.DRjfqh.Checked = False
		Me.DRjfqh.Enabled = True
		Me.DRjfqh.Visible = True
		Me.hjsz.Name = "hjsz"
		Me.hjsz.Text = "环境设置"
		Me.hjsz.Enabled = False
		Me.hjsz.Checked = False
		Me.hjsz.Visible = True
		Me.dqjg.Name = "dqjg"
		Me.dqjg.Text = "当前水费价格"
		Me.dqjg.Checked = False
		Me.dqjg.Enabled = True
		Me.dqjg.Visible = True
		Me.Controls.Add(StatusBar1)
		CType(Me.StatusBar1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.XT.MergeAction = System.Windows.Forms.MergeAction.Remove
		Me.sfgl.MergeAction = System.Windows.Forms.MergeAction.Remove
		Me.yfgl.MergeAction = System.Windows.Forms.MergeAction.Remove
		Me.DYFW.MergeAction = System.Windows.Forms.MergeAction.Remove
		Me.hjsz.MergeAction = System.Windows.Forms.MergeAction.Remove
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.XT, Me.sfgl, Me.yfgl, Me.DYFW, Me.hjsz})
		XT.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.DL, Me.mmxg, Me.aaa, Me.TC})
		sfgl.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.jnsf, Me.cxsf})
		yfgl.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.bjyf, Me.llyf})
		DYFW.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.DRjfqh})
		hjsz.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.dqjg})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class