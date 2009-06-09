<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SQLCompliler
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SQLCompliler))
        Me.TreeView1 = New System.Windows.Forms.TreeView
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.TSOpen = New System.Windows.Forms.ToolStripButton
        Me.TSRefresh = New System.Windows.Forms.ToolStripButton
        Me.TSView = New System.Windows.Forms.ToolStripButton
        Me.TSStart = New System.Windows.Forms.ToolStripButton
        Me.TSLicence = New System.Windows.Forms.ToolStripButton
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TreeView1
        '
        Me.TreeView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TreeView1.ImageIndex = 0
        Me.TreeView1.ImageList = Me.ImageList1
        Me.TreeView1.Location = New System.Drawing.Point(0, 28)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.SelectedImageIndex = 0
        Me.TreeView1.Size = New System.Drawing.Size(303, 385)
        Me.TreeView1.TabIndex = 0
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "db.gif")
        Me.ImageList1.Images.SetKeyName(1, "dbOkay.gif")
        Me.ImageList1.Images.SetKeyName(2, "dbError.gif")
        Me.ImageList1.Images.SetKeyName(3, "sclError.gif")
        Me.ImageList1.Images.SetKeyName(4, "sql.gif")
        Me.ImageList1.Images.SetKeyName(5, "sqlOkay.gif")
        Me.ImageList1.Images.SetKeyName(6, "sqlError.gif")
        Me.ImageList1.Images.SetKeyName(7, "unknown.gif")
        Me.ImageList1.Images.SetKeyName(8, "file.gif")
        Me.ImageList1.Images.SetKeyName(9, "fileokay.gif")
        Me.ImageList1.Images.SetKeyName(10, "fileerror.gif")
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSOpen, Me.TSRefresh, Me.TSView, Me.TSStart, Me.TSLicence})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(303, 25)
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'TSOpen
        '
        Me.TSOpen.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TSOpen.Image = CType(resources.GetObject("TSOpen.Image"), System.Drawing.Image)
        Me.TSOpen.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TSOpen.Name = "TSOpen"
        Me.TSOpen.Size = New System.Drawing.Size(23, 22)
        Me.TSOpen.Tag = "open"
        Me.TSOpen.Text = "ToolStripButton1"
        Me.TSOpen.ToolTipText = "Open SCL file."
        '
        'TSRefresh
        '
        Me.TSRefresh.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TSRefresh.Image = CType(resources.GetObject("TSRefresh.Image"), System.Drawing.Image)
        Me.TSRefresh.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TSRefresh.Name = "TSRefresh"
        Me.TSRefresh.Size = New System.Drawing.Size(23, 22)
        Me.TSRefresh.Tag = "refresh"
        Me.TSRefresh.Text = "ToolStripButton2"
        Me.TSRefresh.ToolTipText = "Reload the current file"
        '
        'TSView
        '
        Me.TSView.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TSView.Image = CType(resources.GetObject("TSView.Image"), System.Drawing.Image)
        Me.TSView.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TSView.Name = "TSView"
        Me.TSView.Size = New System.Drawing.Size(23, 22)
        Me.TSView.Tag = "view"
        Me.TSView.Text = "ToolStripButton3"
        Me.TSView.ToolTipText = "View details"
        '
        'TSStart
        '
        Me.TSStart.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TSStart.Image = CType(resources.GetObject("TSStart.Image"), System.Drawing.Image)
        Me.TSStart.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TSStart.Name = "TSStart"
        Me.TSStart.Size = New System.Drawing.Size(23, 22)
        Me.TSStart.Tag = "start"
        Me.TSStart.Text = "ToolStripButton4"
        Me.TSStart.ToolTipText = "Start compilation"
        '
        'TSLicence
        '
        Me.TSLicence.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TSLicence.Image = CType(resources.GetObject("TSLicence.Image"), System.Drawing.Image)
        Me.TSLicence.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TSLicence.Name = "TSLicence"
        Me.TSLicence.Size = New System.Drawing.Size(23, 22)
        Me.TSLicence.Text = "ToolStripButton1"
        '
        'SQLCompliler
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(303, 413)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.TreeView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "SQLCompliler"
        Me.Text = "SQLCompliler"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents TSOpen As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents TSRefresh As System.Windows.Forms.ToolStripButton
    Friend WithEvents TSView As System.Windows.Forms.ToolStripButton
    Friend WithEvents TSStart As System.Windows.Forms.ToolStripButton
    Friend WithEvents TSLicence As System.Windows.Forms.ToolStripButton
End Class
