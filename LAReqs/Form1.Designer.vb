<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.lstStatus = New System.Windows.Forms.ListBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.lblStatus = New System.Windows.Forms.ToolStripStatusLabel
        Me.pgrWorking = New System.Windows.Forms.ProgressBar
        Me.BGOR_GBP = New System.ComponentModel.BackgroundWorker
        Me.BGOR_L7P = New System.ComponentModel.BackgroundWorker
        Me.BGOR_L6P = New System.ComponentModel.BackgroundWorker
        Me.BGOR_G4P = New System.ComponentModel.BackgroundWorker
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lstStatus
        '
        Me.lstStatus.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstStatus.FormattingEnabled = True
        Me.lstStatus.Location = New System.Drawing.Point(5, 7)
        Me.lstStatus.Name = "lstStatus"
        Me.lstStatus.Size = New System.Drawing.Size(452, 199)
        Me.lstStatus.TabIndex = 4
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblStatus})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 218)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(462, 22)
        Me.StatusStrip1.TabIndex = 5
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'lblStatus
        '
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(447, 17)
        Me.lblStatus.Spring = True
        Me.lblStatus.Text = "ToolStripStatusLabel1"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pgrWorking
        '
        Me.pgrWorking.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pgrWorking.Location = New System.Drawing.Point(0, 208)
        Me.pgrWorking.Name = "pgrWorking"
        Me.pgrWorking.Size = New System.Drawing.Size(462, 10)
        Me.pgrWorking.TabIndex = 6
        '
        'BGOR_GBP
        '
        Me.BGOR_GBP.WorkerReportsProgress = True
        '
        'BGOR_L7P
        '
        Me.BGOR_L7P.WorkerReportsProgress = True
        '
        'BGOR_L6P
        '
        Me.BGOR_L6P.WorkerReportsProgress = True
        '
        'BGOR_G4P
        '
        Me.BGOR_G4P.WorkerReportsProgress = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(462, 240)
        Me.Controls.Add(Me.pgrWorking)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.lstStatus)
        Me.Name = "Form1"
        Me.Text = "LA Open Requisitions [DMS]"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstStatus As System.Windows.Forms.ListBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents lblStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents pgrWorking As System.Windows.Forms.ProgressBar
    Friend WithEvents BGOR_GBP As System.ComponentModel.BackgroundWorker
    Friend WithEvents BGOR_L7P As System.ComponentModel.BackgroundWorker
    Friend WithEvents BGOR_L6P As System.ComponentModel.BackgroundWorker
    Friend WithEvents BGOR_G4P As System.ComponentModel.BackgroundWorker

End Class
