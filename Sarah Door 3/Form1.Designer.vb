<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblStatus = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblDuration = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblConfidence = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblLastInteraction = New System.Windows.Forms.Label
        Me.lblCommand = New System.Windows.Forms.Label
        Me.btnOn = New System.Windows.Forms.Button
        Me.btnOff = New System.Windows.Forms.Button
        Me.btnDoorbell = New System.Windows.Forms.Button
        Me.btnLastVisitor = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 177)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 13)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Last Talk:"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(95, 153)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(48, 13)
        Me.lblStatus.TabIndex = 20
        Me.lblStatus.Text = "Sleeping"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 153)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 13)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Status:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 13)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Last Command:"
        '
        'lblDuration
        '
        Me.lblDuration.AutoSize = True
        Me.lblDuration.Location = New System.Drawing.Point(95, 131)
        Me.lblDuration.Name = "lblDuration"
        Me.lblDuration.Size = New System.Drawing.Size(13, 13)
        Me.lblDuration.TabIndex = 17
        Me.lblDuration.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 131)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 13)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Duration:"
        '
        'lblConfidence
        '
        Me.lblConfidence.AutoSize = True
        Me.lblConfidence.Location = New System.Drawing.Point(95, 109)
        Me.lblConfidence.Name = "lblConfidence"
        Me.lblConfidence.Size = New System.Drawing.Size(13, 13)
        Me.lblConfidence.TabIndex = 15
        Me.lblConfidence.Text = "0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 109)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Confidence:"
        '
        'lblLastInteraction
        '
        Me.lblLastInteraction.AutoSize = True
        Me.lblLastInteraction.Location = New System.Drawing.Point(95, 177)
        Me.lblLastInteraction.Name = "lblLastInteraction"
        Me.lblLastInteraction.Size = New System.Drawing.Size(0, 13)
        Me.lblLastInteraction.TabIndex = 22
        '
        'lblCommand
        '
        Me.lblCommand.AutoSize = True
        Me.lblCommand.Location = New System.Drawing.Point(95, 86)
        Me.lblCommand.Name = "lblCommand"
        Me.lblCommand.Size = New System.Drawing.Size(0, 13)
        Me.lblCommand.TabIndex = 23
        '
        'btnOn
        '
        Me.btnOn.Enabled = False
        Me.btnOn.Location = New System.Drawing.Point(148, 12)
        Me.btnOn.Name = "btnOn"
        Me.btnOn.Size = New System.Drawing.Size(58, 24)
        Me.btnOn.TabIndex = 24
        Me.btnOn.Text = "On"
        Me.btnOn.UseVisualStyleBackColor = True
        '
        'btnOff
        '
        Me.btnOff.Location = New System.Drawing.Point(148, 42)
        Me.btnOff.Name = "btnOff"
        Me.btnOff.Size = New System.Drawing.Size(58, 24)
        Me.btnOff.TabIndex = 25
        Me.btnOff.Text = "Off"
        Me.btnOff.UseVisualStyleBackColor = True
        '
        'btnDoorbell
        '
        Me.btnDoorbell.Location = New System.Drawing.Point(12, 12)
        Me.btnDoorbell.Name = "btnDoorbell"
        Me.btnDoorbell.Size = New System.Drawing.Size(109, 24)
        Me.btnDoorbell.TabIndex = 26
        Me.btnDoorbell.Text = "Doorbell"
        Me.btnDoorbell.UseVisualStyleBackColor = True
        '
        'btnLastVisitor
        '
        Me.btnLastVisitor.Location = New System.Drawing.Point(12, 42)
        Me.btnLastVisitor.Name = "btnLastVisitor"
        Me.btnLastVisitor.Size = New System.Drawing.Size(108, 24)
        Me.btnLastVisitor.TabIndex = 27
        Me.btnLastVisitor.Text = "Why Stopped By?"
        Me.btnLastVisitor.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(230, 196)
        Me.Controls.Add(Me.btnLastVisitor)
        Me.Controls.Add(Me.btnDoorbell)
        Me.Controls.Add(Me.btnOff)
        Me.Controls.Add(Me.btnOn)
        Me.Controls.Add(Me.lblCommand)
        Me.Controls.Add(Me.lblLastInteraction)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblDuration)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lblConfidence)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmMain"
        Me.Text = "S.A.R.A.H. Front Door"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblDuration As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblConfidence As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblLastInteraction As System.Windows.Forms.Label
    Friend WithEvents lblCommand As System.Windows.Forms.Label
    Friend WithEvents btnOn As System.Windows.Forms.Button
    Friend WithEvents btnOff As System.Windows.Forms.Button
    Friend WithEvents btnDoorbell As System.Windows.Forms.Button
    Friend WithEvents btnLastVisitor As System.Windows.Forms.Button

End Class
