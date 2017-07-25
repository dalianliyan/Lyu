<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.txtSource = New System.Windows.Forms.TextBox()
        Me.txtResult = New System.Windows.Forms.TextBox()
        Me.btnApp = New System.Windows.Forms.Button()
        Me.txtXPath = New System.Windows.Forms.TextBox()
        Me.btnOpen = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtSource
        '
        Me.txtSource.Location = New System.Drawing.Point(12, 41)
        Me.txtSource.Multiline = True
        Me.txtSource.Name = "txtSource"
        Me.txtSource.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtSource.Size = New System.Drawing.Size(488, 670)
        Me.txtSource.TabIndex = 0
        Me.txtSource.Text = resources.GetString("txtSource.Text")
        '
        'txtResult
        '
        Me.txtResult.Location = New System.Drawing.Point(515, 41)
        Me.txtResult.Multiline = True
        Me.txtResult.Name = "txtResult"
        Me.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtResult.Size = New System.Drawing.Size(488, 670)
        Me.txtResult.TabIndex = 0
        '
        'btnApp
        '
        Me.btnApp.Location = New System.Drawing.Point(928, 9)
        Me.btnApp.Name = "btnApp"
        Me.btnApp.Size = New System.Drawing.Size(75, 23)
        Me.btnApp.TabIndex = 1
        Me.btnApp.Text = "Go"
        Me.btnApp.UseVisualStyleBackColor = True
        '
        'txtXPath
        '
        Me.txtXPath.Location = New System.Drawing.Point(97, 12)
        Me.txtXPath.Name = "txtXPath"
        Me.txtXPath.Size = New System.Drawing.Size(825, 20)
        Me.txtXPath.TabIndex = 2
        '
        'btnOpen
        '
        Me.btnOpen.Location = New System.Drawing.Point(12, 10)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(75, 23)
        Me.btnOpen.TabIndex = 3
        Me.btnOpen.Text = "Open"
        Me.btnOpen.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AcceptButton = Me.btnApp
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1027, 721)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.txtXPath)
        Me.Controls.Add(Me.btnApp)
        Me.Controls.Add(Me.txtResult)
        Me.Controls.Add(Me.txtSource)
        Me.Name = "Form1"
        Me.Text = "XPath Viewer"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtSource As System.Windows.Forms.TextBox
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    Friend WithEvents btnApp As System.Windows.Forms.Button
    Friend WithEvents txtXPath As System.Windows.Forms.TextBox
    Friend WithEvents btnOpen As System.Windows.Forms.Button

End Class
