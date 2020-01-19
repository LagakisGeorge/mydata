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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.toXML = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.EditConnString = New System.Windows.Forms.Button()
        Me.APO = New System.Windows.Forms.DateTimePicker()
        Me.EOS = New System.Windows.Forms.DateTimePicker()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(625, 14)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(180, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "send INV to AADE"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(12, 225)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(764, 191)
        Me.TextBox2.TabIndex = 1
        '
        'toXML
        '
        Me.toXML.Location = New System.Drawing.Point(12, 3)
        Me.toXML.Name = "toXML"
        Me.toXML.Size = New System.Drawing.Size(264, 23)
        Me.toXML.TabIndex = 2
        Me.toXML.Text = "to-xml"
        Me.toXML.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(12, 32)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(264, 23)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "SendIncomeClassification"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'EditConnString
        '
        Me.EditConnString.Location = New System.Drawing.Point(12, 61)
        Me.EditConnString.Name = "EditConnString"
        Me.EditConnString.Size = New System.Drawing.Size(264, 23)
        Me.EditConnString.TabIndex = 4
        Me.EditConnString.Text = "opendatabase"
        Me.EditConnString.UseVisualStyleBackColor = True
        '
        'APO
        '
        Me.APO.Location = New System.Drawing.Point(394, 3)
        Me.APO.Name = "APO"
        Me.APO.Size = New System.Drawing.Size(130, 20)
        Me.APO.TabIndex = 5
        '
        'EOS
        '
        Me.EOS.Location = New System.Drawing.Point(394, 35)
        Me.EOS.Name = "EOS"
        Me.EOS.Size = New System.Drawing.Size(130, 20)
        Me.EOS.TabIndex = 6
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.Location = New System.Drawing.Point(0, 90)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(1180, 134)
        Me.ListBox2.TabIndex = 8
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(899, 14)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(120, 40)
        Me.Button2.TabIndex = 9
        Me.Button2.Text = "check xml"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1225, 450)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.EOS)
        Me.Controls.Add(Me.APO)
        Me.Controls.Add(Me.EditConnString)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.toXML)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents toXML As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents EditConnString As Button
    Friend WithEvents APO As DateTimePicker
    Friend WithEvents EOS As DateTimePicker
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents Button2 As Button
End Class
