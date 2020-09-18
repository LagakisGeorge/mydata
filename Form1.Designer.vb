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
        Me.UPDATE_TIM = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.CancInv = New System.Windows.Forms.Button()
        Me.RequestTransmittedDocs = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(16, 33)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(352, 34)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "send INV to AADE-2"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(16, 246)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(1148, 234)
        Me.TextBox2.TabIndex = 1
        '
        'toXML
        '
        Me.toXML.Location = New System.Drawing.Point(16, 4)
        Me.toXML.Margin = New System.Windows.Forms.Padding(4)
        Me.toXML.Name = "toXML"
        Me.toXML.Size = New System.Drawing.Size(352, 28)
        Me.toXML.TabIndex = 2
        Me.toXML.Text = "to-xml-1"
        Me.toXML.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(16, 124)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(352, 28)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "4.SendIncomeClassification"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'EditConnString
        '
        Me.EditConnString.Location = New System.Drawing.Point(813, 5)
        Me.EditConnString.Margin = New System.Windows.Forms.Padding(4)
        Me.EditConnString.Name = "EditConnString"
        Me.EditConnString.Size = New System.Drawing.Size(352, 28)
        Me.EditConnString.TabIndex = 4
        Me.EditConnString.Text = "opendatabase"
        Me.EditConnString.UseVisualStyleBackColor = True
        '
        'APO
        '
        Me.APO.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.APO.Location = New System.Drawing.Point(555, 4)
        Me.APO.Margin = New System.Windows.Forms.Padding(4)
        Me.APO.Name = "APO"
        Me.APO.Size = New System.Drawing.Size(172, 22)
        Me.APO.TabIndex = 5
        '
        'EOS
        '
        Me.EOS.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.EOS.Location = New System.Drawing.Point(555, 43)
        Me.EOS.Margin = New System.Windows.Forms.Padding(4)
        Me.EOS.Name = "EOS"
        Me.EOS.Size = New System.Drawing.Size(172, 22)
        Me.EOS.TabIndex = 6
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.ItemHeight = 16
        Me.ListBox2.Location = New System.Drawing.Point(376, 74)
        Me.ListBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(788, 164)
        Me.ListBox2.TabIndex = 8
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(813, 41)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(352, 27)
        Me.Button2.TabIndex = 9
        Me.Button2.Text = "check xml"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'UPDATE_TIM
        '
        Me.UPDATE_TIM.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.UPDATE_TIM.Location = New System.Drawing.Point(16, 74)
        Me.UPDATE_TIM.Margin = New System.Windows.Forms.Padding(4)
        Me.UPDATE_TIM.Name = "UPDATE_TIM"
        Me.UPDATE_TIM.Size = New System.Drawing.Size(352, 43)
        Me.UPDATE_TIM.TabIndex = 10
        Me.UPDATE_TIM.Text = "3.UPDATE TIM WITH RESPONSE & DHMIOYRGIA INC.XML"
        Me.UPDATE_TIM.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(16, 489)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 51
        Me.DataGridView1.Size = New System.Drawing.Size(1149, 292)
        Me.DataGridView1.TabIndex = 11
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button4.Location = New System.Drawing.Point(1220, 5)
        Me.Button4.Margin = New System.Windows.Forms.Padding(4)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(215, 28)
        Me.Button4.TabIndex = 12
        Me.Button4.Text = "RequestInvoices"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(16, 164)
        Me.Button5.Margin = New System.Windows.Forms.Padding(4)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(352, 28)
        Me.Button5.TabIndex = 13
        Me.Button5.Text = "5.UpdateDB with IncResponse"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(383, 6)
        Me.Button6.Margin = New System.Windows.Forms.Padding(4)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(156, 60)
        Me.Button6.TabIndex = 14
        Me.Button6.Text = "ολα μαζί"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'CancInv
        '
        Me.CancInv.BackColor = System.Drawing.Color.Lime
        Me.CancInv.Location = New System.Drawing.Point(1220, 103)
        Me.CancInv.Margin = New System.Windows.Forms.Padding(4)
        Me.CancInv.Name = "CancInv"
        Me.CancInv.Size = New System.Drawing.Size(215, 28)
        Me.CancInv.TabIndex = 15
        Me.CancInv.Text = "CancelInvoice"
        Me.CancInv.UseVisualStyleBackColor = False
        '
        'RequestTransmittedDocs
        '
        Me.RequestTransmittedDocs.BackColor = System.Drawing.Color.Lime
        Me.RequestTransmittedDocs.Location = New System.Drawing.Point(1220, 214)
        Me.RequestTransmittedDocs.Margin = New System.Windows.Forms.Padding(4)
        Me.RequestTransmittedDocs.Name = "RequestTransmittedDocs"
        Me.RequestTransmittedDocs.Size = New System.Drawing.Size(215, 57)
        Me.RequestTransmittedDocs.TabIndex = 16
        Me.RequestTransmittedDocs.Text = "RequestTransmittedDocs"
        Me.RequestTransmittedDocs.UseVisualStyleBackColor = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1604, 814)
        Me.Controls.Add(Me.RequestTransmittedDocs)
        Me.Controls.Add(Me.CancInv)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.UPDATE_TIM)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.EOS)
        Me.Controls.Add(Me.APO)
        Me.Controls.Add(Me.EditConnString)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.toXML)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents UPDATE_TIM As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents CancInv As Button
    Friend WithEvents RequestTransmittedDocs As Button
End Class
