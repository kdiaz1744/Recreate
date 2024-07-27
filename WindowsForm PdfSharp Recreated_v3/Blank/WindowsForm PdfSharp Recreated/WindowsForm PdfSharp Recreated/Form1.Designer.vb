<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnGenerateCertificate = New System.Windows.Forms.Button()
        Me.PO_Number = New System.Windows.Forms.TextBox()
        Me.Data_Recieved = New System.Windows.Forms.TextBox()
        Me.Supplier = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnGenerateCertificate
        '
        Me.btnGenerateCertificate.Location = New System.Drawing.Point(104, 12)
        Me.btnGenerateCertificate.Name = "btnGenerateCertificate"
        Me.btnGenerateCertificate.Size = New System.Drawing.Size(75, 23)
        Me.btnGenerateCertificate.TabIndex = 0
        Me.btnGenerateCertificate.Text = "Recreate"
        Me.btnGenerateCertificate.UseVisualStyleBackColor = True
        '
        'PO_Number
        '
        Me.PO_Number.Location = New System.Drawing.Point(147, 92)
        Me.PO_Number.Name = "PO_Number"
        Me.PO_Number.Size = New System.Drawing.Size(100, 20)
        Me.PO_Number.TabIndex = 5
        '
        'Data_Recieved
        '
        Me.Data_Recieved.Location = New System.Drawing.Point(148, 67)
        Me.Data_Recieved.Name = "Data_Recieved"
        Me.Data_Recieved.Size = New System.Drawing.Size(100, 20)
        Me.Data_Recieved.TabIndex = 4
        '
        'Supplier
        '
        Me.Supplier.Location = New System.Drawing.Point(149, 41)
        Me.Supplier.Name = "Supplier"
        Me.Supplier.Size = New System.Drawing.Size(100, 20)
        Me.Supplier.TabIndex = 0
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(43, 93)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(100, 20)
        Me.TextBox3.TabIndex = 3
        Me.TextBox3.Text = "PO #"
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(43, 67)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 2
        Me.TextBox1.Text = "Data Recieved"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(43, 41)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(100, 20)
        Me.TextBox2.TabIndex = 1
        Me.TextBox2.Text = "Supplier"
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.PO_Number)
        Me.Controls.Add(Me.Data_Recieved)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Supplier)
        Me.Controls.Add(Me.btnGenerateCertificate)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnGenerateCertificate As Button
    Friend WithEvents PO_Number As TextBox
    Friend WithEvents Data_Recieved As TextBox
    Friend WithEvents Supplier As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
End Class
