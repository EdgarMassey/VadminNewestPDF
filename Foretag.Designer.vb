<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Foretag
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Ver = New System.Windows.Forms.GroupBox()
        Me.datum = New System.Windows.Forms.Label()
        Me.Versionl = New System.Windows.Forms.Label()
        Me.Huvud = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.KlientIdTB = New System.Windows.Forms.TextBox()
        Me.LosenTB = New System.Windows.Forms.TextBox()
        Me.LoginK = New System.Windows.Forms.Button()
        Me.AvslutaK = New System.Windows.Forms.Button()
        Me.SparaCB = New System.Windows.Forms.CheckBox()
        Me.Ver.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(138, 101)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "KlientId"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Ver
        '
        Me.Ver.Controls.Add(Me.datum)
        Me.Ver.Controls.Add(Me.Versionl)
        Me.Ver.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Ver.ForeColor = System.Drawing.Color.Gray
        Me.Ver.Location = New System.Drawing.Point(11, 263)
        Me.Ver.Name = "Ver"
        Me.Ver.Size = New System.Drawing.Size(197, 34)
        Me.Ver.TabIndex = 1
        Me.Ver.TabStop = False
        Me.Ver.Text = "Version"
        '
        'datum
        '
        Me.datum.Location = New System.Drawing.Point(45, 16)
        Me.datum.Name = "datum"
        Me.datum.Size = New System.Drawing.Size(115, 15)
        Me.datum.TabIndex = 1
        Me.datum.Text = "Datum"
        Me.datum.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Versionl
        '
        Me.Versionl.AutoSize = True
        Me.Versionl.Location = New System.Drawing.Point(22, 22)
        Me.Versionl.Name = "Versionl"
        Me.Versionl.Size = New System.Drawing.Size(0, 15)
        Me.Versionl.TabIndex = 0
        '
        'Huvud
        '
        Me.Huvud.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Huvud.ForeColor = System.Drawing.Color.Maroon
        Me.Huvud.Location = New System.Drawing.Point(36, 12)
        Me.Huvud.Name = "Huvud"
        Me.Huvud.Size = New System.Drawing.Size(406, 50)
        Me.Huvud.TabIndex = 10
        Me.Huvud.Text = "Behörighetskontrol"
        Me.Huvud.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label2.Location = New System.Drawing.Point(147, 142)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 21)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Lösen"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'KlientIdTB
        '
        Me.KlientIdTB.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KlientIdTB.Location = New System.Drawing.Point(216, 101)
        Me.KlientIdTB.Name = "KlientIdTB"
        Me.KlientIdTB.Size = New System.Drawing.Size(121, 22)
        Me.KlientIdTB.TabIndex = 1
        '
        'LosenTB
        '
        Me.LosenTB.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LosenTB.Location = New System.Drawing.Point(216, 142)
        Me.LosenTB.Name = "LosenTB"
        Me.LosenTB.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.LosenTB.Size = New System.Drawing.Size(121, 22)
        Me.LosenTB.TabIndex = 2
        '
        'LoginK
        '
        Me.LoginK.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LoginK.ForeColor = System.Drawing.Color.Maroon
        Me.LoginK.Location = New System.Drawing.Point(359, 239)
        Me.LoginK.Name = "LoginK"
        Me.LoginK.Size = New System.Drawing.Size(98, 24)
        Me.LoginK.TabIndex = 6
        Me.LoginK.Text = "Login"
        Me.LoginK.UseVisualStyleBackColor = True
        '
        'AvslutaK
        '
        Me.AvslutaK.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AvslutaK.ForeColor = System.Drawing.Color.Maroon
        Me.AvslutaK.Location = New System.Drawing.Point(359, 282)
        Me.AvslutaK.Name = "AvslutaK"
        Me.AvslutaK.Size = New System.Drawing.Size(98, 25)
        Me.AvslutaK.TabIndex = 7
        Me.AvslutaK.Text = "Avsluta"
        Me.AvslutaK.UseVisualStyleBackColor = True
        '
        'SparaCB
        '
        Me.SparaCB.AutoSize = True
        Me.SparaCB.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SparaCB.Location = New System.Drawing.Point(216, 186)
        Me.SparaCB.Name = "SparaCB"
        Me.SparaCB.Size = New System.Drawing.Size(236, 19)
        Me.SparaCB.TabIndex = 8
        Me.SparaCB.Text = "Spara inloggnings information i datorn"
        Me.SparaCB.UseVisualStyleBackColor = True
        '
        'Foretag
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(489, 319)
        Me.Controls.Add(Me.SparaCB)
        Me.Controls.Add(Me.AvslutaK)
        Me.Controls.Add(Me.LoginK)
        Me.Controls.Add(Me.LosenTB)
        Me.Controls.Add(Me.KlientIdTB)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Huvud)
        Me.Controls.Add(Me.Ver)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Times New Roman", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Foretag"
        Me.Text = "Foretag"
        Me.Ver.ResumeLayout(False)
        Me.Ver.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Ver As System.Windows.Forms.GroupBox
    Friend WithEvents Versionl As System.Windows.Forms.Label
    Friend WithEvents Huvud As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents KlientIdTB As System.Windows.Forms.TextBox
    Friend WithEvents LosenTB As System.Windows.Forms.TextBox
    Friend WithEvents LoginK As System.Windows.Forms.Button
    Friend WithEvents AvslutaK As System.Windows.Forms.Button
    Friend WithEvents SparaCB As System.Windows.Forms.CheckBox
    Friend WithEvents datum As System.Windows.Forms.Label

End Class
