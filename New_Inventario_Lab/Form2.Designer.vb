<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormIngreso
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormIngreso))
        Me.TxtBxUsuario = New System.Windows.Forms.TextBox()
        Me.TxtBxContraseña = New System.Windows.Forms.TextBox()
        Me.BtnAceptar = New System.Windows.Forms.Button()
        Me.BtnCancelar = New System.Windows.Forms.Button()
        Me.LblUsuario = New System.Windows.Forms.Label()
        Me.LblContraseña = New System.Windows.Forms.Label()
        Me.GBValidacionUsuario = New System.Windows.Forms.GroupBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GBValidacionUsuario.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TxtBxUsuario
        '
        Me.TxtBxUsuario.Location = New System.Drawing.Point(279, 42)
        Me.TxtBxUsuario.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TxtBxUsuario.Name = "TxtBxUsuario"
        Me.TxtBxUsuario.Size = New System.Drawing.Size(168, 23)
        Me.TxtBxUsuario.TabIndex = 0
        '
        'TxtBxContraseña
        '
        Me.TxtBxContraseña.Location = New System.Drawing.Point(279, 109)
        Me.TxtBxContraseña.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TxtBxContraseña.Name = "TxtBxContraseña"
        Me.TxtBxContraseña.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TxtBxContraseña.Size = New System.Drawing.Size(168, 23)
        Me.TxtBxContraseña.TabIndex = 1
        '
        'BtnAceptar
        '
        Me.BtnAceptar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnAceptar.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.BtnAceptar.Location = New System.Drawing.Point(44, 179)
        Me.BtnAceptar.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.BtnAceptar.Name = "BtnAceptar"
        Me.BtnAceptar.Size = New System.Drawing.Size(169, 41)
        Me.BtnAceptar.TabIndex = 2
        Me.BtnAceptar.Text = "Aceptar"
        Me.BtnAceptar.UseVisualStyleBackColor = True
        '
        'BtnCancelar
        '
        Me.BtnCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancelar.ForeColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.BtnCancelar.Location = New System.Drawing.Point(279, 179)
        Me.BtnCancelar.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.BtnCancelar.Name = "BtnCancelar"
        Me.BtnCancelar.Size = New System.Drawing.Size(169, 41)
        Me.BtnCancelar.TabIndex = 3
        Me.BtnCancelar.Text = "Cancelar"
        Me.BtnCancelar.UseVisualStyleBackColor = True
        '
        'LblUsuario
        '
        Me.LblUsuario.AutoSize = True
        Me.LblUsuario.Location = New System.Drawing.Point(62, 47)
        Me.LblUsuario.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblUsuario.Name = "LblUsuario"
        Me.LblUsuario.Size = New System.Drawing.Size(61, 16)
        Me.LblUsuario.TabIndex = 4
        Me.LblUsuario.Text = "Usuario"
        '
        'LblContraseña
        '
        Me.LblContraseña.AutoSize = True
        Me.LblContraseña.Location = New System.Drawing.Point(62, 116)
        Me.LblContraseña.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblContraseña.Name = "LblContraseña"
        Me.LblContraseña.Size = New System.Drawing.Size(85, 16)
        Me.LblContraseña.TabIndex = 5
        Me.LblContraseña.Text = "Contraseña"
        '
        'GBValidacionUsuario
        '
        Me.GBValidacionUsuario.Controls.Add(Me.TxtBxUsuario)
        Me.GBValidacionUsuario.Controls.Add(Me.LblContraseña)
        Me.GBValidacionUsuario.Controls.Add(Me.TxtBxContraseña)
        Me.GBValidacionUsuario.Controls.Add(Me.LblUsuario)
        Me.GBValidacionUsuario.Controls.Add(Me.BtnAceptar)
        Me.GBValidacionUsuario.Controls.Add(Me.BtnCancelar)
        Me.GBValidacionUsuario.ForeColor = System.Drawing.SystemColors.InactiveCaption
        Me.GBValidacionUsuario.Location = New System.Drawing.Point(64, 77)
        Me.GBValidacionUsuario.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GBValidacionUsuario.Name = "GBValidacionUsuario"
        Me.GBValidacionUsuario.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.GBValidacionUsuario.Size = New System.Drawing.Size(534, 253)
        Me.GBValidacionUsuario.TabIndex = 6
        Me.GBValidacionUsuario.TabStop = False
        Me.GBValidacionUsuario.Text = "Validacion de Usuario"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Inventario_Lab.My.Resources.Resources.Admin
        Me.PictureBox1.Location = New System.Drawing.Point(628, 77)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(244, 253)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 7
        Me.PictureBox1.TabStop = False
        '
        'FormIngreso
        '
        Me.AcceptButton = Me.BtnAceptar
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DarkSlateGray
        Me.CancelButton = Me.BtnCancelar
        Me.ClientSize = New System.Drawing.Size(928, 401)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.GBValidacionUsuario)
        Me.Font = New System.Drawing.Font("Arial Rounded MT Bold", 10.2!)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Name = "FormIngreso"
        Me.Text = "Control de Acceso"
        Me.GBValidacionUsuario.ResumeLayout(False)
        Me.GBValidacionUsuario.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TxtBxUsuario As TextBox
    Friend WithEvents TxtBxContraseña As TextBox
    Friend WithEvents BtnAceptar As Button
    Friend WithEvents BtnCancelar As Button
    Friend WithEvents LblUsuario As Label
    Friend WithEvents LblContraseña As Label
    Friend WithEvents GBValidacionUsuario As GroupBox
    Friend WithEvents PictureBox1 As PictureBox
End Class
