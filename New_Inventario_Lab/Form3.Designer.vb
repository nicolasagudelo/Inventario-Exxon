<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Eliminar_Equ_Prod = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Combo_Equ_Prod = New System.Windows.Forms.ComboBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Agregar_Equ_Prod = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Eliminar_Prov_Prod = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Combo_Prov_Prod = New System.Windows.Forms.ComboBox()
        Me.Agregar_Prov_Prod = New System.Windows.Forms.Button()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.Aforo_Ubicacion = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.DataGridView4 = New System.Windows.Forms.DataGridView()
        Me.Eliminar_Ubicacion = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Agregar_Ubicacion = New System.Windows.Forms.Button()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.Mov_ind = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Cantidad_Mov_Reg = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DataGridView5 = New System.Windows.Forms.DataGridView()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DataGridView5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(32, 41)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(4)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1075, 470)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.DarkSlateGray
        Me.TabPage1.Controls.Add(Me.Eliminar_Equ_Prod)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.Combo_Equ_Prod)
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Controls.Add(Me.Agregar_Equ_Prod)
        Me.TabPage1.Location = New System.Drawing.Point(4, 29)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(4)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(4)
        Me.TabPage1.Size = New System.Drawing.Size(1067, 437)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Equipos/Productos"
        '
        'Eliminar_Equ_Prod
        '
        Me.Eliminar_Equ_Prod.Location = New System.Drawing.Point(721, 319)
        Me.Eliminar_Equ_Prod.Margin = New System.Windows.Forms.Padding(4)
        Me.Eliminar_Equ_Prod.Name = "Eliminar_Equ_Prod"
        Me.Eliminar_Equ_Prod.Size = New System.Drawing.Size(173, 49)
        Me.Eliminar_Equ_Prod.TabIndex = 8
        Me.Eliminar_Equ_Prod.Text = "Eliminar"
        Me.Eliminar_Equ_Prod.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label1.Location = New System.Drawing.Point(73, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 20)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Label1"
        '
        'Combo_Equ_Prod
        '
        Me.Combo_Equ_Prod.FormattingEnabled = True
        Me.Combo_Equ_Prod.Location = New System.Drawing.Point(721, 173)
        Me.Combo_Equ_Prod.Name = "Combo_Equ_Prod"
        Me.Combo_Equ_Prod.Size = New System.Drawing.Size(172, 28)
        Me.Combo_Equ_Prod.TabIndex = 2
        '
        'DataGridView1
        '
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.DarkSlateGray
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(77, 104)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(613, 264)
        Me.DataGridView1.TabIndex = 1
        '
        'Agregar_Equ_Prod
        '
        Me.Agregar_Equ_Prod.Location = New System.Drawing.Point(721, 104)
        Me.Agregar_Equ_Prod.Margin = New System.Windows.Forms.Padding(4)
        Me.Agregar_Equ_Prod.Name = "Agregar_Equ_Prod"
        Me.Agregar_Equ_Prod.Size = New System.Drawing.Size(173, 53)
        Me.Agregar_Equ_Prod.TabIndex = 0
        Me.Agregar_Equ_Prod.Text = "Button1"
        Me.Agregar_Equ_Prod.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.DarkSlateGray
        Me.TabPage2.Controls.Add(Me.Eliminar_Prov_Prod)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.Combo_Prov_Prod)
        Me.TabPage2.Controls.Add(Me.Agregar_Prov_Prod)
        Me.TabPage2.Controls.Add(Me.DataGridView2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 29)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(4)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(4)
        Me.TabPage2.Size = New System.Drawing.Size(1067, 437)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Proveedores/Productos"
        '
        'Eliminar_Prov_Prod
        '
        Me.Eliminar_Prov_Prod.Location = New System.Drawing.Point(719, 285)
        Me.Eliminar_Prov_Prod.Margin = New System.Windows.Forms.Padding(4)
        Me.Eliminar_Prov_Prod.Name = "Eliminar_Prov_Prod"
        Me.Eliminar_Prov_Prod.Size = New System.Drawing.Size(173, 49)
        Me.Eliminar_Prov_Prod.TabIndex = 7
        Me.Eliminar_Prov_Prod.Text = "Eliminar"
        Me.Eliminar_Prov_Prod.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label2.Location = New System.Drawing.Point(77, 21)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 20)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Label2"
        '
        'Combo_Prov_Prod
        '
        Me.Combo_Prov_Prod.FormattingEnabled = True
        Me.Combo_Prov_Prod.Location = New System.Drawing.Point(719, 139)
        Me.Combo_Prov_Prod.Name = "Combo_Prov_Prod"
        Me.Combo_Prov_Prod.Size = New System.Drawing.Size(172, 28)
        Me.Combo_Prov_Prod.TabIndex = 4
        '
        'Agregar_Prov_Prod
        '
        Me.Agregar_Prov_Prod.Location = New System.Drawing.Point(719, 70)
        Me.Agregar_Prov_Prod.Margin = New System.Windows.Forms.Padding(4)
        Me.Agregar_Prov_Prod.Name = "Agregar_Prov_Prod"
        Me.Agregar_Prov_Prod.Size = New System.Drawing.Size(173, 49)
        Me.Agregar_Prov_Prod.TabIndex = 3
        Me.Agregar_Prov_Prod.Text = "Button1"
        Me.Agregar_Prov_Prod.UseVisualStyleBackColor = True
        '
        'DataGridView2
        '
        Me.DataGridView2.BackgroundColor = System.Drawing.Color.DarkSlateGray
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(81, 70)
        Me.DataGridView2.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowTemplate.Height = 24
        Me.DataGridView2.Size = New System.Drawing.Size(613, 264)
        Me.DataGridView2.TabIndex = 2
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.DarkSlateGray
        Me.TabPage3.Controls.Add(Me.Aforo_Ubicacion)
        Me.TabPage3.Controls.Add(Me.Label5)
        Me.TabPage3.Controls.Add(Me.Label4)
        Me.TabPage3.Controls.Add(Me.DataGridView4)
        Me.TabPage3.Controls.Add(Me.Eliminar_Ubicacion)
        Me.TabPage3.Controls.Add(Me.Label3)
        Me.TabPage3.Controls.Add(Me.Agregar_Ubicacion)
        Me.TabPage3.Controls.Add(Me.DataGridView3)
        Me.TabPage3.Location = New System.Drawing.Point(4, 29)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(4)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(4)
        Me.TabPage3.Size = New System.Drawing.Size(1067, 437)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Ubicaciones"
        '
        'Aforo_Ubicacion
        '
        Me.Aforo_Ubicacion.Location = New System.Drawing.Point(407, 391)
        Me.Aforo_Ubicacion.Name = "Aforo_Ubicacion"
        Me.Aforo_Ubicacion.Size = New System.Drawing.Size(100, 27)
        Me.Aforo_Ubicacion.TabIndex = 12
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label5.Location = New System.Drawing.Point(347, 394)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(54, 20)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Aforo"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label4.Location = New System.Drawing.Point(95, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(201, 20)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Ubicaciones Existentes"
        '
        'DataGridView4
        '
        Me.DataGridView4.BackgroundColor = System.Drawing.Color.DarkSlateGray
        Me.DataGridView4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView4.Location = New System.Drawing.Point(8, 53)
        Me.DataGridView4.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView4.Name = "DataGridView4"
        Me.DataGridView4.RowTemplate.Height = 24
        Me.DataGridView4.Size = New System.Drawing.Size(516, 301)
        Me.DataGridView4.TabIndex = 9
        '
        'Eliminar_Ubicacion
        '
        Me.Eliminar_Ubicacion.Location = New System.Drawing.Point(543, 380)
        Me.Eliminar_Ubicacion.Margin = New System.Windows.Forms.Padding(4)
        Me.Eliminar_Ubicacion.Name = "Eliminar_Ubicacion"
        Me.Eliminar_Ubicacion.Size = New System.Drawing.Size(149, 49)
        Me.Eliminar_Ubicacion.TabIndex = 8
        Me.Eliminar_Ubicacion.Text = "Eliminar"
        Me.Eliminar_Ubicacion.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label3.Location = New System.Drawing.Point(574, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 20)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Label3"
        '
        'Agregar_Ubicacion
        '
        Me.Agregar_Ubicacion.Location = New System.Drawing.Point(156, 380)
        Me.Agregar_Ubicacion.Margin = New System.Windows.Forms.Padding(4)
        Me.Agregar_Ubicacion.Name = "Agregar_Ubicacion"
        Me.Agregar_Ubicacion.Size = New System.Drawing.Size(149, 49)
        Me.Agregar_Ubicacion.TabIndex = 4
        Me.Agregar_Ubicacion.Text = "Agregar"
        Me.Agregar_Ubicacion.UseVisualStyleBackColor = True
        '
        'DataGridView3
        '
        Me.DataGridView3.BackgroundColor = System.Drawing.Color.DarkSlateGray
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(543, 53)
        Me.DataGridView3.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.RowTemplate.Height = 24
        Me.DataGridView3.Size = New System.Drawing.Size(516, 301)
        Me.DataGridView3.TabIndex = 3
        '
        'TabPage4
        '
        Me.TabPage4.BackColor = System.Drawing.Color.DarkSlateGray
        Me.TabPage4.Controls.Add(Me.Mov_ind)
        Me.TabPage4.Controls.Add(Me.Label7)
        Me.TabPage4.Controls.Add(Me.Cantidad_Mov_Reg)
        Me.TabPage4.Controls.Add(Me.Label6)
        Me.TabPage4.Controls.Add(Me.DataGridView5)
        Me.TabPage4.Location = New System.Drawing.Point(4, 29)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(1067, 437)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Movimiento"
        '
        'Mov_ind
        '
        Me.Mov_ind.Location = New System.Drawing.Point(862, 190)
        Me.Mov_ind.Name = "Mov_ind"
        Me.Mov_ind.Size = New System.Drawing.Size(92, 27)
        Me.Mov_ind.TabIndex = 11
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial Rounded MT Bold", 13.8!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(824, 142)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(141, 28)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Pendientes"
        '
        'Cantidad_Mov_Reg
        '
        Me.Cantidad_Mov_Reg.Location = New System.Drawing.Point(829, 235)
        Me.Cantidad_Mov_Reg.Margin = New System.Windows.Forms.Padding(4)
        Me.Cantidad_Mov_Reg.Name = "Cantidad_Mov_Reg"
        Me.Cantidad_Mov_Reg.Size = New System.Drawing.Size(149, 49)
        Me.Cantidad_Mov_Reg.TabIndex = 9
        Me.Cantidad_Mov_Reg.Text = "Ingresar / Sacar"
        Me.Cantidad_Mov_Reg.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label6.Location = New System.Drawing.Point(41, 37)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(525, 20)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "El movimiento solo se finaliza cuando los pendientes esten en 0"
        '
        'DataGridView5
        '
        Me.DataGridView5.BackgroundColor = System.Drawing.Color.DarkSlateGray
        Me.DataGridView5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView5.Location = New System.Drawing.Point(45, 80)
        Me.DataGridView5.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView5.Name = "DataGridView5"
        Me.DataGridView5.RowTemplate.Height = 24
        Me.DataGridView5.Size = New System.Drawing.Size(732, 319)
        Me.DataGridView5.TabIndex = 7
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.SlateGray
        Me.ClientSize = New System.Drawing.Size(1147, 542)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Arial Rounded MT Bold", 10.2!)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Form3"
        Me.Text = "Relaciones"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.DataGridView4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        CType(Me.DataGridView5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents Agregar_Equ_Prod As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents DataGridView3 As DataGridView
    Friend WithEvents Combo_Equ_Prod As ComboBox
    Friend WithEvents Combo_Prov_Prod As ComboBox
    Friend WithEvents Agregar_Prov_Prod As Button
    Friend WithEvents Agregar_Ubicacion As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Eliminar_Equ_Prod As Button
    Friend WithEvents Eliminar_Prov_Prod As Button
    Friend WithEvents Eliminar_Ubicacion As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents DataGridView4 As DataGridView
    Friend WithEvents Aforo_Ubicacion As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents Label6 As Label
    Friend WithEvents DataGridView5 As DataGridView
    Friend WithEvents Mov_ind As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Cantidad_Mov_Reg As Button
End Class
