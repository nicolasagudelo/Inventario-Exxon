Imports MySql.Data.MySqlClient
Imports System.Configuration

Public Class Form3
    Dim conn As New MySqlConnection
    Private Sub Connect()
        conn.ConnectionString = ConfigurationManager.ConnectionStrings("cs").ConnectionString
        Try
            conn.Open()
            Console.WriteLine("conectandose a la base de datos")
        Catch ex As Exception
            MsgBox(ex.Message)
            End
        End Try
        conn.Close()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Connect()
        Cargar()
    End Sub

    Private Sub Consulta(ByVal Query As String)
        Try
            conn.Open()
            Dim Comando As New MySqlCommand(Query, conn)
            Dim Adaptador As New MySqlDataAdapter(Comando)
            Dim Tabla As New DataTable
            Adaptador.Fill(Tabla)
                Select Case Bandera_Rel
                    Case 1
                        DataGridView1.DataSource = Tabla
                        DataGridView1.ReadOnly = True
                        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
                        DataGridView1.AutoResizeColumns()
                        DataGridView1.Columns(2).Visible = False
                    Case 2
                        DataGridView2.DataSource = Tabla
                        DataGridView2.ReadOnly = True
                        DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
                        DataGridView2.AutoResizeColumns()
                        DataGridView2.Columns(2).Visible = False
                    Case 3
                        DataGridView3.DataSource = Tabla
                        DataGridView3.ReadOnly = True
                        DataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect
                        DataGridView3.AutoResizeColumns()
                    Case 4
                        DataGridView5.DataSource = Tabla
                        DataGridView5.ReadOnly = True
                        DataGridView5.SelectionMode = DataGridViewSelectionMode.FullRowSelect
                        DataGridView5.AutoResizeColumns()
                End Select
            conn.Close()
        Catch ex As Exception
            MsgBox("Error." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
            conn.Close()
        End Try
    End Sub

    Dim Id1 As String
    Dim Id2 As String

    Private Sub Cargar()
        Id1 = ""
        Id2 = ""
        For i = 1 To 4
            If Me.Controls.Find("TabPage" & i, True).Count = 1 Then
                Dim b As TabPage = Me.Controls.Find("TabPage" & i, True)(0)
                b.Parent = Nothing
            End If
        Next
        Select Case Bandera_Rel
            Case 1
                TabPage1.Parent = TabControl1
                Consulta(Consulta_rel)
                Agregar_Equ_Prod.Text = "Agregar " & Tabla_Rel
                Eliminar_Equ_Prod.Text = "Eliminar " & Tabla_Rel
                With Combo_Equ_Prod
                    Try
                        conn.Open()
                        Dim consulta As String = "Select * from " & Tabla_Rel & " Where Activo='1'"
                        Dim cmd As New MySqlCommand(consulta, conn)
                        Dim MysqlDadap As New MySqlDataAdapter(cmd)
                        Dim MysqlDset As New DataSet
                        MysqlDadap.Fill(MysqlDset)
                        conn.Close()
                        .DataSource = MysqlDset.Tables(0)
                        If Tabla_Rel = "Equipos" Then
                            Label1.Text = "Gestor para agregar Equipos al Producto: " & Elemento_rel
                            Id1 = Id_elem
                            .DisplayMember = "Nombre_Equipo"
                            .ValueMember = "Id_Equipo"
                        ElseIf Tabla_Rel = "Productos" Then
                            Label1.Text = "Gestor para agregar Productos al Equipo: " & Elemento_rel
                            Id2 = Id_elem
                            .DisplayMember = "Nombre_Producto"
                            .ValueMember = "Id_Producto"
                        End If
                        .SelectedValue = -1
                        '.Enabled = False
                    Catch ex As Exception
                        MsgBox("Error." & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
                        conn.Close()
                    End Try
                End With
                'Case 2
                '    TabPage2.Parent = TabControl1
                '    Consulta(Consulta_rel)
                '    Agregar_Prov_Prod.Text = "Agregar " & Tabla_Rel
                '    Eliminar_Prov_Prod.Text = "Eliminar " & Tabla_Rel
                '    With Combo_Prov_Prod
                '        Try
                '            Dim conexion As New MySqlConnection(datasource)
                '            conexion.Open()
                '            Dim consulta As String = "Select * from " & Tabla_Rel & " Where Activo='1'"
                '            Dim MysqlDadap As New MySqlDataAdapter(consulta, conexion)
                '            Dim MysqlDset As New DataSet
                '            MysqlDadap.Fill(MysqlDset)
                '            conexion.Close()
                '            .DataSource = MysqlDset.Tables(0)
                '            If Tabla_Rel = "Proveedores" Then
                '                Id1 = Id_elem
                '                Label2.Text = "Gestor para agregar Proveedores al Producto: " & Elemento_rel
                '                .DisplayMember = "Nombre_Proveedor"
                '                .ValueMember = "Nit_Proveedor"
                '            ElseIf Tabla_Rel = "Productos" Then
                '                Id2 = Id_elem
                '                Label2.Text = "Gestor para agregar Productos al Proveedor: " & Elemento_rel
                '                .DisplayMember = "Nombre_Producto"
                '                .ValueMember = "Id_Producto"
                '            End If
                '            .SelectedValue = -1
                '            '.Enabled = False
                '        Catch ex As Exception
                '            MessageBox.Show(ex.Message)
                '        End Try
                '    End With
                'Case 3
                '    Id1 = Id_elem
                '    TabPage3.Parent = TabControl1
                '    Consulta(Consulta_rel)
                '    Label3.Text = "Ubicaciones asignadas al producto: " & Elemento_rel
                '    Agregar_Ubicacion.Text = "Agregar " & Tabla_Rel
                '    Eliminar_Ubicacion.Text = "Eliminar " & Tabla_Rel
                '    Using conexion1 As New MySqlConnection(datasource)
                '        Dim Comando1 As New MySqlCommand("select * from ubicaciones", conexion1)
                '        Dim Adaptador1 As New MySqlDataAdapter(Comando1)
                '        Dim Tabla1 As New DataTable
                '        DataGridView3.Columns(6).Visible = False
                '        Adaptador1.Fill(Tabla1)
                '        DataGridView4.DataSource = Tabla1
                '        DataGridView4.ReadOnly = True
                '        DataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect
                '        DataGridView4.AutoResizeColumns()
                '        DataGridView4.Columns(0).Visible = False
                '        conexion1.Close()
                '    End Using
                'Case 4
                '    If Cantidad_a_mover = 0 Then
                '        Label7.ForeColor = Color.Green
                '        Cantidad_Mov_Reg.Enabled = False
                '    Else
                '        Label7.ForeColor = Color.Red
                '        Cantidad_Mov_Reg.Enabled = True
                '    End If
                '    TabPage4.Parent = TabControl1
                '    Consulta(Consulta_rel)
                '    Label7.Text = "Pendientes: " & Cantidad_a_mover
                '    DataGridView5.Columns("IdRel_Ubicaciones_Productos").Visible = False
        End Select
    End Sub

    Private Sub Agregar_Equ_Prod_Click(sender As Object, e As EventArgs) Handles Agregar_Equ_Prod.Click
        If Combo_Equ_Prod.SelectedValue = Nothing Then
            Exit Sub
        End If
        If Id1 <> "" Then
                Id2 = Combo_Equ_Prod.SelectedValue
            Else
                Id1 = Combo_Equ_Prod.SelectedValue
            End If
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("INSERT INTO rel_productos_equipos (Id_Producto, Id_equipo)
                            VALUES (@IdProducto, @IdEquipo)", conn)
                With cmd.Parameters
                    .AddWithValue("IdProducto", Id1)
                    .AddWithValue("IdEquipo", Id2)
                End With
                cmd.ExecuteNonQuery()
                'MessageBox.Show("Registro MODIFICADO")
                conn.Close()
            Catch ex As Exception
                MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try
        Cargar()
        For Each row As DataGridViewRow In DataGridView1.Rows
            For Each cell As DataGridViewCell In row.Cells
                If (Not cell.Value Is Nothing) AndAlso cell.Value.ToString() = Combo_Equ_Prod.Text Then
                    row.Selected = True
                    Exit Sub
                End If
            Next
        Next
    End Sub

    Private Sub Eliminar_Equ_Prod_Click(sender As Object, e As EventArgs) Handles Eliminar_Equ_Prod.Click
        Dim result As Integer = MessageBox.Show("Esta seguro que desea ELIMINAR el registro", "Alerta", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        ElseIf result = DialogResult.Yes Then
            Dim a As Integer = DataGridView1.CurrentRow.Index
            Try
                conn.Open()
                Dim query As String = "DELETE FROM rel_productos_equipos WHERE " & DataGridView1.Columns(2).Name & " = '" & DataGridView1.Item(2, a).Value & "'"
                Dim cmd As New MySqlCommand(query, conn)
                cmd.ExecuteNonQuery()
                MessageBox.Show("Registro ELIMINADO")
                conn.Close()
            Catch ex As Exception
                MsgBox("El registro no pudo Eliminarse por: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
            End Try
        End If
        Cargar()
    End Sub
End Class