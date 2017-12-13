Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Data
Imports System.Windows.Forms
Imports System.Configuration
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions

Public Class Form1
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
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Connect()
        TabControl1.Visible = False
        TabControl1.BackColor = Color.Transparent
        TabControl2.Visible = False
        TabControl2.BackColor = Color.Transparent
        Foto_Usuario.AllowDrop = True
        Foto_Equipo.AllowDrop = True
        Foto_Producto.AllowDrop = True
        Cargar_CBTabUsuarios()
    End Sub

    Private Sub Cargar_CBTabUsuarios()
        With Perfiles_Usuario
            Try
                conn.Open()
                Dim query As String = "Select ID_Perfil, Nombre_Perfil from perfiles"
                Dim cmd As New MySqlCommand(query, conn)
                Dim sqlAdap As New MySqlDataAdapter(cmd)
                Dim dtRecord As New DataTable
                sqlAdap.Fill(dtRecord)
                .DataSource = dtRecord
                .DisplayMember = "Nombre_Perfil"
                .ValueMember = "ID_Perfil"
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar los perfiles de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End With
        With Doag_Usuarios
            Try
                conn.Open()
                Dim query As String = "Select ID_Doag, Nombre_Doag from doag"
                Dim cmd As New MySqlCommand(query, conn)
                Dim sqlAdap As New MySqlDataAdapter(cmd)
                Dim dtRecord As New DataTable
                sqlAdap.Fill(dtRecord)
                .DataSource = dtRecord
                .DisplayMember = "Nombre_Doag"
                .ValueMember = "ID_Doag"
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar los Doag de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End With
    End Sub

    Private Sub Esconder_tabpages()
        TabControl1.Visible = True
        TabControl2.Visible = False
        For i = 1 To 6
            If Me.Controls.Find("TabPage" & i, True).Count = 1 Then
                Dim b As TabPage = Me.Controls.Find("TabPage" & i, True)(0)
                b.Parent = Nothing
            End If
        Next
    End Sub
    Private Sub Esconder_tabpages_submenu()
        TabControl2.Visible = True
        For i = 6 To 16
            If Me.Controls.Find("TabPage" & i, True).Count = 1 Then
                Dim b As TabPage = Me.Controls.Find("TabPage" & i, True)(0)
                b.Parent = Nothing
            End If
        Next
    End Sub
    Private Sub Menu_Seleccionado(ByVal Bandera_Menu As Integer)
        TabControl2.Visible = False
        Select Case Bandera_Menu
            Case 1
                Esconder_tabpages()
                TabPage1.Parent = TabControl1 'Administrar
            Case 2
                Esconder_tabpages()
                TabPage2.Parent = TabControl1 'Movimientos
            Case 3
                Esconder_tabpages()
                TabPage3.Parent = TabControl1 'Equipos
            Case 4
                Esconder_tabpages()
                TabPage4.Parent = TabControl1 'Productos
            Case 5
                Esconder_tabpages()
                TabPage5.Parent = TabControl1 'Proveedores
            Case 6
                Esconder_tabpages()
                TabPage6.Parent = TabControl1 'Reportes
        End Select
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Menu_Seleccionado(1)
        cant_reg_encon = 0
        z = "USUARIOS"
    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Menu_Seleccionado(2)
        Esconder_tabpages_submenu()
        TabPage12.Parent = TabControl2
    End Sub

    Dim reg_bus_equ As String
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Menu_Seleccionado(3)
        Esconder_tabpages_submenu()
        TabPage13.Parent = TabControl2
        'cant_reg_encon = 0
        'z = "EQUIPOS"
        'Recorrer_Equipos()
    End Sub

    Dim reg_bus_produ As String
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Menu_Seleccionado(4)
        Esconder_tabpages_submenu()
        TabPage14.Parent = TabControl2
        'cant_reg_encon = 0
        'z = "PRODUCTOS"
        'Recorrer_Productos()
    End Sub
    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Menu_Seleccionado(5)
        Esconder_tabpages_submenu()
        'TabPage15.Parent = TabControl2
        'cant_reg_encon = 0
        'z = "PROVEEDORES"
        'Recorrer_Proveedores()
    End Sub
    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        Menu_Seleccionado(6)
        Esconder_tabpages_submenu()
        TabPage16.Parent = TabControl2
    End Sub


    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        FormIngreso.Close()
    End Sub

    Private Sub Gestion_Usuario_Click(sender As Object, e As EventArgs) Handles Gestion_Usuario.Click
        Modificar_Usuario = 0
        Agregar_Usuario = 0
        HabilitarControlesUsuario()
        ActivarCamposContrasena()
        Esconder_tabpages_submenu()
        TabPage7.Parent = TabControl2 'Usuarios
        TabPage8.Parent = TabControl2 'Monto
        CargarDGVMontos()
        TabPage9.Parent = TabControl2 'Perfiles
        CargarPerfiles()
        Recorrer_Usuarios()
    End Sub


    Private Sub Gestion_Almacen_Click(sender As Object, e As EventArgs) Handles Gestion_Almacen.Click
        Esconder_tabpages_submenu()
        TabPage10.Parent = TabControl2
        CargarCBTabCategorias()
    End Sub

    Dim Id_Ubicacion As Integer
    Dim Categ As String
    Dim CategLoad As Boolean = False
    Private Sub CargarCBTabCategorias()
        With Nombre_Categoria
            Try
                conn.Open()
                Dim query As String = "Select Id_Categoria, Nombre_Categoria from categorias"
                Dim cmd As New MySqlCommand(query, conn)
                Dim sqlAdap As New MySqlDataAdapter(cmd)
                Dim dtRecord As New DataTable
                sqlAdap.Fill(dtRecord)
                .DataSource = dtRecord
                .DisplayMember = "Nombre_Categoria"
                .ValueMember = "Id_Categoria"
                .SelectedValue = dtRecord.Rows(0).Item(0)
                Categ = .SelectedValue
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar las categorias de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End With
        CategLoad = True
        CargarSubCategorias()
        TabPage11.Parent = TabControl2
        Cargar_Tabla("*", "Ubicaciones")
        Id_Ubicacion = Tabla1.Rows(0).ItemArray(0)
        Estantes.Text = Tabla1.Rows(0).ItemArray(1).ToString
        Entrepanos.Text = Tabla1.Rows(0).ItemArray(2).ToString
        Cajas_Colores.Text = Tabla1.Rows(0).ItemArray(3).ToString
        Zonas.Text = Tabla1.Rows(0).ItemArray(4).ToString
        DataGridView3.DataSource = Nothing
        DataGridView3.DataSource = Tabla1
        DataGridView3.ReadOnly = True
        DataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView3.Columns(0).Visible = False
    End Sub

    Private Sub CargarSubCategorias()
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(String.Format("Select * from categorias_sub where Id_Categoria = @IDCat;"), conn)
            Categ = Nombre_Categoria.SelectedValue
            cmd.Parameters.AddWithValue("IDCat", Categ)
            Dim Adaptador As New MySqlDataAdapter(cmd)
            Dim Tabla As New DataTable
            Adaptador.Fill(Tabla)
            DataGridView2.DataSource = Tabla
            DataGridView2.ReadOnly = False
            DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            DataGridView2.Rows(0).Selected = True
            DataGridView2.CurrentCell = DataGridView2.Rows(0).Cells(1)
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "ERROR")
            conn.Close()
        End Try
        DataGridView2.ColumnHeadersVisible = False
    End Sub

    Private Sub Llenar_Combos_Ubicaciones()
        With Estantes
            Try
                Dim conexion As New MySqlConnection(datasource)
                conexion.Open()
                Dim consulta As String = "Select * from datos_app"
                Dim MysqlDadap As New MySqlDataAdapter(consulta, conexion)
                Dim MysqlDset As New DataSet
                MysqlDadap.Fill(MysqlDset)
                conexion.Close()
                .Items.Clear()
                Dim a As String = MysqlDset.Tables(0).Rows(0).Item(2)
                Dim i1 As Integer = 0
                For i1 = 1 To a
                    .Items.Add(i1)
                Next
                .SelectedValue = Estantes.Items(0)
                '.Enabled = False
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End With
        With Entrepanos
            Try
                Dim conexion As New MySqlConnection(datasource)
                conexion.Open()
                Dim consulta As String = "Select * from datos_app"
                Dim MysqlDadap As New MySqlDataAdapter(consulta, conexion)
                Dim MysqlDset As New DataSet
                MysqlDadap.Fill(MysqlDset)
                conexion.Close()
                .Items.Clear()
                Dim a As String = MysqlDset.Tables(0).Rows(1).Item(2)
                Dim i1 As Integer = 0
                For i1 = 1 To a
                    .Items.Add(i1)
                Next
                .SelectedValue = Entrepanos.Items(0)
                '.Enabled = False
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End With
        With Cajas_Colores
            Try
                Dim conexion As New MySqlConnection(datasource)
                conexion.Open()
                Dim consulta As String = "Select * from datos_app"
                Dim MysqlDadap As New MySqlDataAdapter(consulta, conexion)
                Dim MysqlDset As New DataSet
                MysqlDadap.Fill(MysqlDset)
                conexion.Close()
                .Items.Clear()
                Dim a As String = MysqlDset.Tables(0).Rows(2).Item(2)
                Dim i1 As Integer = 0
                For i1 = 1 To a
                    .Items.Add(i1)
                Next
                .Items.Add(MysqlDset.Tables(0).Rows(3).Item(1))
                .Items.Add(MysqlDset.Tables(0).Rows(4).Item(1))
                .Items.Add(MysqlDset.Tables(0).Rows(5).Item(1))
                .Items.Add(MysqlDset.Tables(0).Rows(6).Item(1))
                .SelectedValue = Cajas_Colores.Items(0)
                '.Enabled = False
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End With
        With Zonas
            Try
                .Items.Clear()
                .Items.Add("Bodega")
                .Items.Add("Reactivos")
                .Items.Add("Gases")
                .Items.Add("Acceso Especial")
                .SelectedValue = Zonas.Items(0)
                '.Enabled = False
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End With
    End Sub

    Private Sub CargarDGVMontos()
        Cargar_Tabla("*", "Doag")
        Nombre_Doag.Text = Tabla1.Rows(0).ItemArray(1).ToString
        Monto_Doag.Text = Tabla1.Rows(0).ItemArray(2).ToString
        Comentario_Doag.Text = Tabla1.Rows(0).ItemArray(3).ToString
        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = Tabla1
        DataGridView1.ColumnHeadersVisible = False
        DataGridView1.ReadOnly = True
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(2).DefaultCellStyle.Format = "c"
    End Sub

    Private Sub CargarPerfiles()
        Cargar_Tabla("*", "Perfiles")
        Nombre_Perfil.Text = Tabla1.Rows(0).ItemArray(1).ToString
        Nivel_Permisos.Text = Tabla1.Rows(0).ItemArray(2).ToString
    End Sub

    Dim usuario_num As Integer = 1
    Dim Id_Usuario As String
    Dim Tabla1 As New DataTable

    Private Sub Recorrer_Usuarios()
        Label10.Visible = True
        Anterior_Usuario.Visible = True
        Siguiente_Usuario.Visible = True
        Cargar_Tabla("Id_Usuario, Nombre_Usuario, Usuario, Id_Perfil, Foto, Id_Doag", "USUARIOS")
        If usuario_num >= Tabla1.Rows.Count Then
            usuario_num = Tabla1.Rows.Count
        End If
        Label10.Text = "Usuario " & (usuario_num) & " de " & (Tabla1.Rows.Count)
        Id_Usuario = Tabla1.Rows(usuario_num - 1).ItemArray(0).ToString
        Nombre_Usuario.Text = Tabla1.Rows(usuario_num - 1).ItemArray(1).ToString
        Usuario_Nickname.Text = Tabla1.Rows(usuario_num - 1).ItemArray(2).ToString
        Perfiles_Usuario.SelectedValue = Convert.ToInt64(Tabla1.Rows(usuario_num - 1).ItemArray(3))
        Doag_Usuarios.SelectedValue = Convert.ToInt64(Tabla1.Rows(usuario_num - 1).ItemArray(5))
        Try
            Dim b64str As String = Tabla1.Rows(usuario_num - 1).ItemArray(4).ToString
            Dim binaryData() As Byte = Convert.FromBase64String(b64str)
            Dim stream As New MemoryStream(binaryData)
            Foto_Usuario.Image = Image.FromStream(stream)
        Catch ex As Exception
            Foto_Usuario.Image = My.Resources.NoImage
        End Try
    End Sub

    Private Sub Siguiente_Usuario_Click(sender As Object, e As EventArgs) Handles Siguiente_Usuario.Click
        usuario_num += 1
        If usuario_num > (Tabla1.Rows.Count) Then
            usuario_num = 1
            Recorrer_Usuarios()
            Exit Sub
        End If
        Recorrer_Usuarios()
    End Sub

    Private Sub Anterior_Usuario_Click(sender As Object, e As EventArgs) Handles Anterior_Usuario.Click
        usuario_num -= 1
        Dim a = Tabla1.Rows.Count
        If usuario_num = 0 Then
            usuario_num = Tabla1.Rows.Count
            Recorrer_Usuarios()
            Exit Sub
        End If
        Recorrer_Usuarios()

    End Sub

    Private Sub Cargar_Tabla(ByVal Columns As String, ByVal Nombre_Tabla As String)
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(String.Format("SELECT " & Columns & " from " & Nombre_Tabla & ";"), conn)
            Dim Adaptador As New MySqlDataAdapter(cmd)
            Dim Tabla As New DataTable
            Adaptador.Fill(Tabla)
            Tabla1 = Tabla
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
            conn.Close()
        End Try
    End Sub

    Private Sub Foto_Usuario_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Foto_Usuario.DragEnter
        If Agregar_Usuario = 1 Then
            Exit Sub
        End If
        'DataFormats.FileDrop nos devuelve el array de rutas de archivos
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'Los archivos son externos a nuestra aplicación por lo que de indicaremos All ya que dará lo mismo.
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub Foto_Usuario_DragDrop(ByVal sender As Object, e As DragEventArgs) Handles Foto_Usuario.DragDrop
        If Agregar_Usuario = 1 Then
            Exit Sub
        End If

        If MessageBox.Show("¿Esta seguro que desea CAMBIAR la foto?" & vbCrLf & "Esta accion no se puede deshacer", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim strRutaArchivoImagen As String
                strRutaArchivoImagen = e.Data.GetData(DataFormats.FileDrop)(0)
                If Path.GetExtension(strRutaArchivoImagen) = ".jpg" Or Path.GetExtension(strRutaArchivoImagen) = ".png" Or Path.GetExtension(strRutaArchivoImagen) = ".bmp" Then

                    CambiarImagenBD(strRutaArchivoImagen, "Usuarios")
                Else

                    MsgBox("El formato (" & Path.GetExtension(strRutaArchivoImagen) & ") no es soportado", MsgBoxStyle.Critical, "Error")
                End If
            End If
        End If
    End Sub

    Private Sub CambiarImagenBD(ByVal strRutaArchivoImagen As String, ByVal Tabla As String)
        Select Case Tabla
            Case "Usuarios"
                Try
                    Foto_Usuario.Image.Dispose()
                    Foto_Usuario.Image = Nothing
                    Dim FileSize As UInt32
                    Dim rawData() As Byte
                    Dim fs As FileStream
                    fs = New FileStream(strRutaArchivoImagen, FileMode.Open, FileAccess.Read)
                    FileSize = fs.Length - 1

                    rawData = New Byte(FileSize) {}
                    fs.Read(rawData, 0, FileSize)
                    fs.Close()
                    conn.Open()
                    Dim foto As String = Convert.ToBase64String(rawData)
                    Dim cmd As New MySqlCommand(String.Format("UPDATE usuarios set `Foto` = '" & foto & "' where ID_usuario = '" & Id_Usuario & "';"), conn)
                    cmd.ExecuteNonQuery()
                    conn.Close()
                    Foto_Usuario.Image = Image.FromFile(strRutaArchivoImagen)
                Catch ex As Exception
                    MsgBox("Error al cambiar la imagen, revise el estado del archivo y su conexion a la base de datos" & vbCrLf & "Error:" & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                End Try
                Exit Sub
        End Select
    End Sub

    Dim Modificar_Doag As Integer = 0

    Private Sub BtnModificarDoag_Click(sender As Object, e As EventArgs) Handles BtnModificarDoag.Click
        Agregar_Doag = 0
        If Modificar_Doag = 1 Then
            Modificar_Doag = 0
        ElseIf Modificar_Doag = 0 Then
            Modificar_Doag = 1
        End If
        HabilitarControlesDoag()

    End Sub

    Private Sub HabilitarControlesDoag()
        If Modificar_Doag = 1 Or Agregar_Doag = 1 Then
            Nombre_Doag.ReadOnly = False
            Monto_Doag.ReadOnly = False
            Comentario_Doag.ReadOnly = False
        Else
            Nombre_Doag.ReadOnly = True
            Monto_Doag.ReadOnly = True
            Comentario_Doag.ReadOnly = True
        End If
    End Sub

    Dim Modificar_Usuario As Integer = 0

    Private Sub BtnModificarUsuario_Click(sender As Object, e As EventArgs) Handles BtnModificarUsuario.Click
        Agregar_Usuario = 0
        If Modificar_Usuario = 1 Then
            Modificar_Usuario = 0
        ElseIf Modificar_Usuario = 0 Then
            Modificar_Usuario = 1
        End If
        HabilitarControlesUsuario()
        ActivarCamposContrasena()
        Label10.Visible = True
        Anterior_Usuario.Visible = True
        Siguiente_Usuario.Visible = True
        Recorrer_Usuarios()
    End Sub

    Dim Agregar_Usuario As Integer = 0
    Private Sub Nuevo_Usuario_Click(sender As Object, e As EventArgs) Handles Nuevo_Usuario.Click
        Modificar_Usuario = 0
        Agregar_Usuario = 1
        HabilitarControlesUsuario()
        ActivarCamposContrasena()
        Nombre_Usuario.Clear()
        Nombre_Usuario.Focus()
        Usuario_Nickname.Clear()
        Foto_Usuario.Image = My.Resources.NoImage
        Label10.Visible = False
        Anterior_Usuario.Visible = False
        Siguiente_Usuario.Visible = False
    End Sub

    Dim Agregar_Doag As Integer = 0
    Private Sub Nuevo_Doag_Click(sender As Object, e As EventArgs) Handles Nuevo_Doag.Click
        Modificar_Doag = 0
        Agregar_Doag = 1
        Nombre_Doag.Clear()
        Monto_Doag.Clear()
        Comentario_Doag.Clear()
        Nombre_Doag.Focus()
        HabilitarControlesDoag()
    End Sub

    Private Sub ActivarCamposContrasena()
        If Agregar_Usuario = 1 Then
            Label21.Visible = True
            Label21.Enabled = True
            Contrasena_Usuario.Enabled = True
            Contrasena_Usuario.Visible = True
            Contrasena_Usuario.Clear()
        Else
            Label21.Visible = False
            Label21.Enabled = False
            Contrasena_Usuario.Enabled = False
            Contrasena_Usuario.Visible = False
            Contrasena_Usuario.Clear()
        End If
    End Sub

    Private Sub HabilitarControlesUsuario()
        If Modificar_Usuario = 1 Or Agregar_Usuario = 1 Then
            Nombre_Usuario.ReadOnly = False
            Usuario_Nickname.ReadOnly = False
            Perfiles_Usuario.Enabled = True
            Doag_Usuarios.Enabled = True
        Else
            Nombre_Usuario.ReadOnly = True
            Usuario_Nickname.ReadOnly = True
            Perfiles_Usuario.Enabled = False
            Doag_Usuarios.Enabled = False
        End If
    End Sub

    Private Sub TabPage7_Leave(sender As Object, e As EventArgs) Handles TabPage7.Leave
        Modificar_Usuario = 0
        Agregar_Usuario = 0
        HabilitarControlesUsuario()
        ActivarCamposContrasena()
    End Sub

    Private Sub TabPage8_Leave(sender As Object, e As EventArgs) Handles TabPage8.Leave
        Modificar_Doag = 0
        Agregar_Doag = 0
        HabilitarControlesDoag()
    End Sub


    Dim Id_Doag As String
    Private Sub DataGridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Try
            Dim fila_actual As Integer = (DataGridView1.CurrentRow.Index)
            If fila_actual = (DataGridView1.Rows.Count - 1) Then
                'reg_bus_doag = "nuevo"
                Nombre_Doag.Text = ""
                Monto_Doag.Text = ""
                Comentario_Doag.Text = ""
            Else
                Cargar_Tabla("*", "Doag")
                Id_Doag = Tabla1.Rows(fila_actual).ItemArray(0).ToString
                Nombre_Doag.Text = Tabla1.Rows(fila_actual).ItemArray(1).ToString
                Monto_Doag.Text = Format(Tabla1.Rows(fila_actual).ItemArray(2).ToString, "Currency") 'Text1.Text = Format(Numero, "Currency")
                Comentario_Doag.Text = Tabla1.Rows(fila_actual).ItemArray(3).ToString
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub Guardar_Doag_Click(sender As Object, e As EventArgs) Handles Guardar_Doag.Click
        If Modificar_Doag = 1 Then
            Dim Nombre As String = Nombre_Doag.Text.Trim
            Dim Suma As String = CType(Monto_Doag.Text.Trim, Integer).ToString
            Dim Comentario As String = Comentario_Doag.Text.Trim
            Try
                conn.Open()
                Dim query As String = "UPDATE doag SET Nombre_Doag = @Nombre, Monto = @Monto, Comentario = @Comentario
                                       WHERE Id_Doag = @IDDoag;"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Nombre", Nombre)
                    .AddWithValue("Monto", Suma)
                    .AddWithValue("Comentario", Comentario)
                    .AddWithValue("IDDoag", Id_Doag)
                End With
                cmd.ExecuteNonQuery()
                MsgBox("Registro modificado", MsgBoxStyle.Information, "Info.")
                conn.Close()
                CargarDGVMontos()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try

        ElseIf Agregar_Doag = 1 Then
            Dim Nombre As String = Nombre_Doag.Text
            Dim Monto As String = CType(Monto_Doag.Text, Integer).ToString
            Dim Comentario As String = Comentario_Doag.Text.Trim
            Try
                conn.Open()
                Dim query As String = "INSERT INTO doag (Nombre_Doag, Monto, Comentario) VALUES (@Nombre, @Monto, @Comentario);"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Nombre", Nombre)
                    .AddWithValue("@Monto", Monto)
                    .AddWithValue("@Comentario", Comentario)
                End With
                cmd.ExecuteNonQuery()
                MsgBox("Registro Agregado", MsgBoxStyle.Information, "Info.")
                conn.Close()
                CargarDGVMontos()
            Catch ex As Exception
                MsgBox("Error al intentar agregar el registro: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try
            Agregar_Doag = 0
            HabilitarControlesDoag()
        End If
    End Sub

    Private Sub Guardar_Usuario_Click(sender As Object, e As EventArgs) Handles Guardar_Usuario.Click
        If Modificar_Usuario = 1 Then
            Dim reader As MySqlDataReader
            Dim Nombre As String = Nombre_Usuario.Text.Trim
            Dim Usuario As String = Usuario_Nickname.Text.Trim
            Dim IDPerfil As String = Perfiles_Usuario.SelectedValue.ToString
            Dim IDDoag As String = Doag_Usuarios.SelectedValue.ToString

            Try
                conn.Open()
                Dim query As String = "UPDATE usuarios SET Nombre_Usuario = @nombre, Usuario = @usuario
                                , Id_Perfil = @IDPerfil, Id_Doag = @doag WHERE Id_Usuario = @IDUsu;"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("nombre", Nombre)
                    .AddWithValue("usuario", Usuario)
                    .AddWithValue("IDPerfil", IDPerfil)
                    .AddWithValue("doag", IDDoag)
                    .AddWithValue("IDUsu", Id_Usuario)
                End With
                reader = cmd.ExecuteReader
                MsgBox("Usuario modificado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try

        ElseIf Agregar_Usuario = 1 Then

            Dim reader As MySqlDataReader
            Dim Nombre As String = Nombre_Usuario.Text.Trim
            Dim Usuario As String = Usuario_Nickname.Text.Trim
            Dim IDPerfil As String = Perfiles_Usuario.SelectedValue.ToString
            Dim IDDoag As String = Doag_Usuarios.SelectedValue.ToString
            Dim Contrasena As String = Contrasena_Usuario.Text
            Dim NewSalt As String = GenerateSalt()
            If Contrasena = "" Then
                Contrasena = "123"
            End If
            If Nombre = "" Or Usuario = "" Then
                MsgBox("Todos los campos son obligatorios", MsgBoxStyle.Exclamation, "Error")
                Exit Sub
            End If

            Contrasena = NewSalt + Contrasena
            Contrasena = ComputeHashOfString(Of SHA256CryptoServiceProvider)(Contrasena)

            Try
                conn.Open()
                Dim query As String = "INSERT into usuarios (Nombre_Usuario, Usuario, Salt, Hash, Id_Perfil, Id_Doag)
                                      VALUES (@nombre, @usuario, @Salt, @Hash, @IDPerfil, @doag);"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("nombre", Nombre)
                    .AddWithValue("usuario", Usuario)
                    .AddWithValue("Salt", NewSalt)
                    .AddWithValue("Hash", Contrasena)
                    .AddWithValue("IDPerfil", IDPerfil)
                    .AddWithValue("doag", IDDoag)
                    .AddWithValue("IDUsu", Id_Usuario)
                End With
                reader = cmd.ExecuteReader
                MsgBox("Usuario Agregado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try
            Agregar_Usuario = 0
            HabilitarControlesUsuario()
            ActivarCamposContrasena()
            Recorrer_Usuarios()
        End If
    End Sub

    Public Function ComputeHashOfString(Of T As HashAlgorithm)(ByVal str As String,
                                                                             Optional ByVal enc As Encoding = Nothing) As String
        If (enc Is Nothing) Then
            enc = Encoding.Default
        End If
        Using algorithm As HashAlgorithm = DirectCast(Activator.CreateInstance(GetType(T)), HashAlgorithm)
            Dim data As Byte() = enc.GetBytes(str)
            Dim hash As Byte() = algorithm.ComputeHash(data)
            Dim sb As New StringBuilder(capacity:=hash.Length * 2)
            For Each b As Byte In hash
                sb.Append(b.ToString("X2"))
            Next
            Return sb.ToString.ToLower()
        End Using

    End Function

    Private Function GenerateSalt()
        Dim saltsize As Integer = 47
        Dim saltbytes() As Byte
        saltbytes = New Byte(saltsize - 1) {}
        Dim rng As RNGCryptoServiceProvider
        rng = New RNGCryptoServiceProvider
        rng.GetNonZeroBytes(saltbytes)
        Return Convert.ToBase64String(saltbytes)
    End Function

    Private Sub Eliminar_Usuario_Click(sender As Object, e As EventArgs) Handles Eliminar_Usuario.Click
        If Agregar_Usuario = 1 Then
            Exit Sub
        End If
        If MessageBox.Show("¿Esta seguro que desea ELIMINAR este Usuario?", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Try
                conn.Open()
                Dim query As String = "Delete from Usuarios where ID_Usuario = @IdUsu;"
                Dim cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("IdUsu", Id_Usuario)
                cmd.ExecuteNonQuery()
                MsgBox("Usuario Eliminado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End If
        Recorrer_Usuarios()
    End Sub


    Private Sub Eliminar_Doag_Click(sender As Object, e As EventArgs) Handles Eliminar_Doag.Click
        If Agregar_Doag = 1 Then
            Exit Sub
        End If
        If MessageBox.Show("¿Esta seguro que desea ELIMINAR este Registro?", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Try
                conn.Open()
                Dim query As String = "Delete from doag where Id_Doag = @IdDoag;"
                Dim cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("IdDoag", Id_Doag)
                cmd.ExecuteNonQuery()
                MsgBox("Registro Eliminado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End If
        CargarDGVMontos()
    End Sub


    Dim cant_reg_encon As Integer = 0
    Dim z As String 'memorioa del usuario a buscar

    Private Sub Buscar_Usuario_Click(sender As Object, e As EventArgs) Handles Buscar_Usuario.Click
        If Buscar_Us.Text.Trim = "" Then
            Exit Sub
        End If
        Cargar_Tabla("Id_Usuario, Nombre_Usuario, Usuario, Id_Perfil, Foto, Id_Doag", "USUARIOS")
        If z <> Buscar_Us.Text Then
            cant_reg_encon = 0
        End If
        Try
            conn.Open()
            Dim consulta As String = "Select * from usuarios"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            conn.Close()
            Dim i As Integer = 0
            Dim foundRows() As Data.DataRow
            foundRows = MysqlDset.Tables(0).Select("Nombre_Usuario Like '" & Buscar_Us.Text & "%'")
            z = Buscar_Us.Text
            If cant_reg_encon = 0 And foundRows.Length > 1 Then
                cant_reg_encon = foundRows.Length
                For Each row In Tabla1.Rows
                    If foundRows(cant_reg_encon - 1).Item(1) = row(1) Then
                        'MsgBox(foundRows(cant_reg_encon - 1).Item(1))
                        usuario_num = i + 1
                        Recorrer_Usuarios()
                        cant_reg_encon = cant_reg_encon - 1
                        Exit Sub
                    End If
                    i = i + 1
                Next
            Else
                If foundRows.Length = 0 Then
                    MsgBox("No se encontro ninguna coincidencia")
                ElseIf cant_reg_encon = 0 Then
                    For Each row In Tabla1.Rows
                        If foundRows(cant_reg_encon).Item(1) = row(1) Then
                            usuario_num = i + 1
                            Recorrer_Usuarios()
                            Exit Sub
                        End If
                        i = i + 1
                    Next
                Else
                    For Each row In Tabla1.Rows
                        If foundRows(cant_reg_encon - 1).Item(1) = row(1) Then
                            usuario_num = i + 1
                            Recorrer_Usuarios()
                            cant_reg_encon = cant_reg_encon - 1
                            Exit Sub
                        End If
                        i = i + 1
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox("Error durante la busqueda: " & ex.Message, MsgBoxStyle.Critical, "Error")
            conn.Close()
        End Try
    End Sub

    Private Sub Buscar_Us_KeyDown(sender As Object, e As KeyEventArgs) Handles Buscar_Us.KeyDown
        If e.KeyCode = Keys.Enter Then
            Buscar_Usuario.PerformClick()
        End If
    End Sub

    Private Sub TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Monto_Doag.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

    End Sub

    Private Sub Monto_Doag_Leave(sender As Object, e As EventArgs) Handles Monto_Doag.Leave
        If Monto_Doag.Text = "" Then
            Exit Sub
        End If
        Monto_Doag.Text = FormatCurrency(Monto_Doag.Text)
    End Sub

    Private Sub Nombre_Categoria_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Nombre_Categoria.SelectedIndexChanged
        If CategLoad = False Then
            Exit Sub
        End If
        DataGridView2.DataSource = Nothing
        CargarSubCategorias()
    End Sub


End Class
