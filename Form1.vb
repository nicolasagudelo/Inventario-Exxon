Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Data
Imports System.Windows.Forms
Imports System.Configuration
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
        'cant_reg_encon = 0
        'z = "USUARIOS"
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
        Esconder_tabpages_submenu()
        TabPage7.Parent = TabControl2 'Usuarios
        Recorrer_Usuarios()
        TabPage8.Parent = TabControl2 'Monto
        Cargar_Tabla("*", "Doag")
        Nombre_Doag.Text = Tabla1.Rows(0).ItemArray(1).ToString
        Monto_Doag.Text = Tabla1.Rows(0).ItemArray(2).ToString
        Comentario_Doag.Text = Tabla1.Rows(0).ItemArray(3).ToString
        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = Tabla1
        DataGridView1.ColumnHeadersVisible = False
        DataGridView1.ReadOnly = True
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.Columns(1).Width = 160
        DataGridView1.Columns(2).Width = 160
        DataGridView1.Columns(3).Width = 310
        DataGridView1.Columns(0).Visible = False
        TabPage9.Parent = TabControl2 'Perfiles
        Cargar_Tabla("*", "Perfiles")
        Nombre_Perfil.Text = Tabla1.Rows(0).ItemArray(1).ToString
        Nivel_Permisos.Text = Tabla1.Rows(0).ItemArray(2).ToString

    End Sub

    Dim usuario_num As Integer = 0
    Dim Id_Usuario As String
    Dim Tabla1 As New DataTable

    Private Sub Recorrer_Usuarios()
        Cargar_Tabla("Id_Usuario, Nombre_Usuario, Usuario, Id_Perfil, Foto, Id_Doag", "USUARIOS")
        Label10.Text = "Usuario " & (usuario_num + 1) & " de " & (Tabla1.Rows.Count)
        Id_Usuario = Tabla1.Rows(usuario_num).ItemArray(0).ToString
        Nombre_Usuario.Text = Tabla1.Rows(usuario_num).ItemArray(1).ToString
        Usuario_Nickname.Text = Tabla1.Rows(usuario_num).ItemArray(2).ToString
        Perfiles_Usuario.SelectedValue = Convert.ToInt64(Tabla1.Rows(usuario_num).ItemArray(3))
        Doag_Usuarios.SelectedValue = Convert.ToInt64(Tabla1.Rows(usuario_num).ItemArray(5))
        Try
            Dim b64str As String = Tabla1.Rows(usuario_num).ItemArray(4).ToString
            Dim binaryData() As Byte = Convert.FromBase64String(b64str)
            Dim stream As New MemoryStream(binaryData)
            Foto_Usuario.Image = Image.FromStream(stream)
        Catch ex As Exception
            Foto_Usuario.Image = My.Resources.NoImage
        End Try
    End Sub

    Private Sub Siguiente_Usuario_Click(sender As Object, e As EventArgs) Handles Siguiente_Usuario.Click
        If usuario_num >= (Tabla1.Rows.Count - 1) Then
            usuario_num = 0
            Recorrer_Usuarios()
        Else
            usuario_num = usuario_num + 1
            Recorrer_Usuarios()
        End If
    End Sub
    Private Sub Anterior_Usuario_Click(sender As Object, e As EventArgs) Handles Anterior_Usuario.Click
        If usuario_num = 0 Then
            usuario_num = Tabla1.Rows.Count - 1
            Recorrer_Usuarios()
        Else
            usuario_num = usuario_num - 1
            Recorrer_Usuarios()
        End If
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
        'DataFormats.FileDrop nos devuelve el array de rutas de archivos
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'Los archivos son externos a nuestra aplicación por lo que de indicaremos All ya que dará lo mismo.
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub Foto_Usuario_DragDrop(ByVal sender As Object, e As DragEventArgs) Handles Foto_Usuario.DragDrop

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

End Class
