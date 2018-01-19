Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Data
Imports System.Windows.Forms
Imports System.Configuration
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Net.Mail

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
        For i = 7 To 18
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

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Menu_Seleccionado(3)
        Esconder_tabpages_submenu()
        TabPage13.Parent = TabControl2
        cant_reg_encon = 0
        z = "EQUIPOS"
        Modificar_Equipo = 0
        Agregar_Equipo = 0
        HabilitarControlesEquipo()
        Recorrer_Equipos()
    End Sub

    Dim Equipo_num As Integer = 1
    Dim Id_Equipo As String
    Private Sub Recorrer_Equipos()
        Label60.Visible = True
        Anterior_Equipo.Visible = True
        Siguiente_Equipo.Visible = True
        Cargar_Tabla("*", "Equipos")
        If Equipo_num >= Tabla1.Rows.Count Then
            Equipo_num = Tabla1.Rows.Count
        End If
        Label60.Text = "Equipo " & (Equipo_num) & " de " & (Tabla1.Rows.Count)
        Id_Equipo = Tabla1.Rows(Equipo_num - 1).ItemArray(0).ToString
        Numero_Equipo.Text = Tabla1.Rows(Equipo_num - 1).ItemArray(1).ToString
        Nombre_Equipo.Text = Tabla1.Rows(Equipo_num - 1).ItemArray(2).ToString
        Marca_Equipo.Text = Tabla1.Rows(Equipo_num - 1).ItemArray(3).ToString
        Serie_Equipo.Text = Tabla1.Rows(Equipo_num - 1).ItemArray(4).ToString
        Activo_Equipo.Checked = Tabla1.Rows(Equipo_num - 1).ItemArray(6).ToString
        Try
            Dim b64str As String = Tabla1.Rows(Equipo_num - 1).ItemArray(5).ToString
            Dim binarydata() As Byte = Convert.FromBase64String(b64str)
            Dim stream As New MemoryStream(binarydata)
            Foto_Equipo.Image = Image.FromStream(stream)
        Catch ex As Exception
            Foto_Equipo.Image = My.Resources.NoMachine
        End Try
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Menu_Seleccionado(4)
        Esconder_tabpages_submenu()
        TabPage14.Parent = TabControl2
        cant_reg_encon = 0
        z = "PRODUCTOS"
        Recorrer_Productos()
    End Sub

    Dim Id_Prod As String
    Dim Prod_Num As Integer = 1
    Private Sub Recorrer_Productos()
        Label72.Visible = True
        Anterior_Producto.Visible = True
        Siguiente_Producto.Visible = True
        Cargar_Tabla("*", "Productos")
        If Prod_Num >= Tabla1.Rows.Count Then
            Prod_Num = Tabla1.Rows.Count
        End If
        Label72.Text = "Producto " & (Prod_Num) & " de " & (Tabla1.Rows.Count)
        Id_Prod = Tabla1.Rows(Prod_Num - 1).ItemArray(0).ToString
        Codigo_Producto.Text = Tabla1.Rows(Prod_Num - 1).ItemArray(1).ToString
        Nombre_Producto.Text = Tabla1.Rows(Prod_Num - 1).ItemArray(2).ToString
        Marca_Producto.Text = Tabla1.Rows(Prod_Num - 1).ItemArray(5).ToString
        Serie_Producto.Text = Tabla1.Rows(Prod_Num - 1).ItemArray(6).ToString
        Stock_Minimo.Text = Tabla1.Rows(Prod_Num - 1).ItemArray(8).ToString
        Stock_Maximo.Text = Tabla1.Rows(Prod_Num - 1).ItemArray(9).ToString
        Stock_Existente.Text = Tabla1.Rows(Prod_Num - 1).ItemArray(10).ToString
        Compra_Maxima.Text = Tabla1.Rows(Prod_Num - 1).ItemArray(12).ToString
        Activo_Producto.Checked = Tabla1.Rows(Prod_Num - 1).ItemArray(13).ToString
        Unidades_Producto.Items.Clear()
        With Unidades_Producto
            .Items.Add("Uni")
            .Items.Add("Kg")
            .Items.Add("Lb")
            .Items.Add("L")
            .Items.Add("Gal")
        End With
        Unidades_Producto.SelectedItem = Tabla1.Rows(Prod_Num - 1).ItemArray(11).ToString
        Try
            Dim b64str As String = Tabla1.Rows(Prod_Num - 1).ItemArray(7).ToString
            Dim binarydata() As Byte = Convert.FromBase64String(b64str)
            Dim stream As New MemoryStream(binarydata)
            Foto_Producto.Image = Image.FromStream(stream)
        Catch ex As Exception
            Foto_Producto.Image = My.Resources.NoMachine
        End Try
        CargarCBTabProductos(Prod_Num - 1)
    End Sub

    Private Sub CargarCBTabProductos(ByVal numero_Producto As Integer)
        With Categoria_Producto
            Try
                conn.Open()
                Dim consulta As String = "Select * from categorias"
                Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
                Dim MysqlDset As New DataSet
                MysqlDadap.Fill(MysqlDset)
                conn.Close()
                .DataSource = MysqlDset.Tables(0)
                .DisplayMember = "Nombre_Categoria" 'elnombre de tu columna de tu base de datos q deseas mostrar
                .ValueMember = "Id_Categoria" 'el ide de tu tabla relacionada con el nombre que muestras muy importante para saber el ide de quien seleccionas en tu combobox
                .SelectedValue = Tabla1.Rows(numero_Producto).ItemArray(3)
                '.Enabled = False
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                conn.Close()
            End Try
        End With
        With SubCategoria_Producto
            Try
                conn.Open()
                Dim consulta As String = "Select * from categorias_sub where Id_Categoria='" & Tabla1.Rows(numero_Producto).ItemArray(3) & "'"
                Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
                Dim MysqlDset As New DataSet
                MysqlDadap.Fill(MysqlDset)
                conn.Close()
                .DataSource = MysqlDset.Tables(0)
                .DisplayMember = "Nombre_SubCategoria" 'elnombre de tu columna de tu base de datos q deseas mostrar
                .ValueMember = "Id_SubCategoria" 'el ide de tu tabla relacionada con el nombre que muestras muy importante para saber el ide de quien seleccionas en tu combobox
                .SelectedValue = Tabla1.Rows(numero_Producto).ItemArray(4)
                '.Enabled = False
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                conn.Close()
            End Try
        End With
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Menu_Seleccionado(5)
        Esconder_tabpages_submenu()
        TabPage15.Parent = TabControl2
        cant_reg_encon = 0
        z = "PROVEEDORES"
        Recorrer_Proveedores()
    End Sub


    Dim Proveedor_Num As Integer = 1
    Dim ID_Prov As String
    Private Sub Recorrer_Proveedores()
        Label59.Visible = True
        Anterior_Proveedor.Visible = True
        Siguiente_Proveedor.Visible = True
        Cargar_Tabla("*", "Proveedores")
        If Proveedor_Num >= Tabla1.Rows.Count Then
            Proveedor_Num = Tabla1.Rows.Count
        End If
        Label59.Text = "Proveedor " & (Proveedor_Num) & " de " & (Tabla1.Rows.Count)
        ID_Prov = Tabla1.Rows(Proveedor_Num - 1).ItemArray(0).ToString
        Nit_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(0).ToString
        Nombre_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(1).ToString
        Contacto_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(2).ToString
        Direccion_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(3).ToString
        Ciudad_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(4).ToString
        Telefono_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(5).ToString
        Email_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(6).ToString
        Fax_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(7).ToString
        Web_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(8).ToString
        Detalle_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(9).ToString
        Clasificacion_Proveedor.Text = Tabla1.Rows(Proveedor_Num - 1).ItemArray(10).ToString
        Aprovado_Proveedor.Checked = Tabla1.Rows(Proveedor_Num - 1).ItemArray(11).ToString
        Activo_Proveedor.Checked = Tabla1.Rows(Proveedor_Num - 1).ItemArray(13).ToString
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        Menu_Seleccionado(6)
        Esconder_tabpages_submenu()
        TabPage16.Parent = TabControl2
        With ComboBox1
            .DataSource = Nothing
            .Items.Clear()
            Try
                conn.Open()
                Dim query As String = "Select Id_Producto, Nombre_Producto from productos"
                Dim cmd As New MySqlCommand(query, conn)
                Dim sqlAdap As New MySqlDataAdapter(cmd)
                Dim dtRecord As New DataTable
                sqlAdap.Fill(dtRecord)
                Dim Todos As DataRow = dtRecord.NewRow
                Todos("Id_Producto") = "-1"
                Todos("Nombre_Producto") = "TODOS"
                dtRecord.Rows.InsertAt(Todos, 0)
                .DataSource = dtRecord
                .DisplayMember = "Nombre_Producto"
                .ValueMember = "Id_Producto"
                .SelectedValue = dtRecord.Rows(0).Item(0)
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar los productos de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End With

        With ComboBox2
            .DataSource = Nothing
            .Items.Clear()
            Try
                conn.Open()
                Dim query As String = "Select Nit_Proveedor, Nombre_Proveedor from proveedores"
                Dim cmd As New MySqlCommand(query, conn)
                Dim sqlAdap As New MySqlDataAdapter(cmd)
                Dim dtRecord As New DataTable
                sqlAdap.Fill(dtRecord)
                Dim Todos As DataRow = dtRecord.NewRow
                Todos("Nit_Proveedor") = "-1"
                Todos("Nombre_Proveedor") = "TODOS"
                dtRecord.Rows.InsertAt(Todos, 0)
                .DataSource = dtRecord
                .DisplayMember = "Nombre_Proveedor"
                .ValueMember = "Nit_Proveedor"
                .SelectedValue = dtRecord.Rows(0).Item(0)
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar los productos de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End With
        FechaInicio.Format = DateTimePickerFormat.Custom
        FechaInicio.CustomFormat = "yyyy-MM-dd"
        FechaInicio.Value = FechaInicio.MinDate
        FechaFin.Format = DateTimePickerFormat.Custom
        FechaFin.CustomFormat = "yyyy-MM-dd"

        Try
            conn.Open()
            Dim cmd As New MySqlCommand(String.Format("SELECT NOW();"), conn)
            Dim fecha_servidor As DateTime = cmd.ExecuteScalar()
            FechaFin.MaxDate = fecha_servidor.ToString("yyyy-MM-dd")
            FechaInicio.MaxDate = FechaFin.MaxDate
            FechaFin.Value = FechaFin.MaxDate
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, False, "No se puede obtener la fecha de la base de datos se tomara la hora local")
            conn.Close()
            FechaFin.MaxDate = DateTime.Now.ToString("yyyy-MM-dd")
            FechaInicio.MaxDate = FechaFin.MaxDate
            FechaFin.Value = FechaFin.MaxDate
            Exit Sub
        End Try

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
        'TabPage9.Parent = TabControl2 'Perfiles
        CargarPerfiles()
        Recorrer_Usuarios()
    End Sub


    Private Sub Gestion_Almacen_Click(sender As Object, e As EventArgs) Handles Gestion_Almacen.Click
        Modificar_Categorias = 0
        Agregar_Categoria = 0
        HabilitarControlesCategorias()
        Esconder_tabpages_submenu()
        TabPage10.Parent = TabControl2
        CargarCBTabCategorias()
    End Sub

    Dim Id_Ubicacion As Integer
    Dim CategLoad As Boolean = False
    Dim FirtsLoad As Boolean = True
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
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar las categorias de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End With
        CategLoad = True
        CargarSubCategorias()
        CargarCBTabUbicaciones()
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

    Private Sub CargarCBTabUbicaciones()
        With Estantes
            Try
                conn.Open()
                Dim consulta As String = "Select * from datos_app"
                Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
                Dim MysqlDset As New DataSet
                MysqlDadap.Fill(MysqlDset)
                .Items.Clear()
                Dim a As String = MysqlDset.Tables(0).Rows(0).Item(2)
                Dim i1 As Integer = 0
                For i1 = 1 To a
                    .Items.Add(i1)
                Next
                .SelectedValue = Estantes.Items(0)
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar los estantes de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End With
        With Entrepanos
            Try
                conn.Open()
                Dim consulta As String = "Select * from datos_app"
                Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
                Dim MysqlDset As New DataSet
                MysqlDadap.Fill(MysqlDset)
                .Items.Clear()
                Dim a As String = MysqlDset.Tables(0).Rows(1).Item(2)
                Dim i1 As Integer = 0
                For i1 = 1 To a
                    .Items.Add(i1)
                Next
                .SelectedValue = Entrepanos.Items(0)
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar los entrepaños de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End With
        With Cajas_Colores
            Try
                conn.Open()
                Dim consulta As String = "Select * from datos_app"
                Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
                Dim MysqlDset As New DataSet
                MysqlDadap.Fill(MysqlDset)
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
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
                conn.Close()
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

    Private Sub CargarSubCategorias()
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(String.Format("Select * from categorias_sub where Id_Categoria = @IDCat;"), conn)
            Dim Id_Cat As Integer = Nombre_Categoria.SelectedValue
            cmd.Parameters.AddWithValue("IDCat", Id_Cat)
            Dim Adaptador As New MySqlDataAdapter(cmd)
            Dim Tabla As New DataTable
            Adaptador.Fill(Tabla)
            DataGridView2.DataSource = Tabla
            DataGridView2.ReadOnly = False
            DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            DataGridView2.Columns(0).Visible = False
            DataGridView2.Columns(2).Visible = False
            DataGridView2.Rows(0).Selected = True
            DataGridView2.CurrentCell = DataGridView2.Rows(0).Cells(1)
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "ERROR")
            conn.Close()
        End Try
        DataGridView2.ColumnHeadersVisible = False
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
        Cargar_Tabla("Id_Usuario, Nombre_Usuario, Usuario, Id_Perfil, Foto, Id_Doag, Email", "USUARIOS")
        If usuario_num >= Tabla1.Rows.Count Then
            usuario_num = Tabla1.Rows.Count
        End If
        Label10.Text = "Usuario " & (usuario_num) & " de " & (Tabla1.Rows.Count)
        Id_Usuario = Tabla1.Rows(usuario_num - 1).ItemArray(0).ToString
        Nombre_Usuario.Text = Tabla1.Rows(usuario_num - 1).ItemArray(1).ToString
        Usuario_Nickname.Text = Tabla1.Rows(usuario_num - 1).ItemArray(2).ToString
        Perfiles_Usuario.SelectedValue = Convert.ToInt64(Tabla1.Rows(usuario_num - 1).ItemArray(3))
        Doag_Usuarios.SelectedValue = Convert.ToInt64(Tabla1.Rows(usuario_num - 1).ItemArray(5))
        TxtBxEmail.Text = Tabla1.Rows(usuario_num - 1).ItemArray(6).ToString
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

    Private Sub Siguiente_Equipo_Click(sender As Object, e As EventArgs) Handles Siguiente_Equipo.Click
        Equipo_num += 1
        If Equipo_num > (Tabla1.Rows.Count) Then
            Equipo_num = 1
            Recorrer_Equipos()
            Exit Sub
        End If
        Recorrer_Equipos()
    End Sub

    Private Sub Anterior_Equipo_Click(sender As Object, e As EventArgs) Handles Anterior_Equipo.Click
        Equipo_num -= 1
        Dim a = Tabla1.Rows.Count
        If Equipo_num = 0 Then
            Equipo_num = Tabla1.Rows.Count
            Recorrer_Equipos()
            Exit Sub
        End If
        Recorrer_Equipos()
    End Sub

    Private Sub Siguiente_producto_Click(sender As Object, e As EventArgs) Handles Siguiente_Producto.Click
        Prod_Num += 1
        If Prod_Num > (Tabla1.Rows.Count) Then
            Prod_Num = 1
            Recorrer_Productos()
            Exit Sub
        End If
        Recorrer_Productos()
    End Sub

    Private Sub Anterior_Producto_Click(sender As Object, e As EventArgs) Handles Anterior_Producto.Click
        Prod_Num -= 1
        Dim a = Tabla1.Rows.Count
        If Prod_Num = 0 Then
            Prod_Num = Tabla1.Rows.Count
            Recorrer_Productos()
            Exit Sub
        End If
        Recorrer_Productos()
    End Sub


    Private Sub Anterior_Proveedor_Click(sender As Object, e As EventArgs) Handles Anterior_Proveedor.Click
        Proveedor_Num -= 1
        Dim a = Tabla1.Rows.Count
        If Proveedor_Num = 0 Then
            Proveedor_Num = Tabla1.Rows.Count
            Recorrer_Proveedores()
            Exit Sub
        End If
        Recorrer_Proveedores()
    End Sub

    Private Sub Siguiente_Proveedor_Click(sender As Object, e As EventArgs) Handles Siguiente_Proveedor.Click
        Proveedor_Num += 1
        If Proveedor_Num > (Tabla1.Rows.Count) Then
            Proveedor_Num = 1
            Recorrer_Proveedores()
            Exit Sub
        End If
        Recorrer_Proveedores()
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

    Private Sub Foto_Equipo_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Foto_Equipo.DragEnter
        If Agregar_Equipo = 1 Then
            Exit Sub
        End If
        'DataFormats.FileDrop nos devuelve el array de rutas de archivos
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'Los archivos son externos a nuestra aplicación por lo que de indicaremos All ya que dará lo mismo.
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub Foto_Equipo_DragDrop(ByVal sender As Object, e As DragEventArgs) Handles Foto_Equipo.DragDrop
        If Agregar_Equipo = 1 Then
            Exit Sub
        End If

        If MessageBox.Show("¿Esta seguro que desea CAMBIAR la foto?" & vbCrLf & "Esta accion no se puede deshacer", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim strRutaArchivoImagen As String
                strRutaArchivoImagen = e.Data.GetData(DataFormats.FileDrop)(0)
                If Path.GetExtension(strRutaArchivoImagen) = ".jpg" Or Path.GetExtension(strRutaArchivoImagen) = ".png" Or Path.GetExtension(strRutaArchivoImagen) = ".bmp" Then

                    CambiarImagenBD(strRutaArchivoImagen, "Equipos")
                Else

                    MsgBox("El formato (" & Path.GetExtension(strRutaArchivoImagen) & ") no es soportado", MsgBoxStyle.Critical, "Error")
                End If
            End If
        End If
    End Sub

    Private Sub Foto_Producto_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Foto_Producto.DragEnter
        If Agregar_Producto = 1 Then
            Exit Sub
        End If
        'DataFormats.FileDrop nos devuelve el array de rutas de archivos
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'Los archivos son externos a nuestra aplicación por lo que de indicaremos All ya que dará lo mismo.
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub Foto_Producto_DragDrop(ByVal sender As Object, e As DragEventArgs) Handles Foto_Producto.DragDrop
        If Agregar_Producto = 1 Then
            Exit Sub
        End If

        If MessageBox.Show("¿Esta seguro que desea CAMBIAR la foto?" & vbCrLf & "Esta accion no se puede deshacer", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim strRutaArchivoImagen As String
                strRutaArchivoImagen = e.Data.GetData(DataFormats.FileDrop)(0)
                If Path.GetExtension(strRutaArchivoImagen) = ".jpg" Or Path.GetExtension(strRutaArchivoImagen) = ".png" Or Path.GetExtension(strRutaArchivoImagen) = ".bmp" Then

                    CambiarImagenBD(strRutaArchivoImagen, "Productos")
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
            Case "Equipos"
                Try
                    Foto_Equipo.Image.Dispose()
                    Foto_Equipo.Image = Nothing
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
                    Dim cmd As New MySqlCommand(String.Format("UPDATE Equipos set `Foto` = '" & foto & "' where ID_Equipo = '" & Id_Equipo & "';"), conn)
                    cmd.ExecuteNonQuery()
                    conn.Close()
                    Foto_Equipo.Image = Image.FromFile(strRutaArchivoImagen)
                Catch ex As Exception
                    MsgBox("Error al cambiar la imagen, revise el estado del archivo y su conexion a la base de datos" & vbCrLf & "Error:" & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                End Try
                Exit Sub
            Case "Productos"
                Try
                    Foto_Producto.Image.Dispose()
                    Foto_Producto.Image = Nothing
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
                    Dim cmd As New MySqlCommand(String.Format("UPDATE productos set `Foto` = '" & foto & "' where ID_Producto = '" & Id_Prod & "';"), conn)
                    cmd.ExecuteNonQuery()
                    conn.Close()
                    Foto_Producto.Image = Image.FromFile(strRutaArchivoImagen)
                Catch ex As Exception
                    MsgBox("Error al cambiar la imagen, revise el estado del archivo y su conexion a la base de datos" & vbCrLf & "Error:" & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                End Try
                Exit Sub
        End Select
    End Sub

    Dim Modificar_Categorias As Integer = 0

    Private Sub BtnModificarCategoria_Click(sender As Object, e As EventArgs) Handles BtnModificarCategoria.Click
        Agregar_Categoria = 0
        If Modificar_Categorias = 1 Then
            Modificar_Categorias = 0
        ElseIf Modificar_Categorias = 0 Then
            Modificar_Categorias = 1
        End If
        HabilitarControlesCategorias()
    End Sub

    Private Sub HabilitarControlesCategorias()
        TxtBxNuevaCategoria.Clear()
        If Modificar_Categorias = 1 Then
            LblCategorias.Visible = True
            Nombre_Categoria.Visible = True
            LblNuevaCategoria.Text = "Nuevo nombre de la Categoria"
            LblNuevaCategoria.Location = New Point(15, 288)
            LblNuevaCategoria.Visible = True
            TxtBxNuevaCategoria.Location = New Point(15, 328)
            TxtBxNuevaCategoria.Visible = True
        ElseIf Agregar_Categoria = 1 Then
            LblNuevaCategoria.Text = "Nueva Categoria"
            LblCategorias.Visible = False
            Nombre_Categoria.Visible = False
            LblNuevaCategoria.Location = New Point(15, 207)
            LblNuevaCategoria.Visible = True
            TxtBxNuevaCategoria.Location = New Point(15, 240)
            TxtBxNuevaCategoria.Visible = True
        Else
            LblCategorias.Visible = True
            Nombre_Categoria.Visible = True
            LblNuevaCategoria.Visible = False
            TxtBxNuevaCategoria.Visible = False
        End If
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

    Dim Modificar_Equipo As Integer = 0
    Private Sub BtnModificarEquipo_Click(sender As Object, e As EventArgs) Handles BtnModificarEquipo.Click
        Agregar_Equipo = 0
        If Modificar_Equipo = 1 Then
            Modificar_Equipo = 0
        ElseIf Modificar_Equipo = 0 Then
            Modificar_Equipo = 1
        End If
        HabilitarControlesEquipo()
        Label60.Visible = True
        Anterior_Equipo.Visible = True
        Siguiente_Equipo.Visible = True
        Recorrer_Equipos()
    End Sub

    Dim Agregar_Equipo As Integer = 0
    Private Sub BtnNuevoEquipo_Click(sender As Object, e As EventArgs) Handles BtnNuevoEquipo.Click
        Modificar_Equipo = 0
        Agregar_Equipo = 1
        HabilitarControlesEquipo()
        Numero_Equipo.Clear()
        Numero_Equipo.Focus()
        Serie_Equipo.Clear()
        Nombre_Equipo.Clear()
        Marca_Equipo.Clear()
        Activo_Equipo.Checked = True
        Foto_Equipo.Image = My.Resources.NoMachine
        Label60.Visible = False
        Anterior_Equipo.Visible = False
        Siguiente_Equipo.Visible = False
    End Sub

    Private Sub HabilitarControlesEquipo()
        If Modificar_Equipo = 1 Or Agregar_Equipo = 1 Then
            Numero_Equipo.ReadOnly = False
            Serie_Equipo.ReadOnly = False
            Nombre_Equipo.ReadOnly = False
            Marca_Equipo.ReadOnly = False
            Activo_Equipo.Enabled = True
        Else
            Numero_Equipo.ReadOnly = True
            Serie_Equipo.ReadOnly = True
            Nombre_Equipo.ReadOnly = True
            Marca_Equipo.ReadOnly = True
            Activo_Equipo.Enabled = False
        End If
    End Sub

    Dim Modificar_Proveedor As Integer = 0
    Private Sub BtnModificarProveedor_Click(sender As Object, e As EventArgs) Handles BtnModificarProveedor.Click
        Agregar_Proveedor = 0
        If Modificar_Proveedor = 1 Then
            Modificar_Proveedor = 0
        ElseIf Modificar_Proveedor = 0 Then
            Modificar_Proveedor = 1
        End If
        HabilitarControlesProveedor()
        Label59.Visible = True
        Anterior_Proveedor.Visible = True
        Siguiente_Proveedor.Visible = True
        Recorrer_Proveedores()
    End Sub

    Dim Agregar_Proveedor As Integer = 0
    Private Sub BtnAgregarProveedor_Click(sender As Object, e As EventArgs) Handles BtnAgregarProveedor.Click
        Modificar_Proveedor = 0
        Agregar_Proveedor = 1
        HabilitarControlesProveedor()
        Nit_Proveedor.Clear()
        Nit_Proveedor.Focus()
        Nombre_Proveedor.Clear()
        Contacto_Proveedor.Clear()
        Direccion_Proveedor.Clear()
        Ciudad_Proveedor.Clear()
        Telefono_Proveedor.Clear()
        Email_Proveedor.Clear()
        Fax_Proveedor.Clear()
        Web_Proveedor.Clear()
        Detalle_Proveedor.Clear()
        Clasificacion_Proveedor.Clear()
        Activo_Proveedor.Checked = True
        Aprovado_Proveedor.Checked = True
        Label59.Visible = False
        Anterior_Proveedor.Visible = False
        Siguiente_Proveedor.Visible = False
    End Sub

    Private Sub HabilitarControlesProveedor()
        If Modificar_Proveedor = 1 Or Agregar_Proveedor = 1 Then
            Nit_Proveedor.ReadOnly = False
            Nombre_Proveedor.ReadOnly = False
            Contacto_Proveedor.ReadOnly = False
            Direccion_Proveedor.ReadOnly = False
            Ciudad_Proveedor.ReadOnly = False
            Telefono_Proveedor.ReadOnly = False
            Email_Proveedor.ReadOnly = False
            Fax_Proveedor.ReadOnly = False
            Web_Proveedor.ReadOnly = False
            Detalle_Proveedor.ReadOnly = False
            Clasificacion_Proveedor.ReadOnly = False
            Activo_Proveedor.Enabled = True
            Aprovado_Proveedor.Enabled = True
        Else
            Nit_Proveedor.ReadOnly = True
            Nombre_Proveedor.ReadOnly = True
            Contacto_Proveedor.ReadOnly = True
            Direccion_Proveedor.ReadOnly = True
            Ciudad_Proveedor.ReadOnly = True
            Telefono_Proveedor.ReadOnly = True
            Email_Proveedor.ReadOnly = True
            Fax_Proveedor.ReadOnly = True
            Web_Proveedor.ReadOnly = True
            Detalle_Proveedor.ReadOnly = True
            Clasificacion_Proveedor.ReadOnly = True
            Activo_Proveedor.Enabled = False
            Aprovado_Proveedor.Enabled = False
        End If
    End Sub


    Dim Modificar_Producto As Integer = 0
    Private Sub BtnModificarProducto_Click(sender As Object, e As EventArgs) Handles BtnModificarProducto.Click
        Agregar_Producto = 0
        If Modificar_Producto = 1 Then
            Modificar_Producto = 0
        ElseIf Modificar_Producto = 0 Then
            Modificar_Producto = 1
        End If
        HabilitarControlesProducto()
        Label72.Visible = True
        Anterior_Equipo.Visible = True
        Siguiente_Equipo.Visible = True
        Recorrer_Productos()
    End Sub

    Dim Agregar_Producto As Integer = 0

    Private Sub BtnNuevoProducto_Click(sender As Object, e As EventArgs) Handles BtnNuevoProducto.Click
        Modificar_Producto = 0
        Agregar_Producto = 1
        HabilitarControlesProducto()
        Codigo_Producto.Clear()
        Codigo_Producto.Focus()
        Serie_Producto.Clear()
        Nombre_Producto.Clear()
        Marca_Producto.Clear()
        Stock_Existente.Clear()
        Stock_Minimo.Clear()
        Stock_Maximo.Clear()
        Compra_Maxima.Clear()
        Activo_Producto.Checked = True
        Foto_Producto.Image = My.Resources.NoMachine
        Label60.Visible = False
        Anterior_Producto.Visible = False
        Siguiente_Producto.Visible = False
    End Sub

    Private Sub HabilitarControlesProducto()
        If Modificar_Producto = 1 Or Agregar_Producto = 1 Then
            Codigo_Producto.ReadOnly = False
            Serie_Producto.ReadOnly = False
            Nombre_Producto.ReadOnly = False
            Marca_Producto.ReadOnly = False
            Stock_Existente.ReadOnly = False
            Stock_Minimo.ReadOnly = False
            Stock_Maximo.ReadOnly = False
            Compra_Maxima.ReadOnly = False
            Activo_Producto.Enabled = True
            Categoria_Producto.Enabled = True
            SubCategoria_Producto.Enabled = True
            Unidades_Producto.Enabled = True
        Else
            Codigo_Producto.ReadOnly = True
            Serie_Producto.ReadOnly = True
            Nombre_Producto.ReadOnly = True
            Marca_Producto.ReadOnly = True
            Stock_Existente.ReadOnly = True
            Stock_Minimo.ReadOnly = True
            Stock_Maximo.ReadOnly = True
            Compra_Maxima.ReadOnly = True
            Activo_Producto.Enabled = False
            Categoria_Producto.Enabled = False
            SubCategoria_Producto.Enabled = False
            Unidades_Producto.Enabled = False
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
        TxtBxEmail.Clear()
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

    Dim Agregar_Categoria As Integer = 0
    Private Sub Nueva_Categoria_Click(sender As Object, e As EventArgs) Handles Nueva_Categoria.Click
        Modificar_Categorias = 0
        Agregar_Categoria = 1
        HabilitarControlesCategorias()
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
            TxtBxEmail.ReadOnly = False
            Perfiles_Usuario.Enabled = True
            Doag_Usuarios.Enabled = True
        Else
            Nombre_Usuario.ReadOnly = True
            Usuario_Nickname.ReadOnly = True
            TxtBxEmail.ReadOnly = True
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

    Private Sub TabPage10_Leave(sender As Object, e As EventArgs) Handles TabPage10.Leave
        Modificar_Categorias = 0
        Agregar_Categoria = 0
        HabilitarControlesCategorias()
    End Sub

    Private Sub TabPage13_Leave(sender As Object, e As EventArgs) Handles TabPage13.Leave
        Modificar_Equipo = 0
        Agregar_Equipo = 0
        HabilitarControlesEquipo()
    End Sub

    Private Sub TabPage14_Leave(Sender As Object, e As EventArgs) Handles TabPage14.Leave
        Modificar_Producto = 0
        Agregar_Producto = 0
        HabilitarControlesProducto()
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

    Private Sub Guardar_Categoria_Click(sender As Object, e As EventArgs) Handles Guardar_Categoria.Click
        If Modificar_Categorias = 1 Then
            Dim Nombre As String = TxtBxNuevaCategoria.Text.Trim
            Dim Id_cat As Integer = Nombre_Categoria.SelectedValue
            Try
                conn.Open()
                Dim query As String = "UPDATE categorias SET Nombre_Categoria = @Nombre WHERE Id_Categoria = @IdCat;"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Nombre", Nombre)
                    .AddWithValue("IdCat", Id_cat)
                End With
                cmd.ExecuteNonQuery()
                MsgBox("Categoria Modificada", MsgBoxStyle.Information, "Info.")
                conn.Close()
                CargarCBTabCategorias()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "ERROR.")
                conn.Close()
            End Try

        ElseIf Agregar_Categoria = 1 Then
            Dim nombre As String = TxtBxNuevaCategoria.Text.Trim
            Try
                conn.Open()
                Dim query As String = "INSERT INTO categorias(Nombre_Categoria) values (@Nombre);"
                Dim cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("Nombre", nombre)
                cmd.ExecuteNonQuery()
                MsgBox("Categoria Agregada", MsgBoxStyle.Information, "Info.")
                conn.Close()
                CargarCBTabCategorias()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try
            Agregar_Categoria = 0
            HabilitarControlesCategorias()
        End If
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
            Dim EmailUsu As String = TxtBxEmail.Text.Trim
            Try
                Dim em As New MailAddress(EmailUsu)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                Exit Sub
            End Try
            Try
                conn.Open()
                Dim query As String = "UPDATE usuarios SET Nombre_Usuario = @nombre, Usuario = @usuario
                                , Id_Perfil = @IDPerfil, Id_Doag = @doag, Email = @Email WHERE Id_Usuario = @IDUsu;"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("nombre", Nombre)
                    .AddWithValue("usuario", Usuario)
                    .AddWithValue("IDPerfil", IDPerfil)
                    .AddWithValue("doag", IDDoag)
                    .AddWithValue("IDUsu", Id_Usuario)
                    .AddWithValue("Email", EmailUsu)
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
            Dim EmailUsu As String = TxtBxEmail.Text.Trim
            Try
                Dim em As New MailAddress(EmailUsu)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                Exit Sub
            End Try
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
                Dim query As String = "INSERT into usuarios (Nombre_Usuario, Usuario, Salt, Hash, Id_Perfil, Id_Doag, Email)
                                      VALUES (@nombre, @usuario, @Salt, @Hash, @IDPerfil, @doag, @Email);"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("nombre", Nombre)
                    .AddWithValue("usuario", Usuario)
                    .AddWithValue("Salt", NewSalt)
                    .AddWithValue("Hash", Contrasena)
                    .AddWithValue("IDPerfil", IDPerfil)
                    .AddWithValue("doag", IDDoag)
                    .AddWithValue("IDUsu", Id_Usuario)
                    .AddWithValue("Email", EmailUsu)
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

    Private Sub BtnGuardarEquipo_Click(sender As Object, e As EventArgs) Handles BtnGuardarEquipo.Click
        If Modificar_Equipo = 1 Then
            Dim reader As MySqlDataReader
            Dim Numero As String = Numero_Equipo.Text.Trim
            Dim Serie As String = Serie_Equipo.Text.Trim
            Dim Nombre As String = Nombre_Equipo.Text.Trim
            Dim Marca As String = Marca_Equipo.Text.Trim
            Dim Activo As Boolean = Activo_Equipo.Checked

            Try
                conn.Open()
                Dim query As String = "UPDATE equipos SET Cod_Equipo = @Numero, Nombre_Equipo = @Nombre
                                , Marca = @Marca, Serie = @Serie, Activo = @Activo WHERE Id_Equipo= @IDEquipo;"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Numero", Numero)
                    .AddWithValue("Nombre", Nombre)
                    .AddWithValue("Marca", Marca)
                    .AddWithValue("Serie", Serie)
                    .AddWithValue("Activo", Activo)
                    .AddWithValue("IDEquipo", Id_Equipo)
                End With
                reader = cmd.ExecuteReader
                MsgBox("Equipo modificado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try

        ElseIf Agregar_Equipo = 1 Then

            Dim reader As MySqlDataReader
            Dim Numero As String = Numero_Equipo.Text.Trim
            Dim Serie As String = Serie_Equipo.Text.Trim
            Dim Nombre As String = Nombre_Equipo.Text.Trim
            Dim Marca As String = Marca_Equipo.Text.Trim
            Dim Activo As Boolean = Activo_Equipo.Checked

            If Nombre = "" Or Numero = "" Or Serie = "" Or Marca = "" Then
                MsgBox("Todos los campos son obligatorios", MsgBoxStyle.Exclamation, "Error")
                Exit Sub
            End If

            Try
                conn.Open()
                Dim query As String = "INSERT into Equipos (Cod_Equipo, Nombre_Equipo, Marca, Serie, Activo)
                                      VALUES (@Numero, @Nombre, @Marca, @Serie, @Activo);"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Numero", Numero)
                    .AddWithValue("Nombre", Nombre)
                    .AddWithValue("Marca", Marca)
                    .AddWithValue("Serie", Serie)
                    .AddWithValue("Activo", Activo)
                End With
                reader = cmd.ExecuteReader
                MsgBox("Equipo Agregado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try
            Agregar_Equipo = 0
            HabilitarControlesEquipo()
            Recorrer_Equipos()
        End If
    End Sub

    Private Sub BtnGuardarProducto_Click(sender As Object, e As EventArgs) Handles BtnGuardarProducto.Click
        If Codigo_Producto.Text = "" Or Nombre_Producto.Text = "" Or Marca_Producto.Text = "" Then
            MsgBox("Los campos que tengan (*) son obligatorios")
            Exit Sub
        End If
        Dim var = Unidades_Producto.SelectedItem
        If Unidades_Producto.SelectedItem = Nothing Then
            MsgBox("Escoja una unidad para el producto", MsgBoxStyle.Exclamation, "Error.")
            Exit Sub
        End If

        If Convert.ToInt32(Stock_Minimo.Text) > Convert.ToInt32(Stock_Maximo.Text) Then
            MsgBox("El Stock minimo no puede ser mayor que el Stock maximo", MsgBoxStyle.Exclamation, "Error")
            Exit Sub
        End If

        If Convert.ToInt32(Compra_Maxima.Text) > Convert.ToInt32(Stock_Maximo.Text) Then
            MsgBox("El valor de la compra maxima no puede ser mayor que el stock maximo", MsgBoxStyle.Exclamation, "Error")
            Exit Sub
        End If

        If Stock_Existente.Text = "" Then
            Stock_Existente.Text = 0
        End If
        If Stock_Minimo.Text = "" Then
            Stock_Minimo.Text = 0
        End If
        If Stock_Maximo.Text = "" Then
            Stock_Maximo.Text = 0
        End If
        If Compra_Maxima.Text = "" Then
            Compra_Maxima.Text = 0
        End If

        If Modificar_Producto = 1 Then
            Dim reader As MySqlDataReader
            Dim Codigo As String = Codigo_Producto.Text.Trim
            Dim Serie As String = Serie_Producto.Text.Trim
            Dim Nombre As String = Nombre_Producto.Text.Trim
            Dim Marca As String = Marca_Producto.Text.Trim
            Dim Activo As Boolean = Activo_Producto.Checked
            Dim IDCategoria As String = Categoria_Producto.SelectedValue.ToString
            Dim IDSubcategoria As String = SubCategoria_Producto.SelectedValue.ToString

            Try
                conn.Open()
                Dim query As String = "UPDATE productos SET Cod_Producto = @Codigo, Nombre_Producto = @Nombre
                                , Marca = @Marca, Serie = @Serie,ID_Categoria = @Categoria, Id_Subcategoria = @Subcategoria,
                                Activo = @Activo WHERE Id_Producto= @IDProducto;"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Codigo", Codigo)
                    .AddWithValue("Nombre", Nombre)
                    .AddWithValue("Marca", Marca)
                    .AddWithValue("Serie", Serie)
                    .AddWithValue("Categoria", IDCategoria)
                    .AddWithValue("Subcategoria", IDSubcategoria)
                    .AddWithValue("Activo", Activo)
                    .AddWithValue("IDProducto", Id_Prod)
                End With
                reader = cmd.ExecuteReader
                MsgBox("Producto modificado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try

            Dim StockMinimo As String = Stock_Minimo.Text.Trim
            Dim StockMaximo As String = Stock_Maximo.Text.Trim
            Dim StockExistente As String = Stock_Existente.Text.Trim
            Dim Unidades As String = Unidades_Producto.SelectedItem.ToString
            Dim CompraMaxima As String = Compra_Maxima.Text.Trim

            Try
                conn.Open()
                Dim query As String = "UPDATE productos SET Stock_Minimo = @Minimo, Stock_Maximo = @Maximo, Stock_Existente = @Existente,
                                       Unidades = @Unidades, Compra_Maxima = @CMaxima WHERE Id_Producto = @IdProducto"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Minimo", StockMinimo)
                    .AddWithValue("Maximo", StockMaximo)
                    .AddWithValue("Existente", StockExistente)
                    .AddWithValue("Unidades", Unidades)
                    .AddWithValue("CMaxima", CompraMaxima)
                    .AddWithValue("IdProducto", Id_Prod)
                End With
                cmd.ExecuteScalar()
                MsgBox("Stock Actualizado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox("Error Actualizando Stock del Producto:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try

        ElseIf Agregar_Producto = 1 Then

            Dim reader As MySqlDataReader
            Dim Codigo As String = Codigo_Producto.Text.Trim
            Dim Serie As String = Serie_Producto.Text.Trim
            Dim Nombre As String = Nombre_Producto.Text.Trim
            Dim Marca As String = Marca_Producto.Text.Trim
            Dim Activo As Boolean = Activo_Producto.Checked
            Dim IDCategoria As String = Categoria_Producto.SelectedValue.ToString
            Dim IDSubcategoria As String = SubCategoria_Producto.SelectedValue.ToString
            Dim StockMinimo As String = Stock_Minimo.Text.Trim
            Dim StockMaximo As String = Stock_Maximo.Text.Trim
            Dim StockExistente As String = Stock_Existente.Text.Trim
            Dim Unidades As String = Unidades_Producto.SelectedItem.ToString
            Dim CompraMaxima As String = Compra_Maxima.Text.Trim

            Try
                conn.Open()
                Dim query As String = "INSERT into Productos (Cod_Producto, Nombre_Producto, Marca, Serie, Id_Categoria, Id_SubCategoria, Stock_Minimo, Stock_Maximo, Stock_Existente, Unidades, Compra_Maxima, Activo)
                                      VALUES (@Codigo, @Nombre, @Marca, @Serie, @Categoria, @Subcategoria, @Minimo, @Maximo, @Existente, @Unidades, @CMaxima, @Activo);"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Codigo", Codigo)
                    .AddWithValue("Nombre", Nombre)
                    .AddWithValue("Marca", Marca)
                    .AddWithValue("Serie", Serie)
                    .AddWithValue("Categoria", IDCategoria)
                    .AddWithValue("Subcategoria", IDSubcategoria)
                    .AddWithValue("Minimo", StockMinimo)
                    .AddWithValue("Maximo", StockMaximo)
                    .AddWithValue("Existente", StockExistente)
                    .AddWithValue("Unidades", Unidades)
                    .AddWithValue("CMaxima", CompraMaxima)
                    .AddWithValue("Activo", Activo)
                End With
                reader = cmd.ExecuteReader
                MsgBox("Producto Agregado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try

            Agregar_Producto = 0
            HabilitarControlesProducto()
            Recorrer_Productos()
        End If
    End Sub

    Private Function IsValidWebFormat(ByVal s As String) As Boolean
        Return (Regex.IsMatch(s, "((((http[s]?|ftp)[:]//)([a-zA-Z0-9.-]+([:][a-zA-Z0-9.&amp;%$-]+)*@)?[a-zA-Z][a-zA-Z0-9.-]+|[a-zA-Z][a-zA-Z0-9]+[.][a-zA-Z][a-zA-Z0-9.-]+)[.](com|edu|gov|mil|net|org|biz|pro|info|name|museum|ac|ad|ae|af|ag|ai|al|am|an|ao|aq|ar|as|at|au|aw|az|ax|ba|bb|bd|be|bf|bg|bh|bi|bj|bm|bn|bo|br|bs|bt|bv|bw|by|bz|ca|cc|cd|cf|cg|ch|ci|ck|cl|cm|cn|co|cr|cs|cu|cv|cx|cy|cz|de|dj|dk|dm|do|dz|ec|ee|eg|eh|er|es|et|eu|fi|fj|fk|fm|fo|fr|ga|gb|gd|ge|gf|gg|gh|gi|gl|gm|gn|gp|gq|gr|gs|gt|gu|gw|hk|hm|hn|hr|ht|hu|id|ie|il|im|in|io|iq|ir|is|it|je|jm|jo|jp|ke|kg|kh|ki|km|kn|kp|kr|kw|ky|kz|la|lb|lc|li|lk|lr|ls|lt|lu|lv|ly|ma|mc|md|mg|mh|mk|ml|mm|mn|mo|mp|mq|mr|ms|mt|mu|mv|mw|mx|my|mz|na|nc|ne|nf|ng|ni|nl|no|np|nr|nu|nz|om|pa|pe|pf|pg|ph|pk|pl|pm|pn|pr|ps|pt|pw|py|qa|re|ro|ru|rw|sa|sb|sc|sd|se|sg|sh|si|sj|sk|sl|sm|sn|so|sr|st|sv|sy|sz|tc|td|tf|tg|th|tj|tk|tl|tn|to|tp|tr|tt|tv|tw|tz|ua|ug|uk|um|us|uy|uz|va|vc|ve|vg|vi|vn|vu|wf|ws|ye|yt|yu|za|zm|zw)([:][0-9]+)*(/[a-zA-Z0-9.,;?'\\+&amp;%$#=~_-]+)*)"))
    End Function


    Private Sub BtnGuardarProveedor_Click(sender As Object, e As EventArgs) Handles BtnGuardarProveedor.Click
        If Modificar_Proveedor = 1 Then
            Dim reader As MySqlDataReader
            Dim Nit As String = Nit_Proveedor.Text.Trim
            Dim Nombre As String = Nombre_Proveedor.Text.Trim
            Dim Contacto As String = Contacto_Proveedor.Text.Trim
            Dim Direccion As String = Direccion_Proveedor.Text.Trim
            Dim Ciudad As String = Ciudad_Proveedor.Text.Trim
            Dim Telefono As String = Telefono_Proveedor.Text.Trim
            Dim Email As String = Email_Proveedor.Text.Trim
            Try
                Dim em As New MailAddress(Email)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                Exit Sub
            End Try
            Dim Fax As String = Fax_Proveedor.Text.Trim
            Dim Web As String = Web_Proveedor.Text.Trim
            If IsValidWebFormat(LCase(Web)) = False Then
                MsgBox("El formato de la pagina web no es valido", MsgBoxStyle.Exclamation, "Error")
                Exit Sub
            End If
            Dim Detalle As String = Detalle_Proveedor.Text.Trim
            Dim Clasificacion As String = Clasificacion_Proveedor.Text.Trim
            Dim Aprovado As Boolean = Aprovado_Proveedor.Checked
            Dim Activo As Boolean = Activo_Proveedor.Checked

            Try
                conn.Open()
                Dim query As String = "UPDATE proveedores SET Nit_Proveedor = @NIT, Nombre_Proveedor = @Nombre,
                                Nombre_Contacto = @Contacto, Direccion = @Direccion, Ciudad = @Ciudad, Numero_Telefono = @Telefono,
                                Email_Contacto = @Email, Numero_Fax = @Fax, Pagina_Web = @Web, Detalle = @Detalle, 
                                Clasificacion_OIMS = @Clasificacion, Aprovado = @Aprovado, Activo = @Activo 
                                WHERE Nit_Proveedor = @Id_Prov"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("NIT", Nit)
                    .AddWithValue("Nombre", Nombre)
                    .AddWithValue("Contacto", Contacto)
                    .AddWithValue("Direccion", Direccion)
                    .AddWithValue("Ciudad", Ciudad)
                    .AddWithValue("Telefono", Telefono)
                    .AddWithValue("Email", Email)
                    .AddWithValue("Fax", Fax)
                    .AddWithValue("Web", Web)
                    .AddWithValue("Detalle", Detalle)
                    .AddWithValue("Clasificacion", Clasificacion)
                    .AddWithValue("Aprovado", Aprovado)
                    .AddWithValue("Activo", Activo)
                    .AddWithValue("Id_Prov", ID_Prov)
                End With
                reader = cmd.ExecuteReader
                MsgBox("Proveedor modificado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try
        ElseIf Agregar_Proveedor = 1 Then

            Dim reader As MySqlDataReader
            Dim Nit As String = Nit_Proveedor.Text.Trim
            Dim Nombre As String = Nombre_Proveedor.Text.Trim
            Dim Contacto As String = Contacto_Proveedor.Text.Trim
            Dim Direccion As String = Direccion_Proveedor.Text.Trim
            Dim Ciudad As String = Ciudad_Proveedor.Text.Trim
            Dim Telefono As String = Telefono_Proveedor.Text.Trim
            Dim Email As String = Email_Proveedor.Text.Trim
            Try
                Dim em As New MailAddress(Email)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                Exit Sub
            End Try
            Dim Fax As String = Fax_Proveedor.Text.Trim
            Dim Web As String = Web_Proveedor.Text.Trim
            If IsValidWebFormat(LCase(Web)) = False Then
                MsgBox("El formato de la pagina web no es valido", MsgBoxStyle.Exclamation, "Error")
                Exit Sub
            End If
            Dim Detalle As String = Detalle_Proveedor.Text.Trim
            Dim Clasificacion As String = Clasificacion_Proveedor.Text.Trim
            Dim Aprovado As Boolean = Aprovado_Proveedor.Checked
            Dim Activo As Boolean = Activo_Proveedor.Checked

            Try
                conn.Open()
                Dim query As String = "INSERT INTO proveedores (Nit_Proveedor, Nombre_Proveedor, Nombre_Contacto, Direccion, Ciudad, Numero_Telefono, Email_Contacto, Numero_Fax, Pagina_Web, Detalle, Clasificacion_OIMS, Aprovado,  Activo)
                                      VALUES (@NIT, @Nombre, @Contacto, @Direccion, @Ciudad, @Telefono, @Email, @Fax, @Web, @Detalle, @Clasificacion, @Aprovado, @Activo);"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("NIT", Nit)
                    .AddWithValue("Nombre", Nombre)
                    .AddWithValue("Contacto", Contacto)
                    .AddWithValue("Direccion", Direccion)
                    .AddWithValue("Ciudad", Ciudad)
                    .AddWithValue("Telefono", Telefono)
                    .AddWithValue("Email", Email)
                    .AddWithValue("Fax", Fax)
                    .AddWithValue("Web", Web)
                    .AddWithValue("Detalle", Detalle)
                    .AddWithValue("Clasificacion", Clasificacion)
                    .AddWithValue("Aprovado", Aprovado)
                    .AddWithValue("Activo", Activo)
                End With
                reader = cmd.ExecuteReader
                MsgBox("Proveedor Agregado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try

            Agregar_Proveedor = 0
            HabilitarControlesProveedor()
            Recorrer_Proveedores()
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
    Private Sub BtnEliminarEquipo_Click(sender As Object, e As EventArgs) Handles BtnEliminarEquipo.Click
        If Agregar_Equipo = 1 Then
            Exit Sub
        End If
        If MessageBox.Show("¿Esta seguro que desea ELIMINAR este Equipo?", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Try
                conn.Open()
                Dim query As String = "Delete from Equipos where Id_Equipo = @IdEquipo;"
                Dim cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("IdEquipo", Id_Equipo)
                cmd.ExecuteNonQuery()
                MsgBox("Equipo Eliminado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End If
        Recorrer_Equipos()
    End Sub

    Private Sub BtnEliminarProducto_Click(sender As Object, e As EventArgs) Handles BtnEliminarProducto.Click
        If Agregar_Producto = 1 Then
            Exit Sub
        End If
        If MessageBox.Show("¿Esta seguro que desea ELIMINAR este Producto?", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Try
                conn.Open()
                Dim query As String = "Delete from Productos where Id_Producto = @IdProducto;"
                Dim cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("IdProducto", Id_Prod)
                cmd.ExecuteNonQuery()
                MsgBox("Producto Eliminado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End If
        Recorrer_Productos()
    End Sub

    Private Sub BtnEliminarProveedor_Click(sender As Object, e As EventArgs) Handles BtnEliminarProveedor.Click
        If Agregar_Proveedor = 1 Then
            Exit Sub
        End If
        If MessageBox.Show("¿Esta seguro que desea ELIMINAR este Proveedor?", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Try
                conn.Open()
                Dim query As String = "Delete from Proveedores where Nit_Proveedor = @NIT;"
                Dim cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("NIT", ID_Prov)
                cmd.ExecuteNonQuery()
                MsgBox("Proveedor Eliminado", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End If
        Recorrer_Proveedores()
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

    Private Sub Eliminar_Categoria_Click(sender As Object, e As EventArgs) Handles Eliminar_Categoria.Click
        If Agregar_Categoria = 1 Then
            Exit Sub
        End If
        Dim Id_Cat As Integer = Nombre_Categoria.SelectedValue
        If MessageBox.Show("¿Esta seguro que desea ELIMINAR esta Categoria?", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Try
                conn.Open()
                Dim query As String = "Delete from categorias where Id_Categoria = @IdCat;"
                Dim cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("IdCat", Id_Cat)
                cmd.ExecuteNonQuery()
                MsgBox("Categoria Eliminada", MsgBoxStyle.Information, "Info.")
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End If
        CargarCBTabCategorias()
    End Sub


    Private Sub Eliminar_SubCategoria_Click(sender As Object, e As EventArgs) Handles Eliminar_SubCategoria.Click
        Dim fila As Integer
        If DataGridView2.CurrentRow.Index.ToString <> Nothing Then
            fila = DataGridView2.CurrentRow.Index
        Else
            fila = 0
        End If
        Try
            conn.Open()
            Dim query As String = "DELETE from categorias_sub WHERE Id_SubCategoria = @IdSubCat;"
            Dim cmd As New MySqlCommand(query, conn)
            cmd.Parameters.AddWithValue("IdSubCat", DataGridView2.Item(0, fila).FormattedValue)
            cmd.ExecuteNonQuery()
            MsgBox("Subcategoria Eliminada", MsgBoxStyle.Information, "Info.")
            conn.Close()
        Catch ex As Exception
            MsgBox("Error durante la eliminacion: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "ERROR.")
            conn.Close()
        End Try
        CargarCBTabCategorias()
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

    Private Sub Equipos_Consultar_Click(sender As Object, e As EventArgs) Handles Equipos_Consultar.Click
        If Buscar_Eq.Text.Trim = "" Then
            Exit Sub
        End If
        Cargar_Tabla("*", "EQUIPOS")
        If z <> Buscar_Eq.Text Then
            cant_reg_encon = 0
        End If
        Try
            conn.Open()
            Dim consulta As String = "Select * from Equipos"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            conn.Close()
            Dim i As Integer = 0
            Dim foundRows() As Data.DataRow
            foundRows = MysqlDset.Tables(0).Select("Nombre_Equipo Like '" & Buscar_Eq.Text & "%'")
            z = Buscar_Eq.Text
            If cant_reg_encon = 0 And foundRows.Length > 1 Then
                cant_reg_encon = foundRows.Length
                For Each row In Tabla1.Rows
                    If foundRows(cant_reg_encon - 1).Item(1) = row(1) Then
                        'MsgBox(foundRows(cant_reg_encon - 1).Item(1))
                        Equipo_num = i + 1
                        Recorrer_Equipos()
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
                            Equipo_num = i + 1
                            Recorrer_Equipos()
                            Exit Sub
                        End If
                        i = i + 1
                    Next
                Else
                    For Each row In Tabla1.Rows
                        If foundRows(cant_reg_encon - 1).Item(1) = row(1) Then
                            Equipo_num = i + 1
                            Recorrer_Equipos()
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

    Private Sub Productos_Consultar_Click(sender As Object, e As EventArgs) Handles Productos_Consultar.Click
        If Buscar_Produ.Text.Trim = "" Then
            Exit Sub
        End If
        Cargar_Tabla("*", "PRODUCTOS")
        If z <> Buscar_Produ.Text Then
            cant_reg_encon = 0
        End If
        Try
            conn.Open()
            Dim consulta As String = "Select * from productos"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            conn.Close()
            Dim i As Integer = 0
            Dim foundRows() As Data.DataRow
            foundRows = MysqlDset.Tables(0).Select("Nombre_Producto Like '" & Buscar_Produ.Text & "%'")
            z = Buscar_Produ.Text
            If cant_reg_encon = 0 And foundRows.Length > 1 Then
                cant_reg_encon = foundRows.Length
                For Each row In Tabla1.Rows
                    If foundRows(cant_reg_encon - 1).Item(1) = row(1) Then
                        'MsgBox(foundRows(cant_reg_encon - 1).Item(1))
                        Prod_Num = i + 1
                        Recorrer_Productos()
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
                            Prod_Num = i + 1
                            Recorrer_Productos()
                            Exit Sub
                        End If
                        i = i + 1
                    Next
                Else
                    For Each row In Tabla1.Rows
                        If foundRows(cant_reg_encon - 1).Item(1) = row(1) Then
                            Prod_Num = i + 1
                            Recorrer_Productos()
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

    Private Sub Proveedores_Consultar_Click(sender As Object, e As EventArgs) Handles Proveedores_Consultar.Click
        If Buscar_Prov.Text.Trim = "" Then
            Exit Sub
        End If
        Cargar_Tabla("*", "PROVEEDORES")
        If z <> Buscar_Prov.Text Then
            cant_reg_encon = 0
        End If
        Try
            conn.Open()
            Dim consulta As String = "Select * from proveedores"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            conn.Close()
            Dim i As Integer = 0
            Dim foundRows() As Data.DataRow
            foundRows = MysqlDset.Tables(0).Select("Nombre_Proveedor Like '" & Buscar_Prov.Text & "%'")
            z = Buscar_Prov.Text
            If cant_reg_encon = 0 And foundRows.Length > 1 Then
                cant_reg_encon = foundRows.Length
                For Each row In Tabla1.Rows
                    If foundRows(cant_reg_encon - 1).Item(1) = row(1) Then
                        'MsgBox(foundRows(cant_reg_encon - 1).Item(1))
                        Proveedor_Num = i + 1
                        Recorrer_Proveedores()
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
                            Proveedor_Num = i + 1
                            Recorrer_Proveedores()
                            Exit Sub
                        End If
                        i = i + 1
                    Next
                Else
                    For Each row In Tabla1.Rows
                        If foundRows(cant_reg_encon - 1).Item(1) = row(1) Then
                            Proveedor_Num = i + 1
                            Recorrer_Proveedores()
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

    Private Sub Buscar_Eq_KeyDown(sender As Object, e As KeyEventArgs) Handles Buscar_Eq.KeyDown
        If e.KeyCode = Keys.Enter Then
            Equipos_Consultar_Click(Me.Equipos_Consultar, Nothing)
        End If
    End Sub

    Private Sub Buscar_Produ_KeyDown(sender As Object, e As KeyEventArgs) Handles Buscar_Produ.KeyDown
        If e.KeyCode = Keys.Enter Then
            Productos_Consultar_Click(Me.Productos_Consultar, Nothing)
        End If
    End Sub

    Private Sub TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Monto_Doag.KeyPress, Stock_Existente.KeyPress, Stock_Maximo.KeyPress, Stock_Minimo.KeyPress, Compra_Maxima.KeyPress, Cantidad_Movimiento.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub TextBoxPhone_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles Telefono_Proveedor.KeyPress, Fax_Proveedor.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                If Asc(e.KeyChar) = 40 Or Asc(e.KeyChar) = 41 Or Asc(e.KeyChar) = 43 Or Asc(e.KeyChar) = 45 Then
                    Exit Sub
                End If
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

    Private Sub Guardar_SubCategoria_Click(sender As Object, e As EventArgs) Handles Guardar_SubCategoria.Click
        Dim fila As Integer = (DataGridView2.Rows.Count - 2)
        For i = 0 To fila
            Dim NombreSubCat As String = DataGridView2.Item(1, i).FormattedValue
            Dim IdCat As Integer = Nombre_Categoria.SelectedValue
            If DataGridView2.Item(0, i).FormattedValue = "" Then
                Try
                    conn.Open()
                    Dim cmd As New MySqlCommand("INSERT INTO categorias_sub (Nombre_SubCategoria, Id_Categoria) " &
                            "VALUES (@NombreSubCat, @IdCat)", conn)
                    With cmd.Parameters
                        .AddWithValue("NombreSubCat", NombreSubCat)
                        .AddWithValue("IdCat", IdCat)
                    End With
                    cmd.ExecuteNonQuery()
                    'MessageBox.Show("Registro AGREGADO")
                    conn.Close()
                Catch ex As Exception
                    MsgBox("El registro no pudo Agregarse por: " & vbCrLf & ex.Message)
                End Try
            Else
                Dim Id_SubCat As Integer = DataGridView2.Item(0, i).FormattedValue
                Dim var As String = DataGridView2.Item(1, i).FormattedValue
                Try
                    conn.Open()
                    Dim cmd As New MySqlCommand("UPDATE categorias_sub SET Nombre_SubCategoria = @NombreSubCat WHERE Id_SubCategoria = @Id_SubCat", conn)
                    With cmd.Parameters
                        .AddWithValue("NombreSubCat", NombreSubCat)
                        .AddWithValue("Id_SubCat", Id_SubCat)
                    End With
                    cmd.ExecuteNonQuery()
                    'MessageBox.Show("Registro MODIFICADO")
                    conn.Close()
                Catch ex As Exception
                    MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
                End Try
            End If

        Next
        MessageBox.Show("SubCategorias Actualizadas")
        CargarCBTabCategorias()
    End Sub

    Private Sub Agregar_Estante_Click(sender As Object, e As EventArgs) Handles Agregar_Estante.Click
        Dim NumeroEstantes As Integer = 0
        Dim IDDatosApp As String = ""
        Try
            conn.Open()
            Dim consulta As String = "Select * from datos_app"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            NumeroEstantes = Convert.ToInt32(MysqlDset.Tables(0).Rows(0).Item(2))
            IDDatosApp = MysqlDset.Tables(0).Rows(0).Item(0)
            Dim cmd As New MySqlCommand("UPDATE datos_app SET Detalles = '" & (NumeroEstantes + 1) & "' " &
                        "WHERE IdDatos_App = '" & IDDatosApp & "'", conn)
            cmd.ExecuteNonQuery()
            conn.Close()
        Catch ex As Exception
            MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
            conn.Close()
        End Try
        CargarCBTabUbicaciones()
    End Sub

    Private Sub Eliminar_Estante_Click(sender As Object, e As EventArgs) Handles Eliminar_Estante.Click
        Dim NumeroEstantes As Integer = 0
        Dim IDDatosApp As String = ""
        Try
            conn.Open()
            Dim consulta As String = "Select * from datos_app"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            conn.Close()
            NumeroEstantes = Convert.ToInt32(MysqlDset.Tables(0).Rows(0).Item(2))
            IDDatosApp = MysqlDset.Tables(0).Rows(0).Item(0)
            '.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
        Dim Num_men As Integer = 0
        For Each row As DataGridViewRow In Me.DataGridView3.Rows
            'obtenemos el valor de la columna en la variable declarada
            If Convert.ToInt32(row.Cells(1).Value) > Num_men Then
                Num_men = row.Cells(1).Value 'donde (0) es la columna a recorrer
            End If
        Next
        If Num_men < NumeroEstantes Then
            Try

                conn.Open()
                    Dim cmd As New MySqlCommand("UPDATE datos_app SET Detalles = '" & (NumeroEstantes - 1) & "' " &
                                "WHERE IdDatos_App = '" & IDDatosApp & "'", conn)
                    cmd.ExecuteNonQuery()
                'MessageBox.Show("Registro MODIFICADO")
                conn.Close()
            Catch ex As Exception
                MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
                conn.Close()
            End Try
        Else
            MsgBox("Su numero minimo de estantes: " & NumeroEstantes)
        End If
        CargarCBTabUbicaciones()
    End Sub

    Private Sub Agregar_Entrepano_Click(sender As Object, e As EventArgs) Handles Agregar_Entrepano.Click
        Dim a As Integer = 0
        Dim b As String = ""
        Try
            conn.Open()
            Dim consulta As String = "Select * from datos_app"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            a = Convert.ToInt64(MysqlDset.Tables(0).Rows(1).Item(2))
            b = MysqlDset.Tables(0).Rows(1).Item(0)
            Dim cmd As New MySqlCommand("UPDATE datos_app SET Detalles = '" & (a + 1) & "' " &
                            "WHERE IdDatos_App = '" & b & "'", conn)
            cmd.ExecuteNonQuery()
            'MessageBox.Show("Registro MODIFICADO")
            conn.Close()
        Catch ex As Exception
            MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
            conn.Close()
        End Try
        CargarCBTabUbicaciones()
    End Sub

    Private Sub Eliminar_Entrepano_Click(sender As Object, e As EventArgs) Handles Eliminar_Entrepano.Click
        Dim a As Integer = 0
        Dim b As String = ""
        Try
            conn.Open()
            Dim consulta As String = "Select * from datos_app"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            conn.Close()
            a = Convert.ToInt64(MysqlDset.Tables(0).Rows(1).Item(2))
            b = MysqlDset.Tables(0).Rows(1).Item(0)
            '.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
        Dim Num_men As Integer = 0
        For Each row As DataGridViewRow In Me.DataGridView3.Rows
            'obtenemos el valor de la columna en la variable declarada
            If Convert.ToInt64(row.Cells(2).Value) > Num_men Then
                Num_men = row.Cells(2).Value 'donde (0) es la columna a recorrer
            End If
        Next
        If Num_men < a Then
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("UPDATE datos_app SET Detalles = '" & (a - 1) & "' " &
                            "WHERE IdDatos_App = '" & b & "'", conn)
                cmd.ExecuteNonQuery()
                'MessageBox.Show("Registro MODIFICADO")
                conn.Close()
            Catch ex As Exception
                MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
                conn.Close()
            End Try
        Else
            MsgBox("Su numero minimo de entrepaños es: " & a)
        End If
        CargarCBTabUbicaciones()
    End Sub

    Private Sub Agregar_Caja_Click(sender As Object, e As EventArgs) Handles Agregar_Caja.Click
        Dim a As Integer = 0
        Dim b As String = ""
        Try
            conn.Open()
            Dim consulta As String = "Select * from datos_app"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            a = Convert.ToInt64(MysqlDset.Tables(0).Rows(2).Item(2))
            b = MysqlDset.Tables(0).Rows(2).Item(0)
            Dim cmd As New MySqlCommand("UPDATE datos_app SET Detalles = '" & (a + 1) & "' " &
                            "WHERE IdDatos_App = '" & b & "'", conn)
            cmd.ExecuteNonQuery()
            'MessageBox.Show("Registro MODIFICADO")
            conn.Close()
        Catch ex As Exception
            MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
            conn.Close()
        End Try
        CargarCBTabUbicaciones()
    End Sub

    Private Sub Eliminar_Caja_Click(sender As Object, e As EventArgs) Handles Eliminar_Caja.Click
        Dim a As Integer = 0
        Dim b As String = ""
        Try
            conn.Open()
            Dim consulta As String = "Select * from datos_app"
            Dim MysqlDadap As New MySqlDataAdapter(consulta, conn)
            Dim MysqlDset As New DataSet
            MysqlDadap.Fill(MysqlDset)
            conn.Close()
            a = Convert.ToInt64(MysqlDset.Tables(0).Rows(2).Item(2))
            b = MysqlDset.Tables(0).Rows(2).Item(0)
            '.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            conn.Close()
        End Try
        Dim Num_men As Integer = 0
        For Each row As DataGridViewRow In Me.DataGridView3.Rows
            'obtenemos el valor de la columna en la variable declarada
            If row.Cells(3).Value = "Azul" Or row.Cells(3).Value = "Rojo" Or row.Cells(3).Value = "Amarillo" Or row.Cells(3).Value = "Blanco" Then
            Else
                If Convert.ToInt64(row.Cells(3).Value) > Num_men Then
                    Num_men = row.Cells(3).Value 'donde (0) es la columna a recorrer
                End If
            End If

        Next
        If Num_men < a Then
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("UPDATE datos_app SET Detalles = '" & (a - 1) & "' " &
                            "WHERE IdDatos_App = '" & b & "'", conn)
                cmd.ExecuteNonQuery()
                'MessageBox.Show("Registro MODIFICADO")
                conn.Close()
            Catch ex As Exception
                MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
                conn.Close()
            End Try
        Else
            MsgBox("Su numero minimo de cajas es: " & a)
        End If
        CargarCBTabUbicaciones()
    End Sub

    Dim IDUbicacion As Integer
    Private Sub DataGridView3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView3.SelectionChanged
        Try
            Dim fila_actual As Integer = (DataGridView3.CurrentRow.Index)
            If fila_actual = (DataGridView1.Rows.Count - 1) Then
                IDUbicacion = -1
                Nombre_Doag.Text = ""
                Monto_Doag.Text = ""
                Comentario_Doag.Text = ""
            Else
                Cargar_Tabla("*", "Ubicaciones")
                IDUbicacion = Tabla1.Rows(fila_actual).ItemArray(0)
                Estantes.Text = Tabla1.Rows(fila_actual).ItemArray(1).ToString
                Entrepanos.Text = Tabla1.Rows(fila_actual).ItemArray(2).ToString
                Cajas_Colores.Text = Tabla1.Rows(fila_actual).ItemArray(3).ToString
                Zonas.Text = Tabla1.Rows(fila_actual).ItemArray(4).ToString
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Guardar_Ubicacion_Click(sender As Object, e As EventArgs) Handles Guardar_Ubicacion.Click
        Cargar_Tabla("*", "ubicaciones")
        If Estantes.SelectedIndex = (-1) Or Entrepanos.SelectedIndex = (-1) Or Cajas_Colores.SelectedIndex = (-1) Or Zonas.SelectedIndex = (-1) Then
            MsgBox("Ningun campo puede ser vacio")
            Exit Sub
        End If
        Dim a As Integer
        Dim consulta As New MySqlCommand(“select * from ubicaciones where Id_Ubicacion=@IdUbicacion;", conn)
        consulta.Parameters.AddWithValue("IdUbicacion", IDUbicacion)
        conn.Open()
        Dim leerbd As MySqlDataReader = consulta.ExecuteReader()
        If leerbd.Read <> False Then
            leerbd.Close()
            Try
                Dim query As String = "UPDATE ubicaciones SET Estante = @Estante, Entrepano = @Entrepano, Caja_Color = @Caja, Zona = @Zona
                                       where Id_ubicacion = @IdUbicacion;"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Estante", Estantes.Text.Trim)
                    .AddWithValue("Entrepano", Entrepanos.Text.Trim)
                    .AddWithValue("Caja", Cajas_Colores.Text.Trim)
                    .AddWithValue("Zona", Zonas.Text.Trim)
                    .AddWithValue("IdUbicacion", IDUbicacion)
                End With
                cmd.ExecuteNonQuery()
                a = IDUbicacion
                conn.Close()
                MsgBox("Registro Modificado", MsgBoxStyle.Information, "Info.")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try
        Else
            leerbd.Close()
            Try
                Dim query As String = "INSERT INTO ubicaciones (Estante, Entrepano, Caja_Color, Zona)
                                       VALUES(@Estante, @Entrepano, @Caja, @Zona);"
                Dim cmd As New MySqlCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("Estante", Estantes.Text.Trim)
                    .AddWithValue("Entrepano", Entrepanos.Text.Trim)
                    .AddWithValue("Caja", Cajas_Colores.Text.Trim)
                    .AddWithValue("Zona", Zonas.Text.Trim)
                    .AddWithValue("IdDatos", IDUbicacion)
                End With
                cmd.ExecuteNonQuery()
                a = IDUbicacion
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try
            Cargar_Tabla("*", "ubicaciones")
            a = Tabla1.Rows((Tabla1.Rows.Count - 1)).ItemArray(0)
        End If

        DataGridView3.DataSource = Nothing
        DataGridView3.DataSource = Tabla1
        DataGridView3.ReadOnly = True
        DataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView3.Columns(0).Visible = False
        For Each row As DataGridViewRow In Me.DataGridView3.Rows
            'obtenemos el valor de la columna en la variable declarada
            If Convert.ToInt64(row.Cells(0).Value) = a Then
                DataGridView3.CurrentCell = DataGridView3(1, row.Index)
            End If
        Next
    End Sub

    Private Sub Nueva_Ubicacion_Click(sender As Object, e As EventArgs) Handles Nueva_Ubicacion.Click
        IDUbicacion = -1
        CargarCBTabUbicaciones()
        Estantes.SelectedIndex = -1
        Entrepanos.SelectedItem = -1
        Cajas_Colores.SelectedItem = -1
        Zonas.SelectedItem = -1
    End Sub

    Private Sub Eliminar_Ubicacion_Click(sender As Object, e As EventArgs) Handles Eliminar_Ubicacion.Click
        If MessageBox.Show("¿Esta seguro que desea ELIMINAR este registro?" & vbCrLf & "Esta accion no se puede deshacer", "Alerta", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Try
                conn.Open()
                Dim query As String = "DELETE FROM ubicaciones WHERE `Id_Ubicacion`=@IdUbicacion;"
                Dim cmd As New MySqlCommand(query, conn)
                cmd.Parameters.AddWithValue("IdUbicacion", IDUbicacion)
                cmd.ExecuteNonQuery()
                conn.Close()
                MsgBox("Registro Eliminado", MsgBoxStyle.Information, "Info.")
            Catch ex As Exception
                MsgBox("Error al eliminar el registro:" & vbCrLf & ex.Message, MsgBoxStyle.Information, "Info.")
                conn.Close()
            End Try
            Dim a = Tabla1.Rows((Tabla1.Rows.Count - 1)).ItemArray(0)
            Cargar_Tabla("*", "ubicaciones")
            DataGridView3.DataSource = Nothing
            DataGridView3.DataSource = Tabla1
            DataGridView3.ReadOnly = True
            DataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            DataGridView3.Columns(0).Visible = False
            For Each row As DataGridViewRow In Me.DataGridView3.Rows
                'obtenemos el valor de la columna en la variable declarada
                If Convert.ToInt64(row.Cells(0).Value) = a Then
                    DataGridView3.CurrentCell = DataGridView3(1, row.Index)
                End If
            Next
        Else
            Exit Sub
        End If
    End Sub

    Private Sub Movimiento_Ingreso_Click(sender As Object, e As EventArgs) Handles Movimiento_Ingreso.Click
        Esconder_tabpages_submenu()
        TabPage12.Parent = TabControl2
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(String.Format("SELECT NOW();"), conn)
            Dim fecha_servidor As DateTime = cmd.ExecuteScalar()
            Fecha_Movimiento.Text = fecha_servidor.ToString("yyyy-MM-dd")
            conn.Close()
        Catch
            MsgBox("No se puede recuperar la fecha de la base de datos", MsgBoxStyle.Exclamation, "Error")
            conn.Close()
        End Try
        Label86.Text = "Proveedor: *"
        Label82.Text = "Orden de Compra: *"
        Tipo_Movimiento.Text = "INGRESO"
        'Haciendo los controles necesarios para el ingreso visibles
        CBN_SolicitudSalida.Visible = False
        Label86.Visible = True
        Label85.Visible = True
        Label84.Visible = True
        Label83.Visible = True
        Observaciones_Movimiento.Visible = True
        N_Orden_Movimiento.Visible = True
        Monto_Movimiento.Visible = True
        N_Referencia_Movimiento.Visible = True
        N_Referencia_Movimiento.Enabled = True
        Proveedor_Movimiento.Visible = True
        N_Orden_Movimiento.Enabled = True
        N_Orden_Movimiento.Text = ""
        N_Referencia_Movimiento.Text = ""
        Cargar_Tabla("*", "Proveedores")
        With Proveedor_Movimiento
            '.Items.Clear()
            .Text = ""
            .DataSource = Tabla1
            .DisplayMember = "Nombre_Proveedor" 'elnombre de tu columna de tu base de datos q deseas mostrar
            .ValueMember = "Nit_Proveedor"
        End With
        Datos_Movimientos.Visible = True
        Confirmar_Transaccion.Visible = True
    End Sub

    Dim Ord_Movimiento_num = 0
    Dim Id_Ord_Movimiento As Integer
    Private Sub Confirmar_Transaccion_Click(sender As Object, e As EventArgs) Handles Confirmar_Transaccion.Click
        If Tipo_Movimiento.Text = "INGRESO" Then
            If N_Orden_Movimiento.Text.Trim = "" Or N_Referencia_Movimiento.Text.Trim = "" Then
                MsgBox("Los campos con * son obligatorios")
                Exit Sub
            End If

            Try
                conn.Open()
                Dim read As MySqlDataReader
                Dim cmd As New MySqlCommand("Select * from orden_movimientos where N_Orden_Compra = @NumOrden and N_Referencia = @Referencia", conn)
                With cmd.Parameters
                    .AddWithValue("NumOrden", UCase(N_Orden_Movimiento.Text.Trim))
                    .AddWithValue("Referencia", UCase(N_Referencia_Movimiento.Text.Trim))
                End With
                read = cmd.ExecuteReader
                If read.Read Then
                    MsgBox("Un numero de orden ingresado ya se encuentra con el mismo numero de remision en la base de datos, Revise los datos e intentelo de nuevo.", MsgBoxStyle.Exclamation, "Alerta.")
                    read.Close()
                    conn.Close()
                    Exit Sub
                End If
                read.Close()
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try
            Try
                conn.Open()
                Dim read As MySqlDataReader
                Dim cmd As New MySqlCommand("Select * from orden_movimientos where N_Orden_Compra = @NumOrden", conn)
                With cmd.Parameters
                    .AddWithValue("NumOrden", UCase(N_Orden_Movimiento.Text.Trim))
                End With
                read = cmd.ExecuteReader
                If read.Read Then
                    DataGridView4.Visible = True
                    Dim cmd2 As New MySqlCommand("SELECT nombre_producto as 'Producto', Cantidad, Precio_Compra as 'Precio', Descripcion
                                                FROM movimientos inner join productos on movimientos.Id_Producto = productos.Id_Producto
                                                inner join orden_movimientos on movimientos.IdOrden_Movimiento = orden_movimientos.IdOrden_Movimiento
                                                WHERE N_Orden_Compra = @Orden and tipo = 'INGRESO';", conn)
                    With cmd2.Parameters
                        .AddWithValue("Orden", N_Orden_Movimiento.Text.Trim)
                    End With
                    read.Close()
                    Dim reader As MySqlDataReader
                    Dim Tabla As New DataTable
                    reader = cmd2.ExecuteReader
                    Tabla.Load(reader)
                    DataGridView4.DataSource = Tabla
                    DataGridView4.ReadOnly = True
                    DataGridView4.AllowUserToResizeColumns = True
                    DataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect
                    reader.Close()
                    Dim cmd3 As New MySqlCommand("Select max(monto) from orden_movimientos where N_Orden_Compra = @Orden;", conn)
                    With cmd3.Parameters
                        .AddWithValue("Orden", N_Orden_Movimiento.Text.Trim)
                    End With
                    Monto_Movimiento.Text = cmd3.ExecuteScalar.ToString
                    conn.Close()
                End If
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
                conn.Close()
            End Try
            If Monto_Movimiento.Text.Trim = "" Then
                Monto_Movimiento.Text = 0
            End If
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("INSERT INTO orden_movimientos (N_Orden_Compra, N_Referencia, Tipo, Fecha, Observaciones, Nit_Proveedor,Monto)
                                            VALUES (@NumOrden, @NumRef, @Tipo, @Fecha, @Obs, @NitProv, @Monto);", conn)
                With cmd.Parameters
                    .AddWithValue("NumOrden", UCase(N_Orden_Movimiento.Text.Trim))
                    .AddWithValue("NumRef", UCase(N_Referencia_Movimiento.Text.Trim))
                    .AddWithValue("Tipo", UCase(Tipo_Movimiento.Text.Trim))
                    .AddWithValue("Fecha", UCase(Fecha_Movimiento.Text.Trim))
                    .AddWithValue("Obs", UCase(Observaciones_Movimiento.Text.Trim))
                    .AddWithValue("NitProv", UCase(Proveedor_Movimiento.SelectedValue))
                    .AddWithValue("Monto", Monto_Movimiento.Text.Trim)
                End With
                cmd.ExecuteNonQuery()
                conn.Close()
            Catch ex As Exception
                MsgBox("No se pudo registrar el movimiento:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
                Exit Sub
            End Try
            Gru_Movimiento.Visible = True
            Producto_Movimiento.SelectedItem = -1
            Usuario_Movimiento.SelectedItem = -1
            Usuario_Movimiento.Visible = False
            Label95.Text = "Precio"
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("SELECT * FROM rel_productos_proveedores INNER JOIN productos 
                                            ON (rel_productos_proveedores.Id_Producto = productos.Id_Producto) 
                                            INNER JOIN proveedores ON (rel_productos_proveedores.Nit_Proveedor= proveedores.Nit_Proveedor)
                                            WHERE rel_productos_proveedores.Nit_Proveedor= @NitProv;", conn)
                cmd.Parameters.AddWithValue("NitProv", Proveedor_Movimiento.SelectedValue)
                Dim adaptador As New MySqlDataAdapter(cmd)
                Dim tabla As New DataTable
                Try
                    adaptador.Fill(tabla)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                    Exit Sub
                End Try
                conn.Close()
                Tabla1 = tabla
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
                Exit Sub
            End Try
            With Producto_Movimiento
                .DataSource = Tabla1
                .DisplayMember = "Nombre_Producto" 'elnombre de tu columna de tu base de datos q deseas mostrar
                .ValueMember = "Id_Producto" 'el ide de tu tabla relacionada con el nombre que muestras muy importante para saber el ide de quien seleccionas en tu combobox
                '.Enabled = False
            End With
        ElseIf Tipo_Movimiento.Text = "SALIDA" Then
            CBN_SolicitudSalida.Enabled = False
            Gru_Movimiento.Visible = True
            Precio_Movimiento.Visible = False
            Label95.Text = "Solicitado Por:"
            With Usuario_Movimiento
                .DataSource = Nothing
                .Items.Clear()
                .Enabled = False
                Try
                    conn.Open()
                    Dim query As String = "SELECT distinct usuarios.Nombre_Usuario, solicitador 
                                           from solicitud_salida inner join usuarios on id_usuario = solicitador 
                                           where NumeroOrdenTrabajo = @Orden"
                    Dim cmd As New MySqlCommand(query, conn)
                    With cmd.Parameters
                        .AddWithValue("Orden", CBN_SolicitudSalida.SelectedValue)
                    End With
                    Dim sqladap As New MySqlDataAdapter(cmd)
                    Dim dtRecord As New DataTable
                    sqladap.Fill(dtRecord)
                    .DataSource = dtRecord
                    .DisplayMember = "Nombre_Usuario"
                    .ValueMember = "solicitador"
                    conn.Close()
                Catch ex As Exception
                    MsgBox("No se pudo recuperar el nombre de quien solicito la salida" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                    Exit Sub
                End Try
            End With

            With Producto_Movimiento
                .DataSource = Nothing
                .Items.Clear()
                .Enabled = True
                Try
                    conn.Open()
                    Dim query As String = "Select productos.Nombre_Producto, productos.Id_Producto
                                           From solicitud_salida inner Join productos On solicitud_salida.producto = productos.Id_Producto 
                                           Where NumeroOrdenTrabajo = @Orden and pendiente = '1';"
                    Dim cmd As New MySqlCommand(query, conn)
                    With cmd.Parameters
                        .AddWithValue("Orden", CBN_SolicitudSalida.SelectedValue)
                    End With
                    Dim SqlAdap As New MySqlDataAdapter(cmd)
                    Dim dtRecord As New DataTable
                    SqlAdap.Fill(dtRecord)
                    .DataSource = dtRecord
                    .DisplayMember = "Nombre_Producto"
                    .ValueMember = "Id_Producto"
                    conn.Close()
                Catch ex As Exception
                    MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                    Exit Sub
                End Try

            End With

            With DataGridView4
                .DataSource = Nothing
                Try
                    conn.Open()
                    Dim query As String = "SELECT productos.Nombre_Producto as 'Producto Solicitado', solicitud_salida.Cantidad as 'Cantidad Solicitada'
                                           from solicitud_salida inner join productos on solicitud_salida.producto = productos.Id_Producto 
                                           where NumeroOrdenTrabajo = @Orden and pendiente = '1';"
                    Dim cmd As New MySqlCommand(query, conn)
                    With cmd.Parameters
                        .AddWithValue("Orden", CBN_SolicitudSalida.SelectedValue)
                    End With
                    Dim reader As MySqlDataReader
                    reader = cmd.ExecuteReader
                    Dim T As New DataTable
                    T.Load(reader)
                    .DataSource = T
                    .ReadOnly = True
                    .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                    .Visible = True
                    conn.Close()
                Catch ex As Exception
                    MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                    Exit Sub
                End Try

            End With

            Try
                conn.Open()
                Dim cmd As New MySqlCommand("INSERT INTO orden_movimientos (N_Orden_Compra, Tipo, Fecha)
                                            VALUES (@NumOrden, @Tipo, @Fecha);", conn)
                With cmd.Parameters
                    .AddWithValue("NumOrden", UCase(CBN_SolicitudSalida.SelectedValue))
                    .AddWithValue("Tipo", UCase(Tipo_Movimiento.Text.Trim))
                    .AddWithValue("Fecha", UCase(Fecha_Movimiento.Text.Trim))
                End With
                cmd.ExecuteNonQuery()
                conn.Close()
            Catch ex As Exception
                MsgBox("No se pudo registrar el movimiento:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
                Exit Sub
            End Try

        End If
        Movimiento_Ingreso.Enabled = False
        Movimiento_Salida.Enabled = False
        Generar_Movimientos.Visible = True
        Confirmar_Transaccion.Visible = False
        N_Orden_Movimiento.Enabled = False
        N_Referencia_Movimiento.Enabled = False
        Monto_Movimiento.Enabled = False
        Observaciones_Movimiento.Enabled = False
        Proveedor_Movimiento.Enabled = False
        Cargar_Tabla("*", "orden_movimientos")
        Ord_Movimiento_num = Tabla1.Rows.Count - 1
        Id_Ord_Movimiento = Tabla1.Rows(Ord_Movimiento_num).ItemArray(0).ToString
        N_Orden_Movimiento.Text = Tabla1.Rows(Ord_Movimiento_num).ItemArray(2).ToString
        N_Referencia_Movimiento.Text = Tabla1.Rows(Ord_Movimiento_num).ItemArray(3).ToString
        Observaciones_Movimiento.Text = Tabla1.Rows(Ord_Movimiento_num).ItemArray(5).ToString
        Monto_Movimiento.Text = Tabla1.Rows(Ord_Movimiento_num).ItemArray(6).ToString
        Proveedor_Movimiento.SelectedValue = Tabla1.Rows(Ord_Movimiento_num).ItemArray(7).ToString
    End Sub

    Private Sub Modificar_Prod_Equ_Click(sender As Object, e As EventArgs) Handles Modificar_Prod_Equ.Click
        Bandera_Rel = 1

        Consulta_rel = "SELECT Cod_Producto, Nombre_Producto,IdRel_Producto_Equipos FROM rel_productos_equipos INNER JOIN productos
                        ON (rel_productos_equipos.Id_Producto = productos.Id_Producto) INNER JOIN equipos ON (rel_productos_equipos.Id_Equipo = equipos.Id_Equipo)
                        WHERE rel_productos_equipos.Id_Equipo='" & Id_Equipo & "'"
        Id_elem = Id_Equipo
        Tabla_Rel = "Productos"
        Elemento_rel = Nombre_Equipo.Text.Trim
        Form3.ShowDialog()
    End Sub

    Private Sub Equipo_Producto_Click(sender As Object, e As EventArgs) Handles Equipo_Producto.Click
        Bandera_Rel = 1
        Consulta_rel = "SELECT Cod_Equipo, Nombre_Equipo,IdRel_Producto_Equipos FROM rel_productos_equipos " &
                        "INNER JOIN productos " &
                        "ON (rel_productos_equipos.Id_Producto = productos.Id_Producto) " &
                        "INNER JOIN equipos " &
                        "ON (rel_productos_equipos.Id_Equipo = equipos.Id_Equipo) " &
                        "WHERE rel_productos_equipos.Id_Producto='" & Id_Prod & "'"
        Id_elem = Id_Prod
        Tabla_Rel = "Equipos"
        Elemento_rel = Nombre_Producto.Text
        Form3.ShowDialog()
    End Sub

    Private Sub Proveedor_Producto_Click(sender As Object, e As EventArgs) Handles Proveedor_Producto.Click
        Bandera_Rel = 2
        Consulta_rel = "SELECT Nombre_Proveedor, Ciudad, IdRel_Productos_Proveedores FROM rel_productos_proveedores " &
                        "INNER JOIN productos " &
                        "ON (rel_productos_proveedores.Id_Producto = productos.Id_Producto) " &
                        "INNER JOIN proveedores " &
                        "ON (rel_productos_proveedores.Nit_Proveedor= proveedores.Nit_Proveedor) " &
                        "WHERE rel_productos_proveedores.Id_Producto='" & Id_Prod & "'"
        Id_elem = Id_Prod
        Tabla_Rel = "Proveedores"
        Form3.ShowDialog()
    End Sub

    Private Sub Ubicacion_Producto_Click(sender As Object, e As EventArgs) Handles Ubicacion_Producto.Click
        Bandera_Rel = 3
        Consulta_rel = "SELECT Estante,Entrepano,Caja_Color,Zona,Cantidad,Aforo,IdRel_Ubicaciones_Productos FROM rel_ubicaciones_productos " &
                        "INNER JOIN productos " &
                        "ON (rel_ubicaciones_productos.Id_Producto = productos.Id_Producto) " &
                        "INNER JOIN ubicaciones " &
                        "ON (rel_ubicaciones_productos.Id_Ubicacion = ubicaciones.Id_Ubicacion) " &
                        "WHERE rel_ubicaciones_productos.Id_Producto='" & Id_Prod & "'"
        Id_elem = Id_Prod
        Tabla_Rel = "Ubicaciones"
        Elemento_rel = Nombre_Producto.Text
        Form3.ShowDialog()
    End Sub

    Private Sub Modificar_Prod_Prov_Click(sender As Object, e As EventArgs) Handles Modificar_Prod_Prov.Click
        Bandera_Rel = 2
        Consulta_rel = "SELECT Cod_Producto, Nombre_Producto,IdRel_Productos_Proveedores FROM rel_productos_proveedores " &
                        "INNER JOIN productos " &
                        "ON (rel_productos_proveedores.Id_Producto = productos.Id_Producto) " &
                        "INNER JOIN proveedores " &
                        "ON (rel_productos_proveedores.Nit_Proveedor= proveedores.Nit_Proveedor) " &
                        "WHERE rel_productos_proveedores.Nit_Proveedor='" & ID_Prov & "'"
        Id_elem = ID_Prov
        Tabla_Rel = "Productos"
        Elemento_rel = Nombre_Proveedor.Text
        Form3.ShowDialog()
    End Sub

    Private Sub Generar_Movimientos_Click(sender As Object, e As EventArgs) Handles Generar_Movimientos.Click
        DataGridView4.Visible = True
        Dim consulta_movimientos As String = ""
        Dim stock As Integer
        If IsNumeric(Cantidad_Movimiento.Text) Then
            Cantidad_a_mover = Cantidad_Movimiento.Text
            Tipo_Movi = Tipo_Movimiento.Text
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("SELECT Id_Producto, stock_minimo, Stock_Maximo, Stock_Existente FROM productos where Id_Producto = @ProductoMovimiento;", conn)
                With cmd.Parameters
                    .AddWithValue("ProductoMovimiento", Producto_Movimiento.SelectedValue)
                End With
                Dim adaptador As New MySqlDataAdapter(cmd)
                Dim Tabla As New DataTable
                adaptador.Fill(Tabla)
                conn.Close()
                Tabla1 = Tabla
            Catch ex As Exception
                MsgBox("Error: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try
            Id_Prod = Tabla1.Rows(0).ItemArray(0).ToString
            If Tipo_Movimiento.Text = "INGRESO" Then
                If Descripcion_Movimiento.Text.Trim = "" Or Producto_Movimiento.Text = "" Or Precio_Movimiento.Text = "" Then
                    MsgBox("Faltan algunos datos para generar el movimiento")
                    Exit Sub
                End If

                If Convert.ToInt64(Tabla1.Rows(0).ItemArray(2).ToString) < (Convert.ToInt64(Cantidad_Movimiento.Text) + Convert.ToInt64(Tabla1.Rows(0).ItemArray(3).ToString)) Then
                    MsgBox("Atencion NO puede realizar el movimiento, porque su movimiento supera el máximo permitido")
                    Exit Sub
                End If
                stock = Convert.ToInt64(Tabla1.Rows(0).ItemArray(3).ToString) + Convert.ToInt64(Cantidad_Movimiento.Text)

            ElseIf Tipo_Movimiento.Text = "SALIDA" Then

                If Descripcion_Movimiento.Text.Trim = "" Or Cantidad_Movimiento.Text.Trim = "" Then
                    MsgBox("Faltan algunos datos para poder generar la salida", MsgBoxStyle.Information, "Info")
                    Exit Sub
                End If

                If Cantidad_Movimiento.Text = 0 Then
                    If MessageBox.Show("¿Esta seguro que desea rechazar la solicitud de salida de este producto?", "Alerta", MessageBoxButtons.YesNo) = DialogResult.No Then
                        Exit Sub
                    End If
                    Nueva_Transaccion.Visible = True
                    Try
                        conn.Open()
                        Dim cmd As New MySqlCommand(String.Format("UPDATE bd_inventario2.solicitud_salida SET Pendiente = '0' WHERE NumeroOrdenTrabajo = @Orden and Producto = @Prod;"), conn)
                        With cmd.Parameters
                            .AddWithValue("Orden", CBN_SolicitudSalida.SelectedValue)
                            .AddWithValue("Prod", Producto_Movimiento.SelectedValue)
                        End With
                        cmd.ExecuteNonQuery()
                        conn.Close()
                    Catch ex As Exception
                        MsgBox("No se pudo cambiar el estado de pendiente del producto retirado en esta solicitud, por favor realice este cambio manualmente", MsgBoxStyle.Information, "Info.")
                        conn.Close()
                    End Try

                    With Producto_Movimiento
                        .DataSource = Nothing
                        .Items.Clear()
                        .Enabled = True
                        Try
                            conn.Open()
                            Dim query As String = "Select productos.Nombre_Producto, productos.Id_Producto
                                           From solicitud_salida inner Join productos On solicitud_salida.producto = productos.Id_Producto 
                                           Where NumeroOrdenTrabajo = @Orden and pendiente = '1';"
                            Dim cmd As New MySqlCommand(query, conn)
                            With cmd.Parameters
                                .AddWithValue("Orden", CBN_SolicitudSalida.SelectedValue)
                            End With
                            Dim SqlAdap As New MySqlDataAdapter(cmd)
                            Dim dtRecord As New DataTable
                            SqlAdap.Fill(dtRecord)
                            .DataSource = dtRecord
                            .DisplayMember = "Nombre_Producto"
                            .ValueMember = "Id_Producto"
                            conn.Close()
                        Catch ex As Exception
                            MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                            conn.Close()
                            Exit Sub
                        End Try

                    End With

                    With DataGridView4
                        .DataSource = Nothing
                        Try
                            conn.Open()
                            Dim query As String = "SELECT productos.Nombre_Producto as 'Producto Solicitado', solicitud_salida.Cantidad as 'Cantidad Solicitada'
                                           from solicitud_salida inner join productos on solicitud_salida.producto = productos.Id_Producto 
                                           where NumeroOrdenTrabajo = @Orden and pendiente = '1';"
                            Dim cmd As New MySqlCommand(query, conn)
                            With cmd.Parameters
                                .AddWithValue("Orden", CBN_SolicitudSalida.SelectedValue)
                            End With
                            Dim reader As MySqlDataReader
                            reader = cmd.ExecuteReader
                            Dim T As New DataTable
                            T.Load(reader)
                            .DataSource = T
                            .ReadOnly = True
                            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                            .Visible = True
                            conn.Close()
                        Catch ex As Exception
                            MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                            conn.Close()
                            Exit Sub
                        End Try

                    End With
                    Exit Sub
                End If

                If Convert.ToInt64(Tabla1.Rows(0).ItemArray(3).ToString) - Convert.ToInt64(Cantidad_Movimiento.Text) < 0 Then
                    MsgBox("No puede llevar a cabo este movimiento porque la cantidad que quiere retirar supera las existencias en el inventario", MsgBoxStyle.Information, "Error.")
                    Exit Sub
                End If

                If Convert.ToInt64(Tabla1.Rows(0).ItemArray(1).ToString) > (Convert.ToInt64(Tabla1.Rows(0).ItemArray(3).ToString) - Convert.ToInt64(Cantidad_Movimiento.Text)) Then
                    If MessageBox.Show("Alerta al realizar este movimiento va a estar por debajo del minimo establecido para el producto. ¿Desea continuar?", "Alerta", MessageBoxButtons.YesNo) = DialogResult.No Then
                        Exit Sub
                    End If
                End If

                stock = Convert.ToInt64(Tabla1.Rows(0).ItemArray(3).ToString) - Convert.ToInt64(Cantidad_Movimiento.Text)

            End If
            Using conn
                conn.Open()
                Dim comando As New MySqlCommand("SELECT Estante,Entrepano,Caja_Color,Zona,Cantidad,Aforo FROM rel_ubicaciones_productos " &
                            "INNER JOIN productos " &
                            "ON (rel_ubicaciones_productos.Id_Producto = productos.Id_Producto) " &
                            "INNER JOIN ubicaciones " &
                            "ON (rel_ubicaciones_productos.Id_Ubicacion = ubicaciones.Id_Ubicacion) " &
                            "WHERE rel_ubicaciones_productos.Id_Producto= @idProd", conn)
                With comando.Parameters
                    .AddWithValue("idProd", Producto_Movimiento.SelectedValue)
                End With
                Dim adaptador As New MySqlDataAdapter(comando)
                Dim Tabla As New DataTable
                Try
                    adaptador.Fill(Tabla)
                Catch ex As Exception
                Finally
                    If conn.State = ConnectionState.Open Then
                        conn.Close()
                    End If
                End Try
                Tabla1 = Tabla
                conn.Close()
            End Using
            If Tabla1.Rows.Count >= 1 Then
                Bandera_Rel = 4
                Consulta_rel = "SELECT Estante,Entrepano,Caja_Color,Zona,Cantidad,Aforo,IdRel_Ubicaciones_Productos FROM rel_ubicaciones_productos " &
                "INNER JOIN productos " &
                "ON (rel_ubicaciones_productos.Id_Producto = productos.Id_Producto) " &
                "INNER JOIN ubicaciones " &
                "ON (rel_ubicaciones_productos.Id_Ubicacion = ubicaciones.Id_Ubicacion) " &
                "WHERE rel_ubicaciones_productos.Id_Producto='" & Producto_Movimiento.SelectedValue & "'"
                Id_elem = Producto_Movimiento.SelectedValue
                'Tabla_Rel = "Ubicaciones"
                'Elemento_rel = Nombre_Producto.Text
                Try
                    Using conn
                        conn.Open()
                        Dim cmd As New MySqlCommand("UPDATE productos SET Stock_Existente = '" & stock &
                        "' WHERE Id_Producto = @idProd", conn)
                        With cmd.Parameters
                            .AddWithValue("idProd", Producto_Movimiento.SelectedValue)
                        End With
                        cmd.ExecuteNonQuery()
                        'MessageBox.Show("Registro MODIFICADO")
                        conn.Close()
                    End Using
                Catch ex As Exception
                    MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
                End Try
                Nueva_Transaccion.Visible = True
                If Tipo_Movimiento.Text = "INGRESO" Then
                    Try
                        Using conn
                            conn.Open()
                            Dim cmd As New MySqlCommand("INSERT INTO movimientos (IdOrden_Movimiento, Id_Producto, Cantidad, Precio_Compra, Id_Usuario, Fecha, Descripcion)
                                                        VALUES (@IdOrd, @IdProd, @Cantidad, @Precio, @Usuario, @Fecha, @Desc)", conn)
                            With cmd.Parameters
                                .AddWithValue("IdOrd", UCase(Id_Ord_Movimiento))
                                .AddWithValue("IdProd", Producto_Movimiento.SelectedValue)
                                .AddWithValue("Cantidad", Cantidad_Movimiento.Text)
                                .AddWithValue("Precio", Precio_Movimiento.Text)
                                .AddWithValue("Usuario", id_Usuar_Per)
                                .AddWithValue("Fecha", Fecha_Movimiento.Text)
                                .AddWithValue("Desc", UCase(Descripcion_Movimiento.Text))
                            End With
                            cmd.ExecuteNonQuery()
                            Dim cmd2 As New MySqlCommand("UPDATE orden_movimientos SET Monto = @NewMonto WHERE N_Orden_Compra = @Orden and N_Referencia = @Remision;", conn)
                            With cmd2.Parameters
                                .AddWithValue("NewMonto", (Convert.ToInt64(Monto_Movimiento.Text) + Convert.ToInt64(Cantidad_Movimiento.Text) * Convert.ToInt64(Precio_Movimiento.Text)))
                                .AddWithValue("Orden", N_Orden_Movimiento.Text)
                                .AddWithValue("Remision", N_Referencia_Movimiento.Text)
                            End With
                            cmd2.ExecuteNonQuery()
                            Monto_Movimiento.Text = (Convert.ToInt64(Monto_Movimiento.Text) + Convert.ToInt64(Cantidad_Movimiento.Text) * Convert.ToInt64(Precio_Movimiento.Text))
                            'MessageBox.Show("Registro MODIFICADO")
                            conn.Close()
                        End Using
                    Catch ex As Exception
                        MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
                    End Try
                ElseIf Tipo_Movimiento.Text = "SALIDA" Then
                    Try
                        Using conn
                            conn.Open()
                            Dim cmd As New MySqlCommand("INSERT INTO movimientos (IdOrden_Movimiento, Id_Producto, Cantidad, Id_Usuario, Fecha, Descripcion,Id_Usuario_Final)
                                                        VALUES (@IdOrd, @IdProd, @Cantidad, @Usuario, @Fecha, @Desc, @Final)", conn)
                            With cmd.Parameters
                                .AddWithValue("IdOrd", UCase(Id_Ord_Movimiento))
                                .AddWithValue("IdProd", Producto_Movimiento.SelectedValue)
                                .AddWithValue("Cantidad", Cantidad_Movimiento.Text)
                                .AddWithValue("Usuario", id_Usuar_Per)
                                .AddWithValue("Fecha", Fecha_Movimiento.Text)
                                .AddWithValue("Desc", UCase(Descripcion_Movimiento.Text))
                                .AddWithValue("Final", Usuario_Movimiento.SelectedValue)
                            End With
                            cmd.ExecuteNonQuery()
                            conn.Close()
                        End Using
                    Catch ex As Exception
                        MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
                    End Try
                End If

                If Tipo_Movimiento.Text = "INGRESO" Then
                    consulta_movimientos = "SELECT nombre_producto as 'Producto', Cantidad, Precio_Compra as 'Precio', Descripcion
                                                FROM movimientos inner join productos on movimientos.Id_Producto = productos.Id_Producto
                                                inner join orden_movimientos on movimientos.IdOrden_Movimiento = orden_movimientos.IdOrden_Movimiento
                                                WHERE N_Orden_Compra = @Orden and tipo = 'INGRESO';"
                    Using conn
                        conn.Open()
                        Dim cmd As New MySqlCommand(consulta_movimientos, conn)
                        With cmd.Parameters
                            .AddWithValue("Orden", N_Orden_Movimiento.Text)
                        End With
                        Dim adaptador As New MySqlDataAdapter(cmd)
                        Dim tabla As New DataTable
                        Try
                            adaptador.Fill(tabla)
                            conn.Close()
                        Catch ex As Exception
                        Finally
                            If conn.State = ConnectionState.Open Then
                                conn.Close()
                            End If
                        End Try
                        DataGridView4.DataSource = tabla
                    End Using
                ElseIf Tipo_Movimiento.Text = "SALIDA" Then
                    Try
                        conn.Open()
                        Dim cmd As New MySqlCommand(String.Format("UPDATE bd_inventario2.solicitud_salida SET Pendiente = '0' WHERE NumeroOrdenTrabajo = @Orden and Producto = @Prod;"), conn)
                        With cmd.Parameters
                            .AddWithValue("Orden", CBN_SolicitudSalida.SelectedValue)
                            .AddWithValue("Prod", Producto_Movimiento.SelectedValue)
                        End With
                        cmd.ExecuteNonQuery()
                        conn.Close()
                    Catch ex As Exception
                        MsgBox("No se pudo cambiar el estado de pendiente del producto retirado en esta solicitud, por favor realice este cambio manualmente", MsgBoxStyle.Information, "Info.")
                        conn.Close()
                    End Try

                    With Producto_Movimiento
                        .DataSource = Nothing
                        .Items.Clear()
                        .Enabled = True
                        Try
                            conn.Open()
                            Dim query As String = "Select productos.Nombre_Producto, productos.Id_Producto
                                           From solicitud_salida inner Join productos On solicitud_salida.producto = productos.Id_Producto 
                                           Where NumeroOrdenTrabajo = @Orden and pendiente = '1';"
                            Dim cmd As New MySqlCommand(query, conn)
                            With cmd.Parameters
                                .AddWithValue("Orden", CBN_SolicitudSalida.SelectedValue)
                            End With
                            Dim SqlAdap As New MySqlDataAdapter(cmd)
                            Dim dtRecord As New DataTable
                            SqlAdap.Fill(dtRecord)
                            .DataSource = dtRecord
                            .DisplayMember = "Nombre_Producto"
                            .ValueMember = "Id_Producto"
                            conn.Close()
                        Catch ex As Exception
                            MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                            conn.Close()
                            Exit Sub
                        End Try

                    End With

                    With DataGridView4
                        .DataSource = Nothing
                        Try
                            conn.Open()
                            Dim query As String = "SELECT productos.Nombre_Producto as 'Producto Solicitado', solicitud_salida.Cantidad as 'Cantidad Solicitada'
                                           from solicitud_salida inner join productos on solicitud_salida.producto = productos.Id_Producto 
                                           where NumeroOrdenTrabajo = @Orden and pendiente = '1';"
                            Dim cmd As New MySqlCommand(query, conn)
                            With cmd.Parameters
                                .AddWithValue("Orden", CBN_SolicitudSalida.SelectedValue)
                            End With
                            Dim reader As MySqlDataReader
                            reader = cmd.ExecuteReader
                            Dim T As New DataTable
                            T.Load(reader)
                            .DataSource = T
                            .ReadOnly = True
                            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                            .Visible = True
                            conn.Close()
                        Catch ex As Exception
                            MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                            conn.Close()
                            Exit Sub
                        End Try

                    End With
                End If
                Descripcion_Movimiento.Text = ""
                Precio_Movimiento.Text = ""
                Cantidad_Movimiento.Text = ""
                Form3.ShowDialog()
            Else
                MsgBox("Atencion NO puede realizar el movimiento, porque el producto no cuenta con ubicaciones", MsgBoxStyle.Information, "Info.")
                Exit Sub
            End If
        Else
            MsgBox("El campo de cantidad solo acepta valores numericos", MsgBoxStyle.Exclamation, "Error.")
        End If
    End Sub

    Private Sub Solicitudes_Click(sender As Object, e As EventArgs) Handles Solicitudes.Click
        Esconder_tabpages_submenu()
        TabPage17.Parent = TabControl2
        CargarCBTabSolicitudes()
    End Sub

    Dim EquipLoad As Boolean = False
    Private Sub CargarCBTabSolicitudes()
        EquipLoad = False
        With CBEquipos
            Try
                conn.Open()
                Dim query As String = "Select Id_Equipo, Nombre_Equipo from equipos"
                Dim cmd As New MySqlCommand(query, conn)
                Dim sqlAdap As New MySqlDataAdapter(cmd)
                Dim dtRecord As New DataTable
                sqlAdap.Fill(dtRecord)
                .DataSource = dtRecord
                .DisplayMember = "Nombre_Equipo"
                .ValueMember = "Id_Equipo"
                .SelectedValue = dtRecord.Rows(0).Item(0)
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar los equipos de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
            EquipLoad = True
        End With
    End Sub

    Private Sub CBEquipos_SelectedValueChanged(sender As Object, e As EventArgs) Handles CBEquipos.SelectedValueChanged
        If EquipLoad = False Then
            Exit Sub
        End If
        conn.Open()
        Dim query As String = "select productos.Id_Producto, Nombre_Producto as 'Producto', Stock_Existente as 'Stock' from productos inner join rel_productos_equipos on productos.Id_Producto = rel_productos_equipos.Id_Producto
                               where Id_Equipo = @Equipo;"
        Dim cmd As New MySqlCommand(query, conn)
        With cmd.Parameters
            .AddWithValue("Equipo", CBEquipos.SelectedValue)
        End With
        Dim Adaptador As New MySqlDataAdapter(cmd)
        Dim Tabla As New DataTable

        Adaptador.Fill(Tabla)
        DGVProductosEquipo.DataSource = Tabla
        DGVProductosEquipo.ReadOnly = True
        DGVProductosEquipo.Columns(0).Visible = False
        DGVProductosEquipo.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        conn.Close()
        showmsg = True


    End Sub
    Dim showmsg As Boolean = True
    Private Sub BtnAgregarProducto_Click(sender As Object, e As EventArgs) Handles BtnAgregarProducto.Click
        If DGVProductosEquipo.Rows.Count <= 0 Then
            Exit Sub
        End If
        Dim fila_actual_producto As Integer = DGVProductosEquipo.CurrentRow.Index
        Dim IdProducto As Integer = DGVProductosEquipo(0, (fila_actual_producto)).Value
        Dim Producto As String = DGVProductosEquipo(1, (fila_actual_producto)).Value
        Dim stock As Integer = DGVProductosEquipo(2, (fila_actual_producto)).Value
        Dim Cantidad As Integer = CantidadProducto.Value
        Dim find As Boolean
        If DGVListaProductos.Rows.Count > 0 Then
            For row As Integer = 0 To DGVListaProductos.Rows.Count - 1
                If IdProducto = DGVListaProductos(0, row).Value Then
                    DGVListaProductos(2, row).Value = DGVListaProductos(2, row).Value + Cantidad
                    If DGVListaProductos(2, row).Value > stock And showmsg Then
                        MsgBox("La solicitud puede demorarse mas ya que esta solicitando una cantidad mayor a la que hay en stock ahora mismo", MsgBoxStyle.Information, "Info.")
                        showmsg = False
                    End If
                    find = True
                    Exit For
                Else
                    find = False
                End If
            Next
            If find = False Then
                DGVListaProductos.Rows.Add(IdProducto, Producto, Cantidad)
                If Cantidad > stock Then
                    MsgBox("La solicitud puede demorarse mas ya que esta solicitando una cantidad mayor a la que hay en stock ahora mismo", MsgBoxStyle.Information, "Info.")
                    showmsg = False
                End If
            End If
        Else
            DGVListaProductos.Rows.Add(IdProducto, Producto, Cantidad)
            If Cantidad > stock Then
                MsgBox("La solicitud puede demorarse mas ya que esta solicitando una cantidad mayor a la que hay en stock ahora mismo", MsgBoxStyle.Information, "Info.")
                showmsg = False
            End If
        End If

    End Sub

    Private Sub BtnQuitarProducto_Click(sender As Object, e As EventArgs) Handles BtnQuitarProducto.Click
        If DGVListaProductos.Rows.Count <= 0 Then
            Exit Sub
        End If
        Dim Fila_actual As Integer = DGVListaProductos.CurrentRow.Index
        Dim Cantidad As Integer = CantidadProducto.Value
        DGVListaProductos(2, Fila_actual).Value = DGVListaProductos(2, Fila_actual).Value - Cantidad
        If DGVListaProductos(2, Fila_actual).Value <= 0 Then
            DGVListaProductos.Rows.RemoveAt(Fila_actual)
        End If
    End Sub

    Private Sub BtnConfirmarSolicitud_Click(sender As Object, e As EventArgs) Handles BtnConfirmarSolicitud.Click
        If DGVListaProductos.Rows.Count <= 0 Then
            MsgBox("Agregue objetos a la lista antes de enviar la solicitud", MsgBoxStyle.Information, "Info.")
            Exit Sub
        End If
        Dim fecha As String
        Dim consecutivo As String
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(String.Format("SELECT NOW();"), conn)
            Dim fecha_servidor As DateTime = cmd.ExecuteScalar()
            conn.Close()
            fecha = fecha_servidor.ToString("yyyyMMdd")
        Catch ex As Exception
            MsgBox("La solicitud no pudo enviarse error:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
            conn.Close()
            Exit Sub
        End Try

        Try
            conn.Open()
            Dim cmd As New MySqlCommand(String.Format("select count(*)+1 as c from solicitud_salida where NumeroOrdenTrabajo like '" & fecha & "%';"), conn)
            consecutivo = cmd.ExecuteScalar()
            conn.Close()
        Catch ex As Exception
            MsgBox("La solicitud no pudo enviarse error:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
            conn.Close()
            Exit Sub
        End Try
        Dim OrdenSalida As String = fecha & "-" & consecutivo

        For row As Integer = 0 To DGVListaProductos.Rows.Count - 1
            Try
                conn.Open()
                Dim cmd As New MySqlCommand("INSERT INTO solicitud_salida (NumeroOrdenTrabajo, Producto, Cantidad, Solicitador) VALUES (@OrdenS, @Producto, @Cantidad, @Usuario);", conn)
                With cmd.Parameters
                    .AddWithValue("OrdenS", OrdenSalida)
                    .AddWithValue("Producto", DGVListaProductos(0, row).Value)
                    .AddWithValue("Cantidad", DGVListaProductos(2, row).Value)
                    .AddWithValue("Usuario", id_Usuar_Per)
                End With
                cmd.ExecuteNonQuery()
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al intentar solicitar el producto: " & DGVListaProductos(1, row).Value & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try
        Next

        MsgBox("Solicitud Enviada", MsgBoxStyle.Information, "Info.")
        DGVListaProductos.Rows.Clear()

    End Sub

    Private Sub Movimiento_Salida_Click(sender As Object, e As EventArgs) Handles Movimiento_Salida.Click
        Esconder_tabpages_submenu()
        TabPage12.Parent = TabControl2
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(String.Format("SELECT NOW();"), conn)
            Dim fecha_servidor As DateTime = cmd.ExecuteScalar()
            Fecha_Movimiento.Text = fecha_servidor.ToString("yyyy-MM-dd")
            conn.Close()
        Catch
            MsgBox("No se puede recuperar la fecha de la base de datos", MsgBoxStyle.Exclamation, "Error")
            conn.Close()
        End Try
        Tipo_Movimiento.Text = "SALIDA"
        'Haciendo los controles necesarios para el ingreso visibles
        Datos_Movimientos.Visible = True
        CBN_SolicitudSalida.Visible = True
        CBN_SolicitudSalida.Enabled = True
        CargarCBNSolicitudesSalida()
        Label80.Visible = True
        Tipo_Movimiento.Visible = True
        Label81.Visible = True
        Fecha_Movimiento.Visible = True
        CBN_SolicitudSalida.Visible = True
        Label86.Visible = False
        Label85.Visible = False
        Label84.Visible = False
        Label83.Visible = False
        Label82.Text = "Numero orden de trabajo: *"
        Label82.Visible = True
        Observaciones_Movimiento.Visible = False
        N_Orden_Movimiento.Visible = False
        Monto_Movimiento.Visible = False
        N_Referencia_Movimiento.Visible = False
        N_Referencia_Movimiento.Enabled = False
        Proveedor_Movimiento.Visible = False
        N_Orden_Movimiento.Enabled = False
        N_Orden_Movimiento.Text = ""
        N_Referencia_Movimiento.Text = ""
        Confirmar_Transaccion.Visible = True
    End Sub

    Private Sub CargarCBNSolicitudesSalida()
        With CBN_SolicitudSalida
            Try
                conn.Open()
                Dim query As String = "Select distinct NumeroOrdenTrabajo from solicitud_salida where pendiente = 1"
                Dim cmd As New MySqlCommand(query, conn)
                Dim sqlAdap As New MySqlDataAdapter(cmd)
                Dim dtRecord As New DataTable
                sqlAdap.Fill(dtRecord)
                .DataSource = dtRecord
                .DisplayMember = "NumeroOrdenTrabajo"
                .ValueMember = "NumeroOrdenTrabajo"
                conn.Close()
            Catch ex As Exception
                MsgBox("Error al cargar los perfiles de la base de datos", MsgBoxStyle.Exclamation, "Error")
                conn.Close()
            End Try
        End With
    End Sub

    Private Sub Nueva_Transaccion_Click(sender As Object, e As EventArgs) Handles Nueva_Transaccion.Click
        DataGridView4.Visible = False
        DataGridView4.DataSource = Nothing
        Gru_Movimiento.Visible = False
        Movimiento_Ingreso.Enabled = True
        Movimiento_Salida.Enabled = True
        Generar_Movimientos.Visible = False
        Confirmar_Transaccion.Visible = True
        N_Orden_Movimiento.Enabled = True
        N_Referencia_Movimiento.Enabled = True
        Monto_Movimiento.Enabled = True
        Observaciones_Movimiento.Enabled = True
        Proveedor_Movimiento.Enabled = True
        N_Orden_Movimiento.Text = ""
        N_Referencia_Movimiento.Text = ""
        Monto_Movimiento.Text = ""
        Observaciones_Movimiento.Text = ""
        Proveedor_Movimiento.SelectedItem = -1
        Nueva_Transaccion.Visible = False
        If Tipo_Movimiento.Text = "SALIDA" Then
            Movimiento_Salida_Click(sender, e)
        ElseIf Tipo_Movimiento.Text = "INGRESO" Then
            Movimiento_Ingreso_Click(sender, e)
        End If
    End Sub

    Private Sub Movimiento_Consulta_Click(sender As Object, e As EventArgs) Handles Movimiento_Consulta.Click
        Esconder_tabpages_submenu()
        TabPage18.Parent = TabControl2
        With DataGridView5
            .DataSource = Nothing
            Try
                conn.Open()
                Dim cmd As New MySqlCommand(String.Format("SELECT Cod_Producto as 'Codigo del Producto', Nombre_Producto as 'Producto', Nombre_Categoria as 'Categoria',
                                                            Nombre_SubCategoria as 'Sub-Categoria', Marca, Serie, Stock_Minimo as 'Minimo', Stock_Maximo as 'Maximo',
                                                            stock_Existente as 'Unidades Existentes', compra_Maxima as 'Compra Maxima', Activo
                                                            from productos inner join categorias on productos.id_categoria = categorias.id_categoria inner join categorias_sub on
                                                            productos.Id_SubCategoria = categorias_sub.Id_SubCategoria;"), conn)
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader
                Dim T As New DataTable
                T.Load(reader)
                .DataSource = T
                .ReadOnly = True
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .Visible = True
                conn.Close()
            Catch ex As Exception
                MsgBox("No se pudieron recuperar los datos de los productos de la base de datos:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error.")
                conn.Close()
            End Try
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If FechaInicio.Value > FechaFin.Value Then
            MsgBox("La fecha inicial no puede estar despues de la fecha final", MsgBoxStyle.Exclamation, "Error")
            Exit Sub
        End If

        Dim fecha_inicial As String = FechaInicio.Value.ToString("yyyy-MM-dd")
        Dim fecha_final As String = FechaFin.Value.ToString("yyyy-MM-dd")

        If ComboBox1.SelectedValue = "-1" And ComboBox2.SelectedValue = "-1" Then

            With DataGridView6
                .DataSource = Nothing
                Try
                    conn.Open()
                    Dim query As String = "SELECT orden_movimientos.N_Orden_Compra as 'Orden de Compra #', orden_movimientos.N_Referencia as 'Numero de Remision', 
                                            productos.Nombre_Producto as 'Producto', movimientos.Cantidad, Precio_Compra as 'Precio', movimientos.fecha as' Fecha',proveedores.Nombre_Proveedor as 'Comprado a'
                                            from productos inner join movimientos on movimientos.Id_Producto = productos.Id_Producto
                                            inner join orden_movimientos on orden_movimientos.IdOrden_Movimiento = movimientos.IdOrden_Movimiento
                                            inner join proveedores on orden_movimientos.Nit_Proveedor = proveedores.Nit_Proveedor
                                            where orden_movimientos.Tipo = 'INGRESO' and movimientos.fecha between @Inicio and @Fin;"
                    Dim cmd As New MySqlCommand(query, conn)
                    With cmd.Parameters
                        .AddWithValue("Inicio", fecha_inicial)
                        .AddWithValue("Fin", fecha_final)
                    End With
                    Dim reader As MySqlDataReader
                    reader = cmd.ExecuteReader
                    Dim T As New DataTable
                    T.Load(reader)
                    .DataSource = T
                    .ReadOnly = True
                    .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                    .Visible = True
                    conn.Close()
                Catch ex As Exception
                    MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                    Exit Sub
                End Try
            End With
        ElseIf ComboBox1.SelectedValue <> "-1" And ComboBox2.SelectedValue = "-1" Then

            With DataGridView6
                .DataSource = Nothing
                Try
                    conn.Open()
                    Dim query As String = "SELECT orden_movimientos.N_Orden_Compra as 'Orden de Compra #', orden_movimientos.N_Referencia as 'Numero de Remision', 
                                            productos.Nombre_Producto as 'Producto', movimientos.Cantidad, Precio_Compra as 'Precio', movimientos.fecha as' Fecha',proveedores.Nombre_Proveedor as 'Comprado a'
                                            from productos inner join movimientos on movimientos.Id_Producto = productos.Id_Producto
                                            inner join orden_movimientos on orden_movimientos.IdOrden_Movimiento = movimientos.IdOrden_Movimiento
                                            inner join proveedores on orden_movimientos.Nit_Proveedor = proveedores.Nit_Proveedor
                                            where orden_movimientos.Tipo = 'INGRESO' and movimientos.fecha between @Inicio and @Fin and movimientos.Id_Producto = @IdProd;"
                    Dim cmd As New MySqlCommand(query, conn)
                    With cmd.Parameters
                        .AddWithValue("Inicio", fecha_inicial)
                        .AddWithValue("Fin", fecha_final)
                        .AddWithValue("IdProd", ComboBox1.SelectedValue)
                    End With
                    Dim reader As MySqlDataReader
                    reader = cmd.ExecuteReader
                    Dim T As New DataTable
                    T.Load(reader)
                    .DataSource = T
                    .ReadOnly = True
                    .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                    .Visible = True
                    conn.Close()
                Catch ex As Exception
                    MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                    Exit Sub
                End Try

            End With

        ElseIf ComboBox1.SelectedValue = "-1" And ComboBox2.SelectedValue <> "-1" Then

            With DataGridView6
                .DataSource = Nothing
                Try
                    conn.Open()
                    Dim query As String = "SELECT orden_movimientos.N_Orden_Compra as 'Orden de Compra #', orden_movimientos.N_Referencia as 'Numero de Remision', 
                                            productos.Nombre_Producto as 'Producto', movimientos.Cantidad, Precio_Compra as 'Precio', movimientos.fecha as' Fecha',proveedores.Nombre_Proveedor as 'Comprado a'
                                            from productos inner join movimientos on movimientos.Id_Producto = productos.Id_Producto
                                            inner join orden_movimientos on orden_movimientos.IdOrden_Movimiento = movimientos.IdOrden_Movimiento
                                            inner join proveedores on orden_movimientos.Nit_Proveedor = proveedores.Nit_Proveedor
                                            where orden_movimientos.Tipo = 'INGRESO' and movimientos.fecha between @Inicio and @Fin and orden_movimientos.Nit_Proveedor = @Nit;"
                    Dim cmd As New MySqlCommand(query, conn)
                    With cmd.Parameters
                        .AddWithValue("Inicio", fecha_inicial)
                        .AddWithValue("Fin", fecha_final)
                        .AddWithValue("Nit", ComboBox2.SelectedValue)
                    End With
                    Dim reader As MySqlDataReader
                    reader = cmd.ExecuteReader
                    Dim T As New DataTable
                    T.Load(reader)
                    .DataSource = T
                    .ReadOnly = True
                    .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                    .Visible = True
                    conn.Close()
                Catch ex As Exception
                    MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                    Exit Sub
                End Try

            End With
        ElseIf ComboBox1.SelectedValue <> "-1" And ComboBox2.SelectedValue <> "-1" Then

            With DataGridView6
                .DataSource = Nothing
                Try
                    conn.Open()
                    Dim query As String = "SELECT orden_movimientos.N_Orden_Compra as 'Orden de Compra #', orden_movimientos.N_Referencia as 'Numero de Remision', 
                                            productos.Nombre_Producto as 'Producto', movimientos.Cantidad, Precio_Compra as 'Precio', movimientos.fecha as' Fecha',proveedores.Nombre_Proveedor as 'Comprado a'
                                            from productos inner join movimientos on movimientos.Id_Producto = productos.Id_Producto
                                            inner join orden_movimientos on orden_movimientos.IdOrden_Movimiento = movimientos.IdOrden_Movimiento
                                            inner join proveedores on orden_movimientos.Nit_Proveedor = proveedores.Nit_Proveedor
                                            where orden_movimientos.Tipo = 'INGRESO' and movimientos.fecha between @Inicio and @Fin and movimientos.Id_Producto = @IdProd and orden_movimientos.Nit_Proveedor = @Nit;"
                    Dim cmd As New MySqlCommand(query, conn)
                    With cmd.Parameters
                        .AddWithValue("Inicio", fecha_inicial)
                        .AddWithValue("Fin", fecha_final)
                        .AddWithValue("IdProd", ComboBox1.SelectedValue)
                        .AddWithValue("Nit", ComboBox2.SelectedValue)
                    End With
                    Dim reader As MySqlDataReader
                    reader = cmd.ExecuteReader
                    Dim T As New DataTable
                    T.Load(reader)
                    .DataSource = T
                    .ReadOnly = True
                    .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                    .Visible = True
                    conn.Close()
                Catch ex As Exception
                    MsgBox("No se pudieron recuperar los productos de la solicitud:" & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Error")
                    conn.Close()
                    Exit Sub
                End Try

            End With
        End If
        Dim gasto As Decimal = 0
        For row As Integer = 0 To DataGridView6.Rows.Count - 1
            gasto = gasto + DataGridView6(3, row).Value * DataGridView6(4, row).Value
        Next

        MsgBox("Gasto total: " & Format(gasto, "Currency"))

    End Sub

    '    Private Sub Generar_Movimientos_Click(sender As Object, e As EventArgs) Handles Generar_Movimientos.Click
    '        DataGridView4.Visible = True
    '        Dim consulta_movimientos As String = ""
    '        Dim stock As Integer
    '        If IsNumeric(Cantidad_Movimiento.Text) Then
    '            Cantidad_a_mover = Cantidad_Movimiento.Text
    '            Tipo_Movi = Tipo_Movimiento.Text
    '            Try
    '                conn.Open()
    '                Dim Comando As New MySqlCommand("SELECT * FROM stock " &
    '                        "INNER JOIN productos " &
    '                        "ON (stock.Id_Stock = productos.Id_Stock) " &
    '                        "WHERE Id_Producto= @IdProd;", conn)
    '                Comando.Parameters.AddWithValue("IdProd", Producto_Movimiento.SelectedValue)
    '                Dim Adaptador As New MySqlDataAdapter(Comando)
    '                Dim Tabla As New DataTable
    '                Adaptador.Fill(Tabla)
    '                conn.Close()
    '                Tabla1 = Tabla
    '            Catch ex As Exception
    '                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error.")
    '            End Try
    '            'DataGridView4.DataSource = Tabla1
    '            id_stock = Tabla1.Rows(0).ItemArray(0).ToString
    '            If Tipo_Movimiento.Text = "SALIDA" Then
    '                'If Descripcion_Movimiento.Text = "" Or Producto_Movimiento.Text = "" Or Usuario_Movimiento.Text = "" Then
    '                '    MsgBox("Faltan algunos datos para generar el movimiento")
    '                '    GoTo err
    '                'End If
    '                'If Convert.Toint64(Tabla1.Rows(0).ItemArray(3).ToString) < Convert.Toint64(Cantidad_Movimiento.Text) Then
    '                '    MsgBox("Atencion NO puede realizar el movimiento, porque no posee esa cantidad en su inventario")
    '                '    GoTo err
    '                'End If
    '                'stock = Convert.Toint64(Tabla1.Rows(0).ItemArray(3).ToString) - Convert.Toint64(Cantidad_Movimiento.Text)
    '            Else
    '                If Descripcion_Movimiento.Text = "" Or Producto_Movimiento.Text = "" Or Precio_Movimiento.Text = "" Then
    '                    MsgBox("Todos los datos son obligatorios")
    '                    Exit Sub
    '                End If
    '                If Tabla1.Rows(0).ItemArray(1).ToString <> 0 Then
    '                    stock = Convert.Toint64(Tabla1.Rows(0).ItemArray(3).ToString) + Convert.Toint64(Cantidad_Movimiento.Text)
    '                    If Convert.Toint64(Tabla1.Rows(0).ItemArray(2).ToString) < stock) Then
    '                        MsgBox("Atencion NO puede realizar el movimiento, porque su movimiento supera el máximo permitido")
    '                        Exit Sub
    '                    End If
    '                ElseIf Tabla1.Rows(0).ItemArray(1).ToString = 0 Then
    '                    stock = Convert.Toint64(Tabla1.Rows(0).ItemArray(3).ToString) + Convert.Toint64(Cantidad_Movimiento.Text)
    '                End If
    '            End If
    '            Using conexion As New MySqlConnection(datasource)
    '                Dim Comando As New MySqlCommand("SELECT Estante,Entrepano,Caja_Color,Zona,Cantidad,Aforo FROM rel_ubicaciones_productos " &
    '                            "INNER JOIN productos " &
    '                            "ON (rel_ubicaciones_productos.Id_Producto = productos.Id_Producto) " &
    '                            "INNER JOIN ubicaciones " &
    '                            "ON (rel_ubicaciones_productos.Id_Ubicacion = ubicaciones.Id_Ubicacion) " &
    '                            "WHERE rel_ubicaciones_productos.Id_Producto='" & Producto_Movimiento.SelectedValue & "'", conexion)
    '                Dim Adaptador As New MySqlDataAdapter(Comando)
    '                Dim Tabla As New DataTable
    '                Try
    '                    Adaptador.Fill(Tabla)
    '                Catch ex As Exception
    '                Finally
    '                    If conexion.State = ConnectionState.Open Then
    '                        conexion.Close()
    '                    End If
    '                End Try
    '                Tabla1 = Tabla
    '            End Using
    '            If Tabla1.Rows.Count >= 1 Then
    '                Bandera_Rel = 4
    '                Consulta_rel = "SELECT Estante,Entrepano,Caja_Color,Zona,Cantidad,Aforo,IdRel_Ubicaciones_Productos FROM rel_ubicaciones_productos " &
    '                "INNER JOIN productos " &
    '                "ON (rel_ubicaciones_productos.Id_Producto = productos.Id_Producto) " &
    '                "INNER JOIN ubicaciones " &
    '                "ON (rel_ubicaciones_productos.Id_Ubicacion = ubicaciones.Id_Ubicacion) " &
    '                "WHERE rel_ubicaciones_productos.Id_Producto='" & Producto_Movimiento.SelectedValue & "'"
    '                Id_elem = Producto_Movimiento.SelectedValue
    '                'Tabla_Rel = "Ubicaciones"
    '                'Elemento_rel = Nombre_Producto.Text
    '                Try
    '                    Using conn As New MySqlConnection(datasource)
    '                        conn.Open()
    '                        Dim cmd As New MySqlCommand("UPDATE stock SET Stock_Existente = '" & stock &
    '                        "' WHERE Id_Stock = '" & id_stock & "'", conn)
    '                        cmd.ExecuteNonQuery()
    '                        'MessageBox.Show("Registro MODIFICADO")
    '                        conn.Close()
    '                    End Using
    '                Catch ex As Exception
    '                    MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
    '                End Try
    '                Nueva_Transaccion.Visible = True
    '                If Tipo_Movimiento.Text = "SALIDA" Then
    '                    Try
    '                        Using conn As New MySqlConnection(datasource)
    '                            conn.Open()
    '                            Dim cmd As New MySqlCommand("INSERT INTO movimientos (IdOrden_Movimiento, Id_Producto, Cantidad, Id_Usuario, Fecha, Descripcion, Id_Usuario_Final)" &
    '                                "VALUES ('" & UCase(Id_Ord_Movimiento) & "', '" & Producto_Movimiento.SelectedValue & "', '" & Cantidad_Movimiento.Text &
    '                                "', '" & id_Usuar_Per & "', '" & Fecha_Movimiento.Text & "', '" & UCase(Descripcion_Movimiento.Text) & "', '" & Usuario_Movimiento.SelectedValue & "')", conn)
    '                            cmd.ExecuteNonQuery()
    '                            'MessageBox.Show("Registro MODIFICADO")
    '                            conn.Close()
    '                        End Using
    '                    Catch ex As Exception
    '                        MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
    '                    End Try
    '                Else
    '                    Try
    '                        Using conn As New MySqlConnection(datasource)
    '                            conn.Open()
    '                            Dim cmd As New MySqlCommand("INSERT INTO movimientos (IdOrden_Movimiento, Id_Producto, Cantidad, Precio_Compra, Id_Usuario, Fecha, Descripcion)" &
    '                                "VALUES ('" & UCase(Id_Ord_Movimiento) & "', '" & Producto_Movimiento.SelectedValue & "', '" & Cantidad_Movimiento.Text &
    '                                "', '" & Precio_Movimiento.Text & "', '" & id_Usuar_Per & "', '" & Fecha_Movimiento.Text & "', '" & UCase(Descripcion_Movimiento.Text) & "')", conn)
    '                            cmd.ExecuteNonQuery()
    '                            'MessageBox.Show("Registro MODIFICADO")
    '                            conn.Close()
    '                        End Using
    '                    Catch ex As Exception
    '                        MsgBox("El registro no pudo Modificarse por: " & vbCrLf & ex.Message)
    '                    End Try
    '                End If

    '                If Tipo_Movimiento.Text = "SALIDA" Then
    '                    consulta_movimientos = "SELECT Nombre_Producto,Cantidad,Descripcion,Nombre_Usuario FROM movimientos " & 'Nombre_Producto,Cantidad,Precio_Compra,Descripcion,Nombre_Usuario
    '                        "INNER JOIN productos " &
    '                        "ON (movimientos.Id_Producto = productos.Id_Producto) " &
    '                        "INNER JOIN usuarios " &
    '                        "ON (movimientos.Id_Usuario_Final = usuarios.Id_Usuario) " &
    '                        "WHERE IdOrden_Movimiento='" & Id_Ord_Movimiento & "'"
    '                Else
    '                    consulta_movimientos = "SELECT Nombre_Producto,Cantidad,Precio_Compra,Descripcion FROM movimientos " & 'Nombre_Producto,Cantidad,Precio_Compra,Descripcion,Nombre_Usuario
    '                                "INNER JOIN productos " &
    '                                "ON (movimientos.Id_Producto = productos.Id_Producto) " &
    '                                "WHERE IdOrden_Movimiento='" & Id_Ord_Movimiento & "'"
    '                End If
    '                Using conexion As New MySqlConnection(datasource)
    '                    Dim Comando As New MySqlCommand(consulta_movimientos, conexion)
    '                    Dim Adaptador As New MySqlDataAdapter(Comando)
    '                    Dim Tabla As New DataTable
    '                    Try
    '                        Adaptador.Fill(Tabla)
    '                    Catch ex As Exception
    '                    Finally
    '                        If conexion.State = ConnectionState.Open Then
    '                            conexion.Close()
    '                        End If
    '                    End Try
    '                    DataGridView4.DataSource = Tabla
    '                End Using
    '                Descripcion_Movimiento.Text = ""
    '                Precio_Movimiento.Text = ""
    '                Cantidad_Movimiento.Text = ""
    '                Form3.ShowDialog()
    '            Else
    '                MsgBox("Atencion NO puede realizar el movimiento, porque el producto no cuenta con ubicaciones")
    '                GoTo err
    '            End If
    '        Else
    '            MsgBox("El campo de cantidad solo acepta valores numéricos")
    '        End If
    '    End Sub


End Class
