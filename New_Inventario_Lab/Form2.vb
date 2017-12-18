Imports MySql.Data.MySqlClient
Imports System.Configuration
Imports System.Text
Imports System.Security.Cryptography


Module Module1

    Public Usuario_Perfil = ""
    Public Bandera_Rel As Integer = 0
    Public Consulta_rel = ""
    Public Tabla_Rel = "" ' Tabla del origen de la relacion
    Public Elemento_rel = "" ' Nombre elemento que deseo relacionar
    Public Id_elem = "" ' Codigo/Id elemento que deseo relacionar
    Public Cantidad_a_mover
    Public Tipo_Movi
    Public id_stock
    Public id_Usuar_Per

End Module

Public Class FormIngreso
    Dim conn As New MySqlConnection
    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        Dim usuario As String = TxtBxUsuario.Text
        Dim contraseña As String = TxtBxContraseña.Text
        Dim bd_password As String
        Dim salt As String
        Try
            conn.Open()
            Dim cmd As New MySqlCommand(String.Format("Select hash from usuarios where binary usuario = @usuario;"), conn)
            cmd.Parameters.AddWithValue("usuario", usuario)
            bd_password = Convert.ToString(cmd.ExecuteScalar())
            Dim cmd2 As New MySqlCommand(String.Format("SELECT Salt from usuarios where binary usuario = @usuario;"), conn)
            cmd2.Parameters.AddWithValue("usuario", usuario)
            salt = Convert.ToString(cmd2.ExecuteScalar())
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
            conn.Close()
            Exit Sub
        End Try

        If bd_password = Nothing Then
            MsgBox("Credenciales de inicio de sesion no validas")
            Exit Sub
        End If

        contraseña = salt + contraseña

        contraseña = ComputeHashOfString(Of SHA256CryptoServiceProvider)(contraseña)

        If contraseña = bd_password Then
            MsgBox("Bienvenido " & usuario & "", False, "Log-In")
            Dim usuario_id As String
            Dim id_perfil As String
            Try
                conn.Open()
                Dim cmd As New MySqlCommand(String.Format("Select Id_Usuario from usuarios where usuario = @usuario;"), conn)
                Dim cmd2 As New MySqlCommand(String.Format("Select Nombre_Usuario from usuarios where usuario = @usuario;"), conn)
                Dim cmd3 As New MySqlCommand(String.Format("Select Id_perfil from usuarios where usuario = @usuario; "), conn)
                cmd.Parameters.AddWithValue("usuario", usuario)
                cmd2.Parameters.AddWithValue("usuario", usuario)
                cmd3.Parameters.AddWithValue("usuario", usuario)
                usuario_id = Convert.ToString(cmd.ExecuteScalar())
                Dim nombre As String = Convert.ToString(cmd2.ExecuteScalar())
                id_perfil = Convert.ToString(cmd3.ExecuteScalar())
                Dim cmd4 As New MySqlCommand(String.Format("Select Nombre_perfil from perfiles where Id_Perfil = @IDPerfil;"), conn)
                cmd4.Parameters.AddWithValue("IDPerfil", id_perfil)
                Usuario_Perfil = UCase(usuario) & " / " & Convert.ToString(cmd4.ExecuteScalar())
                Form1.Label1.Text = Usuario_Perfil
                id_Usuar_Per = usuario_id
                conn.Close()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
                conn.Close()
                Exit Sub
            End Try
            Me.Hide()
            Form1.Show()
        Else
            MsgBox("Credenciales de inicio de sesion no validas")
            TxtBxContraseña.Text = ""
            Exit Sub
        End If
    End Sub

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

    Private Sub FormIngreso_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Connect()
        With TxtBxUsuario
            .Clear()
            .Focus()
        End With
        With TxtBxContraseña
            .Clear()
        End With
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

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub
End Class
