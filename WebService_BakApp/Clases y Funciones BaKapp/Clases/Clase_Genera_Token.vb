Imports System.Data.SqlClient
Imports System.IdentityModel.Tokens.Jwt
Imports System.Security.Claims
Imports System.Text
Imports Microsoft.IdentityModel.Tokens

Public Class TokenGenerator

    ' Clave secreta para firmar el token (debe mantenerse en secreto)
    Private Shared _SecretKey As String = "tu_clave_secreta_aqui"
    Public Property _BaseConectada As Boolean

    Public Sub New()

        Dim _Sql As New Class_SQL()

        _Global_BaseBk = _Sql.Fx_Trae_Dato("TABCARAC", "NOKOCARAC", "KOTABLA = 'BAKAPP'").ToString.Trim & ".dbo."

        Consulta_sql = "Select Top 1 Rut,Llave1,Llave2,Llave3,Llave4" & vbCrLf &
                       "From " & _Global_BaseBk & "Zw_Licencia Where Activa = 1"

        Dim _Row_Licencia As DataRow = _Sql.Fx_Get_DataRow(Consulta_sql)

        Dim _Llave1 As String = _Row_Licencia.Item("Llave1").ToString.Trim
        Dim _Llave2 As String = _Row_Licencia.Item("Llave2").ToString.Trim
        Dim _Llave3 As String = _Row_Licencia.Item("Llave3").ToString.Trim
        Dim _Llave4 As String = _Row_Licencia.Item("Llave4").ToString.Trim

        _SecretKey = _Llave1 & _Llave2 & _Llave3 & _Llave4

    End Sub

    ' Genera un token de validación con la información del usuario
    Function GenerateToken(username As String, expirationMinutes As Integer) As String

        Dim securityKey As New SymmetricSecurityKey(Encoding.UTF8.GetBytes(_SecretKey))
        Dim credentials As New SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256)

        Dim claims As New List(Of Claim)()
        claims.Add(New Claim(ClaimTypes.Name, username))

        Dim token As New JwtSecurityToken(
            issuer:=Nothing,
            audience:=Nothing,
            claims:=claims,
            expires:=DateTime.UtcNow.AddMinutes(expirationMinutes),
            signingCredentials:=credentials
        )

        Dim tokenHandler As New JwtSecurityTokenHandler()

        Return tokenHandler.WriteToken(token)
    End Function

    ' Verifica y decodifica un token de validación
    Function ValidateToken(token As String) As ClaimsPrincipal
        Dim tokenHandler As New JwtSecurityTokenHandler()
        Dim securityToken As SecurityToken

        Dim validationParameters As New TokenValidationParameters() With {
            .ValidateIssuer = False, ' Puedes establecer esto en True si quieres validar el emisor
            .ValidateAudience = False, ' Puedes establecer esto en True si quieres validar la audiencia
            .ClockSkew = TimeSpan.Zero,
            .RequireSignedTokens = True,
            .ValidateLifetime = True,
            .ValidateIssuerSigningKey = True,
            .IssuerSigningKey = New SymmetricSecurityKey(Encoding.UTF8.GetBytes(_SecretKey))
        }

        Try
            Dim principal = tokenHandler.ValidateToken(token, validationParameters, securityToken)
            Return principal
        Catch ex As Exception
            Return Nothing ' Token inválido o expirado
        End Try
    End Function

End Class
