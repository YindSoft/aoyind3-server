Attribute VB_Name = "modCuentas"
'El Yind 26/08/2010
Option Explicit

Public Function CuentaExiste(ByVal Name As String) As Boolean

CuentaExiste = GetByCampo("SELECT COUNT(Id) as 'Cantidad' FROM cuentas WHERE Nombre=" & Comillas(Name), "Cantidad") = "1"

End Function
Public Function BANCuentaCheck(ByVal Name As String) As Boolean

BANCuentaCheck = CStrNull(GetByCampo("SELECT Ban FROM cuentas WHERE Nombre=" & Comillas(Name), "Ban")) = "1"

End Function

