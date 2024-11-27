On Error Resume Next

Dim conn
Dim rs
Dim query
Dim dbName

' Obtener el argumento (nombre de la base de datos)
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    WScript.Echo "No se proporcionó el nombre de la base de datos."
    WScript.Quit 1
End If

dbName = objArgs(0)
WScript.Echo "Nombre de la base de datos recibido: " & dbName

' Conexión a la base de datos seleccionada
Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=SQLOLEDB; Data Source=192.168.201.12; Initial Catalog=" & dbName & "; User ID=sa; Password=infinity;"
If Err.Number <> 0 Then
    WScript.Echo "Error al conectar a la base de datos: " & Err.Description
    WScript.Quit 1
End If

WScript.Echo "Conexión a la base de datos establecida."

' Consulta SQL para los KPIs
query = "SELECT TotalRecords, CompleteRecords, CorrectionRecords FROM CNAE_KPI_Audit"
Set rs = conn.Execute(query)
If Err.Number <> 0 Then
    WScript.Echo "Error al ejecutar la consulta SQL: " & Err.Description
    WScript.Quit 1
End If

WScript.Echo "Consulta SQL ejecutada correctamente."

' Inicializar variables
Dim TotalRecords, CompleteRecords, CorrectionRecords
TotalRecords = 0
CompleteRecords = 0
CorrectionRecords = 0

If Not (rs.EOF And rs.BOF) Then
    ' Obtener los datos de la primera fila
    TotalRecords = rs.Fields("TotalRecords").Value
    CompleteRecords = rs.Fields("CompleteRecords").Value
    CorrectionRecords = rs.Fields("CorrectionRecords").Value
    WScript.Echo "Datos obtenidos: TotalRecords=" & TotalRecords & ", CompleteRecords=" & CompleteRecords & ", CorrectionRecords=" & CorrectionRecords
Else
    WScript.Echo "No se encontraron registros en la tabla CNAE_KPI_Audit."
End If

' Cerrar la conexión
rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing

WScript.Echo "Conexión cerrada y recursos liberados."

' Continuar con la lógica de los gráficos...
' Aquí puedes agregar el código para actualizar los gráficos en la hoja de Excel