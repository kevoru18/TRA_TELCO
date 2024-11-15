
Dim conn As Object
Dim rs As Object
Dim query As String
Dim dbName As String

' Procedimiento para llenar el ComboBox con los nombres de las bases de datos
Sub RellenarComboBoxBasesDeDatos()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hoja1")
    
    Dim cbo As Object
    Set cbo = ws.OLEObjects("ComboBox2").Object ' Referencia al ComboBox
    
    ' Conectar a SQL Server
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLOLEDB; Data Source=192.168.201.12; Initial Catalog=master; User ID=sa; Password=infinity;"
    
    ' Consulta para obtener los nombres de las bases de datos
    query = "SELECT name FROM sys.databases WHERE state = 0;" ' Excluye bases de datos que están offline
    
    ' Ejecutar la consulta
    Set rs = conn.Execute(query)
    
    ' Limpiar el ComboBox antes de llenarlo
    cbo.Clear
    
    ' Llenar el ComboBox con los nombres de las bases de datos
    Do While Not rs.EOF
        cbo.AddItem rs.Fields("name").Value
        rs.MoveNext
    Loop
    
    ' Cerrar la conexión y liberar recursos
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub

' Evento cuando se selecciona una base de datos del ComboBox
Private Sub ComboBox2_Change()
    dbName = ThisWorkbook.Sheets("Hoja1").OLEObjects("ComboBox2").Object.Value ' Obtener el valor seleccionado
    RellenarGraficosConDatos dbName
End Sub

Private Sub CommandButton1_Click()
    ' Obtener el valor seleccionado del ComboBox2
    
    MsgBox (entra)
    
    Dim dbName As String
    dbName = ThisWorkbook.Sheets("Hoja1").OLEObjects("ComboBox2").Object.Value

    ' Llamar al procedimiento para rellenar los gráficos con los datos de la base de datos seleccionada
    If dbName <> "" Then
        RellenarGraficosConDatos dbName
    Else
        MsgBox "Por favor, seleccione una base de datos.", vbExclamation
    End If
End Sub
Sub RellenarGraficosConDatos(dbName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Hoja1")
    Dim chart1 As ChartObject
    Dim chart2 As ChartObject
    Dim chart3 As ChartObject
    Dim chart4 As ChartObject ' Gráfico de pastel
    Dim conn As Object
    Dim rs As Object
    
    ' Variables para almacenar los datos
    Dim TotalRecords As Long
    Dim CompleteRecords As Long
    Dim CorrectionRecords As Long
    Dim TotalEmpresas As Long
    Dim Contactadas As Long
    
    ' Ejemplo de conexión a la base de datos seleccionada
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLOLEDB; Data Source=192.168.201.12; Initial Catalog=" & dbName & "; User ID=sa; Password=infinity;"
    
    ' Consulta SQL para los KPIs
    Dim query As String
    query = "SELECT TotalRecords, CompleteRecords, CorrectionRecords FROM CNAE_KPI_Audit"
    Set rs = conn.Execute(query)

    ' Inicializar variables
    TotalRecords = 0
    CompleteRecords = 0
    CorrectionRecords = 0

    If Not (rs.EOF And rs.BOF) Then
        ' Obtener los datos de la primera fila
        TotalRecords = rs.Fields("TotalRecords").Value
        CompleteRecords = rs.Fields("CompleteRecords").Value
        CorrectionRecords = rs.Fields("CorrectionRecords").Value
    Else
        MsgBox "No se encontraron registros en la tabla CNAE_KPI_Audit."
    End If

    ' Cerrar la conexión
    rs.Close
    conn.Close
    Set rs = Nothing
    
    ' Crear arrays para los gráficos
    Dim data1(1 To 2, 1 To 2) As Variant ' Para gráfico 1
    Dim data2(1 To 2, 1 To 2) As Variant ' Para gráfico 2
    Dim data3() As Variant ' Array dinámico para el gráfico 3

    ' Rellenar arrays con los datos recuperados
    data1(1, 1) = "Total"
    data1(1, 2) = TotalRecords
    data1(2, 1) = "Completos"
    data1(2, 2) = CompleteRecords
    
    data2(1, 1) = "Modificados"
    data2(1, 2) = CorrectionRecords
    data2(2, 1) = "Completos"
    data2(2, 2) = CompleteRecords

    ' Ahora vamos a llenar los datos para Chart3 a partir de la consulta de clics
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=SQLOLEDB; Data Source=192.168.201.12; Initial Catalog=" & dbName & "; User ID=sa; Password=infinity;"
    
    ' Consulta SQL para obtener datos del log de clics por fecha
    query = "SELECT CAST(ClickDate AS DATE) AS ClickDate, ButtonName, COUNT(*) AS ClickCount FROM LinkedinClickLog GROUP BY CAST(ClickDate AS DATE), ButtonName ORDER BY ClickDate"
    Set rs = conn.Execute(query)

    Dim rowCount As Long
    rowCount = 0
    Dim i As Long

    If Not (rs.EOF And rs.BOF) Then
        ' Contar cuántas filas tiene el resultado de la consulta
        Do While Not rs.EOF
            rowCount = rowCount + 1
            rs.MoveNext
        Loop
        rs.MoveFirst

        ' Redimensionar el array data3 para almacenar los resultados
        ReDim data3(1 To rowCount, 1 To 3) ' Aumentar a 3 columnas para incluir fecha

        ' Rellenar el array con los resultados de la consulta
        i = 1
        Do While Not rs.EOF
            data3(i, 1) = rs.Fields("ClickDate").Value ' Fecha del clic
            data3(i, 2) = rs.Fields("ButtonName").Value ' Nombre del botón
            data3(i, 3) = rs.Fields("ClickCount").Value ' Conteo de clics
            i = i + 1
            rs.MoveNext
        Loop
    Else
        MsgBox "No se encontraron registros en el log de clics."
    End If
    
    ' Cerrar la conexión
    rs.Close
    conn.Close
    Set rs = Nothing
    
    ' Asignar los datos directamente a los gráficos
    Set chart1 = ws.ChartObjects("Chart1")
    chart1.Chart.SeriesCollection.NewSeries
    chart1.Chart.SeriesCollection(1).XValues = Array(data1(1, 1), data1(2, 1))
    chart1.Chart.SeriesCollection(1).Values = Array(data1(1, 2), data1(2, 2))
    
    Set chart2 = ws.ChartObjects("Chart2")
    chart2.Chart.SeriesCollection.NewSeries
    chart2.Chart.SeriesCollection(1).XValues = Array(data2(1, 1), data2(2, 1))
    chart2.Chart.SeriesCollection(1).Values = Array(data2(1, 2), data2(2, 2))
    
    ' Crear y rellenar Chart3 como gráfico de columnas
    Set chart3 = ws.ChartObjects("Chart3")
    chart3.Chart.ChartType = xlColumnClustered ' Cambiar a gráfico de columnas

    ' Limpiar series existentes antes de agregar nuevas series
    On Error Resume Next ' Evitar errores si no hay series
    Do While chart3.Chart.SeriesCollection.count > 0
        chart3.Chart.SeriesCollection(1).Delete
    Loop
    On Error GoTo 0 ' Restablecer el manejo de errores

    ' Crear una colección para los nombres de los botones
    Dim buttonNames As Collection
    Set buttonNames = New Collection

    ' Recorre los datos para agregar series
    For i = 1 To rowCount
        On Error Resume Next
        buttonNames.Add data3(i, 2), CStr(data3(i, 2)) ' Añadir nombre del botón a la colección si no existe
        On Error GoTo 0
    Next i

    ' Crear series para cada botón
    Dim dateArray() As Variant
    Dim clickCountArray() As Variant
    Dim buttonName As Variant

    For Each buttonName In buttonNames
        ' Inicializar arrays para los valores
        Dim count As Long
        count = 0
        ReDim dateArray(1 To rowCount)
        ReDim clickCountArray(1 To rowCount)

        ' Rellenar los arrays con los datos correspondientes
        For i = 1 To rowCount
            If data3(i, 2) = buttonName Then
                count = count + 1
                dateArray(count) = data3(i, 1)
                clickCountArray(count) = data3(i, 3)
            End If
        Next i

        ' Ajustar el tamaño de los arrays a la cantidad de elementos
        If count > 0 Then
            ReDim Preserve dateArray(1 To count)
            ReDim Preserve clickCountArray(1 To count)

            ' Crear la serie en el gráfico
            With chart3.Chart.SeriesCollection.NewSeries
                .Name = buttonName
                .XValues = dateArray
                .Values = clickCountArray
            End With
        End If
    Next buttonName

    ' Configuración del gráfico 3
    With chart3.Chart
        .HasTitle = True
        .ChartTitle.Text = "Clics por Botón a lo Largo del Mes"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Fecha"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Número de Clics"
        If .SeriesCollection.count > 0 Then
            .SeriesCollection(1).HasDataLabels = True ' Mostrar etiquetas de datos para la primera serie
        End If
    End With


    ' --- Gráfico 4 (Pie Chart) ---
    ' Simular la obtención de datos para gráfico 4
    ' Este código asume que Contactadas y TotalEmpresas están definidos previamente.
    ' Agrega tu lógica para obtener estos datos según sea necesario.

      ' Crear array para el grï¿½fico 4
    Dim data4(1 To 2, 1 To 2) As Variant
    data4(1, 1) = "Relleno"
    data4(1, 2) = Contactadas
    data4(2, 1) = "Vacio"
    data4(2, 2) = TotalEmpresas - Contactadas
    
    ' Crear y rellenar Chart4 como grï¿½fico de pastel
    'Set chart4 = ws.ChartObjects.Add(Left:=500, Top:=100, Width:=300, Height:=200)
    'chart4.Name = "Chart4"
    Set chart4 = ws.ChartObjects("Chart4")
    chart4.Chart.ChartType = xlPie
    
    ' Asignar los datos al grï¿½fico de pastel
    With chart4.Chart
        .SetSourceData ws.Range("A1:B2") ' Asignar rango que contiene los datos
        .SeriesCollection(1).XValues = Array(data4(1, 1), data4(2, 1))
        .SeriesCollection(1).Values = Array(data4(1, 2), data4(2, 2))
        .HasTitle = True
        .ChartTitle.Text = "PERSONA CONTACTO OBTENIDAS"
        .SeriesCollection(1).HasDataLabels = True ' Mostrar etiquetas de datos
    End With

    ' Mensaje de finalización
    MsgBox "Gráficos actualizados correctamente."
End Sub

