Sub KMeansClustering()
    Dim ws As Worksheet
    Dim dataRange As Range, centerRange As Range
    Dim clusterAssigned As Range
    Dim i As Long, j As Long, k As Long
    Dim iteration As Long, maxIterations As Long
    Dim changed As Boolean
    
    Dim distances() As Double
    Dim totalMotifs() As Double
    Dim counts() As Long
    Dim numClusters As Long
    
    ' Configuración inicial
    Set ws = ThisWorkbook.Sheets(1) ' Asegúrate de que estás en la hoja correcta
    Set dataRange = ws.Range("B2:L181")        ' Rango de datos (motivos por sentencia)
    Set centerRange = ws.Range("B184:L188")    ' Rango de centros de clusters (5 filas = 5 clusters)
    Set clusterAssigned = ws.Range("M2:M181")  ' Rango donde se asignan los clusters
    
    ' Número de clusters = número de filas de centerRange
    numClusters = centerRange.Rows.Count
    
    ' Dimensionar arrays dinámicos según el número de clusters y columnas
    ReDim distances(1 To numClusters) As Double
    ReDim totalMotifs(1 To numClusters, 1 To dataRange.Columns.Count) As Double
    ReDim counts(1 To numClusters) As Long

    maxIterations = 100 ' Máximo número de iteraciones
    iteration = 0
    changed = True

    ' Iterar hasta que las asignaciones no cambien o alcancemos el máximo de iteraciones
    Do While changed And iteration < maxIterations
        iteration = iteration + 1
        changed = False
         
        ' Paso 1: Calcular distancias y asignar clusters
        For i = 1 To dataRange.Rows.Count
            ' Calcular la distancia a cada centro
            For j = 1 To numClusters
                distances(j) = 0
                For k = 1 To dataRange.Columns.Count
                    distances(j) = distances(j) + (dataRange.Cells(i, k).Value - centerRange.Cells(j, k).Value) ^ 2
                Next k
                distances(j) = Sqr(distances(j)) ' Tomar la raíz cuadrada
            Next j
             
            ' Asignar el cluster más cercano
            Dim minDistance As Double, assignedCluster As Long
            minDistance = distances(1)
            assignedCluster = 1
            For j = 2 To numClusters
                If distances(j) < minDistance Then
                    minDistance = distances(j)
                    assignedCluster = j
                End If
            Next j
             
            ' Verificar si la asignación cambió
            If clusterAssigned.Cells(i, 1).Value <> "C" & assignedCluster Then
                clusterAssigned.Cells(i, 1).Value = "C" & assignedCluster
                changed = True
            End If
        Next i
         
        ' Paso 2: Recalcular los centros
        ' Reiniciar acumuladores
        For j = 1 To numClusters
            counts(j) = 0
            For k = 1 To dataRange.Columns.Count
                totalMotifs(j, k) = 0
            Next k
        Next j
         
        ' Acumular valores para cada cluster
        For i = 1 To dataRange.Rows.Count
            assignedCluster = Val(Mid(clusterAssigned.Cells(i, 1).Value, 2)) ' Quita la "C" y lo pasa a número
            If assignedCluster >= 1 And assignedCluster <= numClusters Then
                For k = 1 To dataRange.Columns.Count
                    totalMotifs(assignedCluster, k) = totalMotifs(assignedCluster, k) + dataRange.Cells(i, k).Value
                Next k
                counts(assignedCluster) = counts(assignedCluster) + 1
            End If
        Next i
         
        ' Calcular los nuevos promedios (nuevos centros)
        For j = 1 To numClusters
            For k = 1 To dataRange.Columns.Count
                If counts(j) > 0 Then
                    centerRange.Cells(j, k).Value = totalMotifs(j, k) / counts(j)
                End If
            Next k
        Next j
    Loop

    ' Finalizar
    MsgBox "Clustering completado en " & iteration & " iteraciones.", vbInformation
End Sub

