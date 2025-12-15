Option Explicit

Sub KMeansElbowMethod()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim dataRange As Range
    Set dataRange = ws.Range("B2:L181")
    
    Dim maxK As Integer: maxK = 20
    Dim SSEs() As Double
    ReDim SSEs(1 To maxK)
    
    Dim k As Integer
    For k = 1 To maxK
        SSEs(k) = RunKMeans(ws, dataRange, k)
    Next k
    
    ' Volcar resultados SSE en hoja
    Dim outputStart As Range
    Set outputStart = ws.Range("N2")
    
    outputStart.Offset(0, 0).Value = "K"
    outputStart.Offset(0, 1).Value = "SSE"
    
    For k = 1 To maxK
        outputStart.Offset(k, 0).Value = k
        outputStart.Offset(k, 1).Value = SSEs(k)
    Next k
    
    ' Crear gráfico del codo
    Dim chartObj As ChartObject
    On Error Resume Next
    ws.ChartObjects("ElbowChart").Delete
    On Error GoTo 0
    
    Set chartObj = ws.ChartObjects.Add(Left:=outputStart.Left + 150, Top:=outputStart.Top, Width:=400, Height:=300)
    chartObj.Name = "ElbowChart"
    chartObj.Chart.ChartType = xlLineMarkers
    chartObj.Chart.SetSourceData Source:=ws.Range(outputStart, outputStart.Offset(maxK, 1))
    chartObj.Chart.HasTitle = True
    chartObj.Chart.ChartTitle.Text = "Técnica del Codo (SSE vs K)"
    
End Sub

Function RunKMeans(ws As Worksheet, dataRange As Range, k As Integer) As Double

    Dim maxIterations As Integer: maxIterations = 100
    Dim tolerance As Double: tolerance = 0.0001
    
    Dim nRows As Integer: nRows = dataRange.Rows.count
    Dim nCols As Integer: nCols = dataRange.Columns.count
    
    Dim data() As Double
    ReDim data(1 To nRows, 1 To nCols)
    
    Dim i As Long, j As Long
    For i = 1 To nRows
        For j = 1 To nCols
            data(i, j) = dataRange.Cells(i, j).Value
        Next j
    Next i
    
    ' Inicializar centroides al azar
    Dim centroids() As Variant
    ReDim centroids(1 To k, 1 To nCols)
    
    Dim randIndex As Long
    For i = 1 To k
        randIndex = Int((nRows - 1 + 1) * Rnd + 1)
        For j = 1 To nCols
            centroids(i, j) = data(randIndex, j)
        Next j
    Next i
    
    Dim assignments() As Integer
    ReDim assignments(1 To nRows)
    
    Dim changed As Boolean
    Dim iterations As Integer
    Dim dist As Double, minDist As Double
    Dim idx As Integer
    Dim sum() As Double, count() As Long
    Dim sse As Double
    
    Do
        changed = False
        iterations = iterations + 1
        
        ' Asignar puntos al centro más cercano
        For i = 1 To nRows
            minDist = 1E+99
            For j = 1 To k
                dist = 0
                For idx = 1 To nCols
                    dist = dist + (data(i, idx) - centroids(j, idx)) ^ 2
                Next idx
                If dist < minDist Then
                    minDist = dist
                    assignments(i) = j
                End If
            Next j
        Next i
        
        ' Calcular nuevos centroides
        ReDim sum(1 To k, 1 To nCols)
        ReDim count(1 To k)
        
        For i = 1 To nRows
            For j = 1 To nCols
                sum(assignments(i), j) = sum(assignments(i), j) + data(i, j)
            Next j
            count(assignments(i)) = count(assignments(i)) + 1
        Next i
        
        changed = False
        For i = 1 To k
            For j = 1 To nCols
                If count(i) > 0 Then
                    Dim newVal As Double
                    newVal = sum(i, j) / count(i)
                    If Abs(centroids(i, j) - newVal) > tolerance Then
                        changed = True
                        centroids(i, j) = newVal
                    End If
                End If
            Next j
        Next i
        
    Loop While changed And iterations < maxIterations
    
    ' Calcular SSE
    sse = 0
    For i = 1 To nRows
        For j = 1 To nCols
            sse = sse + (data(i, j) - centroids(assignments(i), j)) ^ 2
        Next j
    Next i
    
    RunKMeans = sse

End Function
