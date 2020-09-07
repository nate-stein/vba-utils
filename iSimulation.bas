Attribute VB_Name = "iSimulation"
Public Sub MCSim()

' MCSim runs a simple simulation in a spreadsheet
'
' To run a simulation
' 1. Put any number of formulas (Nformula) to simulate in a row
' 2. Put the number of simulation trials in a column underneath the first formula
' 3. If you wish to compute percentiles, enter the percentiles in the first column starting in the 9th row
' 3. Select the region with Nformula+1 columns and 8 or more rows.
'    The region should start one cell to the left of the first formula
' 4. Run the simulation by pressing Ctrl-Shift-M or use Tools | Macro | Run MCSim
'
' Programmer: Mark Broadie
'
' This VBA code is adapted from MonteCarlito created by Martin Auer.
' See www.montecarlito.com

Dim rngMCOut As Range
Set rngMCOut = Application.Selection
     '   user selects a range with 8 rows and N+1 columns, where
     '   1st row:  Blank     formula1   formula2  ...  formulaN
     '   2nd row:  "Trials"  Ntrial
     '   Rows 3 through 8 will contain the CPU seconds and the avg, std dev, std err,
     '         min and max for each formula
     '
Dim Ntrial As Long                ' number of simulation trials
Dim blnHideApplication As Boolean ' true means hide Excel for faster performance
Dim Nformula As Integer           ' number of formulas to simulate
Dim intBatchSize As Integer       ' number of trials before updating display
Dim dblStartTime As Double        ' timing variables
Dim dblFinishTime As Double
Dim dblElapseTime As Double
Dim results() As Variant          ' variables to store results
Dim AvgVec() As Variant
Dim StdDevVec() As Variant
Dim StdErrVec() As Variant
Dim MinVec() As Variant
Dim MaxVec() As Variant
Dim i As Long, j As Integer       ' looping / temporary variables
Dim TempVec() As Variant
Dim Percentiles() As Variant      'for computing percentiles
Dim PercentileValues() As Variant
Dim Npercentile As Integer

dblStartTime = Timer

Nformula = rngMCOut.cells.Columns.count - 1
Ntrial = rngMCOut.cells(2, 2)
If (Ntrial <= 2) Then
   Ntrial = 2
End If
    
ReDim AvgVec(1 To Nformula)
ReDim StdDevVec(1 To Nformula)
ReDim StdErrVec(1 To Nformula)
ReDim MinVec(1 To Nformula)
ReDim MaxVec(1 To Nformula)
ReDim results(1 To Ntrial, 1 To Nformula)
ReDim TempVec(1 To Ntrial)
    
' set up for computing percentiles
If rngMCOut.cells.Rows.count > 8 Then
    Npercentile = rngMCOut.cells.Rows.count - 9
Else
    Npercentile = 0
End If
If Npercentile > 0 Then
    ReDim Percentiles(1 To Npercentile)
    ReDim PercentileValues(1 To Npercentile, 1 To Nformula)
    For i = 1 To Npercentile
        Percentiles(i) = rngMCOut.cells(9 + i, 1)
        If Percentiles(i) < 0 Or Percentiles(i) > 1 Then
            MsgBox "Error: percentiles must be in the range [0,1]. Aborting.", vbCritical, "Error"
            Exit Sub
        End If
    Next i
End If
    
' set intBatchSize
intBatchSize = 100
If Ntrial <= 100 Then
   intBatchSize = 20
End If
If Ntrial > 10000 Then
   intBatchSize = Int(Ntrial / 100)
End If

' Main simulation loop
For i = 1 To Ntrial
    
    ' don't display current trial number unless it is a
    ' multiple of intBatchSize or Ntrial trials is reached
    If (i Mod intBatchSize = 0 Or i = Ntrial) Then
       rngMCOut.cells(2, 2) = i
    End If
   
    ' recalculate the spreadsheet and record the results
    Application.Calculate
    For j = 1 To Nformula
        results(i, j) = rngMCOut.cells(1, 1 + j)
    Next j
Next i
    
' Calculate statistics
For j = 1 To Nformula
    ' temporary variables
    Dim sum As Double
    Dim min As Double
    Dim max As Double
    Dim avg As Double
    Dim var As Double
    Dim eps As Double
    Dim x As Double
    Dim dx As Double
    
    ' first pass over the data to compute the mean, min, and max
    sum = results(1, j)
    min = results(1, j)
    max = results(1, j)
    For i = 2 To Ntrial
        x = results(i, j)
        sum = sum + x
        If x > max Then
            max = x
        ElseIf results(i, j) < min Then
            min = x
        End If
    Next i
    avg = sum / Ntrial
    
    ' second pass over the data to compute the standard deviation
    ' use numerical recipes trick to reduce roundoff error
    eps = 0
    var = 0
    For i = 1 To Ntrial
        x = results(i, j)
        dx = x - avg
        eps = eps + dx
        var = var + dx * dx
    Next i
    var = (var + eps * eps / Ntrial) / (Ntrial - 1)
    
    AvgVec(j) = avg
    StdDevVec(j) = Math.Sqr(var)
    StdErrVec(j) = StdDevVec(j) / Math.Sqr(Ntrial)
    MinVec(j) = min
    MaxVec(j) = max
    
    'compute percentiles, if necessary
    If Npercentile > 0 Then
        'copy the data to a vector
        Matrix2Vec TempVec, results, j, Ntrial
        For i = 1 To Npercentile
            ' compute each percentile
            PercentileValues(i, j) = ComputePercentile(TempVec, Percentiles(i), Ntrial)
        Next i
    End If
Next j
    
dblFinishTime = Timer
dblElapseTime = dblFinishTime - dblStartTime
    
' Print the result to the spreadsheet
rngMCOut.cells(2, 1) = "Number of Trials"
rngMCOut.cells(2, 2) = Ntrial

rngMCOut.cells(3, 1) = "CPU seconds"
rngMCOut.cells(3, 2) = dblElapseTime

rngMCOut.cells(4, 1) = "Average"
Call FillOutput(4, 2, AvgVec(), rngMCOut)

rngMCOut.cells(5, 1) = "Standard deviation"
Call FillOutput(5, 2, StdDevVec(), rngMCOut)

rngMCOut.cells(6, 1) = "Standard error"
Call FillOutput(6, 2, StdErrVec(), rngMCOut)
    
rngMCOut.cells(7, 1) = "Minimum"
Call FillOutput(7, 2, MinVec(), rngMCOut)
    
rngMCOut.cells(8, 1) = "Maximum"
Call FillOutput(8, 2, MaxVec(), rngMCOut)
              
If Npercentile > 0 Then
    rngMCOut.cells(9, 1) = "Percentiles"
    For i = 1 To Npercentile
        rngMCOut.cells(9 + i, 1) = Percentiles(i)
        For j = 1 To Nformula
            rngMCOut.cells(9 + i, 1 + j) = PercentileValues(i, j)
        Next j
    Next i
End If
    
End Sub

' compute a percentile of a data array
' note that the order of elements in the array will change
Function ComputePercentile(data() As Variant, percentile As Variant, n As Long)
    Dim position As Long
    position = Int(percentile * (n - 1)) + 1
    ComputePercentile = PartitionSelect(data, 1, n, position)
End Function

' partition-based order statistic calculation
Private Function PartitionSelect(x() As Variant, offset As Long, length As Long, index As Long) As Variant
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim v As Variant
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim l As Long
    Dim n As Long
    Dim s As Long
    
    'for small arrays, just use an insertion sort
    If length < 7 Then
        For i = offset To length + offset - 1
            For j = i To offset + 1 Step -1
                If x(j - 1) > x(j) Then
                    swap x, j, j - 1
                End If
            Next j
        Next i
        PartitionSelect = x(index)
        Exit Function
    End If
        
    ' choose a partition element, v
    m = offset + length / 2     ' Small arrays, middle element
    If length > 7 Then
        l = offset
        n = offset + length - 1
        If length > 40 Then ' Big arrays, pseudomedian of 9
            s = length / 8
            l = med3(x, l, l + s, l + 2 * s)
            m = med3(x, m - s, m, m + s)
            n = med3(x, n - 2 * s, n - s, n)
        End If
        m = med3(x, l, m, n) ' Mid-size, med of 3
    End If
    v = x(m)
    
    ' Establish Invariant: v* (<v)* (>v)* v*
    a = offset
    b = a
    c = offset + length - 1
    d = c
    Do
        Do
            If b > c Then
                Exit Do
            End If
            If x(b) > v Then
                Exit Do
            End If
            If x(b) = v Then
                swap x, a, b
                a = a + 1
            End If
            b = b + 1
        Loop
        
        Do
            If c < b Then
                Exit Do
            End If
            If x(c) < v Then
                Exit Do
            End If
            If x(c) = v Then
                swap x, c, d
                d = d - 1
            End If
            c = c - 1
        Loop
        
        If b > c Then
            Exit Do
        End If
            
        swap x, b, c
        b = b + 1
        c = c - 1
    Loop
    
    ' Swap partition elements back to middle
    n = offset + length
    s = a - offset
    If s > b - a Then
        s = b - a
    End If
    vecswap x, offset, b - s, s
    s = d - c
    If s > n - d - 1 Then
        s = n - d - 1
    End If
    vecswap x, b, n - s, s

    ' recursively select from proper partition
    ' first partition
    s = b - a
    If index < offset + s Then
        PartitionSelect = PartitionSelect(x, offset, s, index)
        Exit Function
    End If
    ' last partition
    s = d - c
    If index >= n - s Then
        PartitionSelect = PartitionSelect(x, n - s, s, index)
        Exit Function
    End If
    ' it must be the middle partition
    PartitionSelect = v
End Function
            
' utiltity function, swaps x(a) with x(b)
Private Sub swap(x() As Variant, a As Long, b As Long)
    Dim t As Variant
    t = x(a)
    x(a) = x(b)
    x(b) = t
End Sub
    

' utility function, swaps x(a .. (a+n-1)) with x(b .. (b+n-1))
Private Sub vecswap(x() As Variant, a As Long, b As Long, n As Long)
    Dim t As Variant
    Dim i, ap, bp As Long
    ap = a
    bp = b
    For i = 1 To n
        t = x(ap)
        x(ap) = x(bp)
        x(bp) = t
        ap = ap + 1
        bp = bp + 1
    Next i
End Sub

' utility function, returns the index of the median of the three indexed variables
Private Function med3(x() As Variant, a As Long, b As Long, c As Long) As Long
    If x(a) < x(b) Then
        If x(b) < x(c) Then
            med3 = b
        ElseIf x(a) < x(c) Then
            med3 = c
        Else
            med3 = a
        End If
    Else
        If x(b) > x(c) Then
            med3 = b
        ElseIf x(a) > x(c) Then
            med3 = c
        Else
            med3 = a
        End If
    End If
End Function


Sub Matrix2Vec(Bvec() As Variant, Amatrix() As Variant, j As Integer, Nrow As Long)

' Matrix2Vec copies column J of the matrix Amatrix to the vector Bvec

Dim i As Long
For i = 1 To Nrow
    Bvec(i) = Amatrix(i, j)
Next i
    
End Sub

Sub FillOutput(Irow As Integer, JColStart As Integer, Avec() As Variant, rng As Range)
    
' FillOutput copies the vector Avec into row Irow of Rng starting at column JColStart
    
Dim i As Integer
For i = 1 To UBound(Avec)
    rng.cells(Irow, JColStart + i - 1) = Avec(i)
Next i

End Sub





