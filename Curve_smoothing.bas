Option Explicit
'This function smooths the data using LOESS method
'X and Y are the discrete data to be smoothed, error cells are allowed and would be skipped in process
'XDomain is the vector holding the X values after smoothing, i.e., sampling points
'nPts - controls the degree of the smoothness, the larger this number, the more smooth the curve would be
'Dazhong 29/03/2016

Function smooth_data(X As Variant, y As Variant, xDomain As Variant, nPts As Long) As Variant
    
    If TypeName(X) = "Range" Then
        X = X.Value
    End If
    If TypeName(y) = "Range" Then
       y = y.Value
    End If
    If TypeName(xDomain) = "Range" Then
       xDomain = xDomain.Value
    End If
    
    'check if we are dealing with column vectors
    Dim Is_column As Boolean
    Is_column = is_column_vector(X)

    'transpose the vector if it is a row rather than a column
    If Is_column = False Then
        X = Application.WorksheetFunction.Transpose(X)
        y = Application.WorksheetFunction.Transpose(y)
        xDomain = Application.WorksheetFunction.Transpose(xDomain)
    End If
    'Assert the dimensions of X and Y vectors are the same
    If size_of_vector(X) <> size_of_vector(y) Then
        MsgBox "Dimension of two X and Y should be the same"
        Exit Function
    End If
    'size of the vector n
    Dim n As Integer
    n = size_of_vector(X)
    Dim flag() As Boolean
    ReDim flag(1 To n, 1 To 1)
    'size of the array without error
    Dim n_no_error As Integer
    n_no_error = 0
    'count the non-error cell in the input data and assign a flag to it
    Dim i As Integer
    For i = 1 To n
        If IsError(X(i, 1)) = False And IsError(y(i, 1)) = False Then
            n_no_error = n_no_error + 1
            flag(i, 1) = True
        Else
            flag(i, 1) = False
        End If
    Next i
    'X1 and Y1 hold only the non error values
    Dim X1(), Y1(), xDomain1() As Variant
    ReDim X1(1 To n_no_error, 1 To 1)
    ReDim Y1(1 To n_no_error, 1 To 1)
    'add the non-error values to x1 and Y1
    n_no_error = 0
    For i = 1 To n
        If IsError(X(i, 1)) = False And IsError(y(i, 1)) = False Then
            n_no_error = n_no_error + 1
            X1(n_no_error, 1) = X(i, 1)
            Y1(n_no_error, 1) = y(i, 1)
        End If
    Next i
    'count the non-error cell in the xdomain
    n_no_error = 0
    For i = 1 To size_of_vector(xDomain)
        If IsError(xDomain(i, 1)) = False Then
            n_no_error = n_no_error + 1
            flag(i, 1) = True
        Else
            flag(i, 1) = False
        End If
    Next i
    ReDim xDomain1(1 To n_no_error, 1 To 1)
    n_no_error = 0
    For i = 1 To size_of_vector(xDomain)
        If IsError(xDomain(i, 1)) = False Then
            n_no_error = n_no_error + 1
            xDomain1(n_no_error, 1) = xDomain(i, 1)
        End If
    Next i
    Dim ydomain1 As Variant
    Dim yDomain() As Variant
    ReDim yDomain(1 To size_of_vector(xDomain), 1 To 1)
    
    n_no_error = 0

    ydomain1 = LOESS(X1, Y1, xDomain1, nPts)
    
    For i = 1 To size_of_vector(xDomain)
        If IsError(xDomain(i, 1)) = False Then
            n_no_error = n_no_error + 1
            yDomain(i, 1) = ydomain1(n_no_error, 1)
        Else
            yDomain(i, 1) = CVErr(xlErrNA)
        End If
    Next i
    smooth_data = yDomain
End Function
'This returns the a vector
Function size_of_vector(X As Variant) As Integer
    If is_column_vector(X) Then
        size_of_vector = UBound(X, 1) - LBound(X, 1) + 1
    Else
        size_of_vector = UBound(X, 2) - LBound(X, 2) + 1
    End If
End Function
'This function checks if a vector is a column or a row
Function is_column_vector(X As Variant) As Boolean
    Dim n_col As Integer
    n_col = UBound(X, 2) - LBound(X, 2) + 1
    If n_col = 1 Then
        is_column_vector = True
    Else
        is_column_vector = False
    End If
End Function
'LOESS Function -http://peltiertech.com/loess-smoothing-in-excel/
Public Function LOESS(X As Variant, y As Variant, xDomain As Variant, nPts As Long) As Double()
  Dim i As Long
  Dim iMin As Long
  Dim iMax As Long
  Dim iPoint As Long
  Dim iMx As Long
  Dim mx As Variant
  Dim maxDist As Double
  Dim SumWts As Double, SumWtX As Double, SumWtX2 As Double, SumWtY As Double, SumWtXY As Double
  Dim Denom As Double, WLRSlope As Double, WLRIntercept As Double
  Dim xNow As Double
  Dim distance() As Double
  Dim weight() As Double
  Dim yLoess() As Double
  If TypeName(X) = "Range" Then
    X = X.Value
  End If

  If TypeName(y) = "Range" Then
    y = y.Value
  End If

  If TypeName(xDomain) = "Range" Then
    xDomain = xDomain.Value
  End If

  ReDim yLoess(LBound(xDomain, 1) To UBound(xDomain, 1), 1 To 1)

  For iPoint = LBound(xDomain, 1) To UBound(xDomain, 1)

    iMin = LBound(X, 1)
    iMax = UBound(X, 1)

    xNow = xDomain(iPoint, 1)

    ReDim distance(iMin To iMax)
    ReDim weight(iMin To iMax)

    For i = iMin To iMax
      ' populate x, y, distance
      distance(i) = Abs(X(i, 1) - xNow)
    Next

    Do
      ' find the nPts points closest to xNow
      If iMax + 1 - iMin <= nPts Then Exit Do
      If distance(iMin) > distance(iMax) Then
        ' remove first point
        iMin = iMin + 1
      ElseIf distance(iMin) < distance(iMax) Then
        ' remove last point
        iMax = iMax - 1
      Else
        ' remove both points?
        iMin = iMin + 1
        iMax = iMax - 1
      End If
    Loop

    ' Find max distance
    maxDist = -1
    For i = iMin To iMax
      If distance(i) > maxDist Then maxDist = distance(i)
    Next

    ' calculate weights using scaled distances
    For i = iMin To iMax
      weight(i) = (1 - (distance(i) / maxDist) ^ 3) ^ 3
    Next

    ' do the sums of squares
    SumWts = 0
    SumWtX = 0
    SumWtX2 = 0
    SumWtY = 0
    SumWtXY = 0
    For i = iMin To iMax
      SumWts = SumWts + weight(i)
      SumWtX = SumWtX + X(i, 1) * weight(i)
      SumWtX2 = SumWtX2 + (X(i, 1) ^ 2) * weight(i)
      SumWtY = SumWtY + y(i, 1) * weight(i)
      SumWtXY = SumWtXY + X(i, 1) * y(i, 1) * weight(i)
    Next
    Denom = SumWts * SumWtX2 - SumWtX ^ 2

    ' calculate the regression coefficients, and finally the loess value
    WLRSlope = (SumWts * SumWtXY - SumWtX * SumWtY) / Denom
    WLRIntercept = (SumWtX2 * SumWtY - SumWtX * SumWtXY) / Denom
    yLoess(iPoint, 1) = WLRSlope * xNow + WLRIntercept

  Next

  LOESS = yLoess

End Function



