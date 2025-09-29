Attribute VB_Name = "matrix_Functions"
Option Explicit

'https://stackoverflow.com/questions/69697140/derivatives-of-equations-required-for-the-jacobian-vba

'https://riptutorial.com/vba/example/17455/multidimensional-arrays

'Matrix functions in VB: http://www.java2s.com/Tutorial/VB/0040__Data-Type/Calculatethedeterminantofanarray.htm

'Matrix math: https://www.eng-tips.com/viewthread.cfm?qid=272153&amp%3Bpage=1
'Use Add-Ins: https://www.exceltip.com/custom-functions/how-to-use-your-excel-add-in-functions-in-vba.html
'https://newtonexcelbach.com/2010/05/22/linking-alglib-c-to-excel-vba/
'https://newtonexcelbach.com/downloads/
'https://www.alglib.net/download.php

' USAGE EXAMPLES:
' Matrix(Rows)
'A=[1;4;7] => 3 Rows, 1 Column; ReDim A(3,1); A(1,1) = 1; A(2,1) = 4; A(3,1) = 7; A = MatrixCreate(1, 1, 4, 7)
'A=[1 2 3] => 1 Row, 3 Columns; ReDim A(1,3); A(1,1) = 1; A(1,2) = 2; A(1,3) = 3; A = MatrixCreate(3, 1, 2, 3)
'
' Matrix(Rows, Columns)
'A=[1 2 3;4 5 6] => 2 Rows, 3 Columns; ReDim A(2,3); A(1,1) = 1; A(1,2) = 2; A(1,3) = 3; A(2,1) = 4; A(2,2) = 5; A(2,3) = 6; A = MatrixCreate(3, 1, 2, 3, 4, 5, 6)
'
'A = MatrixCreate(3, 2, 2, 3):  [2       2       3] (1 Row, 3 Columns)
'B = MatrixCreate(3, 1, 4, 7):  [1       4       7] (1 Row, 3 Columns)
'C = MatrixSum(A, B):           [3       6      10] (1 Row, 3 Columns)
'C = MatrixSubstract(A, B):     [1      -2      -4] (1 Row, 3 Columns)
'C = MatrixTimes(A, B):         [2       8      21] (1 Row, 3 Columns)
'C = MatrixProductNumber(2, A): [4       4       6] (1 Row, 3 Columns)
'C = MatrixCreate(3, 5, 5, 6): D = MatrixCreate3x3(A, B, C): (3 Rows, 3 Columns)
'                                2       2       3
'                                1       4       7
'                                5       5       6
'MatrixInverse(D):               1.(2)  -0.(3)  -0.(2)     (3 Rows, 3 Columns)
'                               -3.(2)   0.(3)   1.(2)
'                                1.(6)   0      -0.(6)
'MatrixDeterminant(D):          -9
'E = MatrixProduct(A, D):       [21  27  38] (1 Row, 3 Columns)
'F = MatrixTranspose(E):         21          (3 Rows, 1 Column)
'                                27
'                                38

Private Sub TestResults()
  Dim a, b, c, d
  'A = Rot_z(1.04719755)
  a = MatrixCreate(3, 1, 2, 3, 4, 5, 6, 7, 8, 9)
  Debug.Print MatrixDisplay(a)
  b = MatrixCreate(3, 10, 20, 30, 40, 50, 60, 70, 80, 90)
  Debug.Print MatrixDisplay(b)
'  C = MatrixCreate(3, 5, 5, 6)
'  Debug.Print MatrixDisplay(C)
'  D = MatrixCreate3x3(A, B, C)
'  Debug.Print MatrixDisplay(D)
  d = MatrixProduct(a, b)
  Debug.Print MatrixDisplay(d)
  'Debug.Print MatrixDisplay(MatrixInverse(D))
End Sub


Private Sub MatrixCreate_TestAll()
  Dim a
  'good: 3 rows and 3 columns
  a = MatrixCreate(3, 1, 2, 3, 4, 5, 6, 7, 8, 9)
  Debug.Print MatrixDisplay(a)
  'good: 3 Rows, 1 Column
  a = MatrixCreate(1, 1, 4, 7)
  Debug.Print MatrixDisplay(a)
  'good: 1 Row, 3 Columns
  a = MatrixCreate(3, 1, 2, 3)
  Debug.Print MatrixDisplay(a)
  'good with too many arguments: 3 rows and 3 columns
  a = MatrixCreate(3, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
  Debug.Print MatrixDisplay(a)
  'error for too few arguments: 3 rows and 3 columns
  a = MatrixCreate(3, 1, 2, 3, 4, 5, 6, 7, 8)
  Debug.Print MatrixDisplay(a)
  'good with too many arguments: 1 Row, 3 Columns
  a = MatrixCreate(3, 1, 2, 3, 10)
  Debug.Print MatrixDisplay(a)
  'good with too many arguments: 1 Row, 3 Columns
  a = MatrixCreate(3, 1, 2, 3, 10, 11)
  Debug.Print MatrixDisplay(a)
  'good for too few arguments: 1 Row, 3 Columns
  a = MatrixCreate(3, 1, 2)
  Debug.Print MatrixDisplay(a)
  'error invalid column count:
  a = MatrixCreate(0, 1, 2, 3)
  Debug.Print MatrixDisplay(a)
  'error invalid column count:
  a = MatrixCreate(-1, 1, 2, 3)
  Debug.Print MatrixDisplay(a)
  'good with various data type: 3 rows and 3 columns
  a = MatrixCreate(3, 1, "2", 3.3, True, "Text", 6, 7, 8, 9)
  Debug.Print MatrixDisplay(a)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Assign values to a 2D Array in one command                          '
' Returns:      Returns a 2D Array with first dimension given in the first argument '
'               and all array values afterwards.                                    '
' Remarks:      http://www.cpearson.com/excel/OptionalArgumentsToProcedures.aspx    '
'               Use Array() built-in function to create a 1D Array                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixCreate(ColumnCount As Long, ParamArray Values() As Variant) As Variant
  If ColumnCount > 0 Then
    Dim ValuesCount As Long: ValuesCount = UBound(Values) - LBound(Values) + 1
    Dim RowCount As Long: RowCount = RoundUp(ValuesCount / ColumnCount)
    If RowCount > 0 Then
      Dim MP: ReDim MP(1 To RowCount, 1 To ColumnCount)
      Dim i: i = 0
      Dim Row As Long: For Row = LBound(MP) To UBound(MP)
        Dim Col As Long: For Col = LBound(MP, 2) To UBound(MP, 2)
          If i >= ValuesCount Then GoTo ExitLoop
          MP(Row, Col) = Values(i)
          i = i + 1
        Next
      Next
ExitLoop:
      MatrixCreate = MP
      Exit Function
    End If
  End If
  MatrixCreate = Empty 'Error
End Function

Private Sub MatrixDeterminant_TestAll()
  Dim a
  'error: empty
  Debug.Print MatrixDeterminant(a)
  'good: 3 rows and 3 columns
  a = MatrixCreate(3, 1, 2, 3, 4, 5, 6, 7, 8, 9)
  Debug.Print MatrixDeterminant(a)
  'error: 3 Rows, 1 Column
  a = MatrixCreate(1, 1, 4, 7)
  Debug.Print MatrixDeterminant(a)
  'error: 1 Row, 3 Columns
  a = MatrixCreate(3, 1, 2, 3)
  Debug.Print MatrixDeterminant(a)
  'error with various data type: 3 rows and 3 columns
  a = MatrixCreate(3, 1, "2", 3.3, True, "Text", 6, 7, 8, 9)
  Debug.Print MatrixDeterminant(a)
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Get determinant of a 2D matrix                 '
' Returns:      Returns a number                               '
' Remarks:      Possible only for square matrices              '
'               Application.WorksheetFunction.MDeterm(MyArray) '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixDeterminant(Matrix As Variant) As Double
  On Error Resume Next
  MatrixDeterminant = Application.WorksheetFunction.MDeterm(Matrix)
  If Err.Number <> 0 Then MatrixDeterminant = -99999999# 'Error
End Function


Private Sub MatrixTranspose_TestAll()
  Dim a
  'error: empty
  a = MatrixTranspose(a)
  Debug.Print MatrixDisplay(a)
  'error: Number
  a = MatrixTranspose(1)
  Debug.Print MatrixDisplay(a)
  'error: Text
  a = MatrixTranspose("Text")
  Debug.Print MatrixDisplay(a)
  'error: Array with 1 dimension
  a = MatrixTranspose(Array(1, 2, 3))
  Debug.Print MatrixDisplay(a)
  'good: 3 rows and 3 columns
  a = MatrixCreate(3, 1, 2, 3, 4, 5, 6, 7, 8, 9)
  a = MatrixTranspose(a)
  Debug.Print MatrixDisplay(a)
  'good: 3 Rows, 1 Column
  a = MatrixCreate(1, 1, 4, 7)
  a = MatrixTranspose(a)
  Debug.Print MatrixDisplay(a)
  'good: 1 Row, 3 Columns
  a = MatrixCreate(3, 1, 2, 3)
  a = MatrixTranspose(a)
  Debug.Print MatrixDisplay(a)
  'good with various data type: 3 rows and 3 columns
  a = MatrixCreate(3, 1, "2", 3.3, True, "Text", "I am a long text.", 7, 8, "I am a longer text.")
  a = MatrixTranspose(a)
  Debug.Print MatrixDisplay(a)
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Get transposed 2D matrix                           '
' Returns:      Returns a 2D Array                                 '
' Remarks:      https://www.automateexcel.com/vba/transpose-array/ '
'               Application.WorksheetFunction.Transpose(MyArray)   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixTranspose(Matrix As Variant) As Variant
  If IsArray(Matrix) Then
    Dim maxX As Long: maxX = UBound(Matrix, 1)
    Dim minX As Long: minX = LBound(Matrix, 1)
    On Error Resume Next 'if Matrix has only 1 dimension
    Dim MP As Variant, x As Long
    Dim maxY As Long: maxY = UBound(Matrix, 2)
    If Err.Number <> 0 Then 'Matrix has only 1 dimension
      On Error GoTo 0
      ReDim MP(1 To 1, minX To maxX)
      For x = minX To maxX
        MP(1, x) = Matrix(x)
      Next x
    Else 'Matrix has at least 2 dimensions
      On Error GoTo 0
      Dim minY As Long: minY = LBound(Matrix, 2)
      ReDim MP(minY To maxY, minX To maxX)
      For x = minX To maxX
        Dim y As Long: For y = minY To maxY
          MP(y, x) = Matrix(x, y)
        Next y
      Next x
    End If
    MatrixTranspose = MP
    Exit Function
  End If
  MatrixTranspose = Empty 'Error
End Function

Private Sub MatrixInverse_TestAll()
  Dim a
  'error: empty
  a = MatrixInverse(a)
  Debug.Print MatrixDisplay(a)
  'good: 3 rows and 3 columns
  a = MatrixCreate(3, 1, 2, 3, 4, 5, 6, 7, 8, 9)
  a = MatrixInverse(a)
  Debug.Print MatrixDisplay(a)
  'error: 3 Rows, 1 Column
  a = MatrixCreate(1, 1, 4, 7)
  a = MatrixInverse(a)
  Debug.Print MatrixDisplay(a)
  'error: 1 Row, 3 Columns
  a = MatrixCreate(3, 1, 2, 3)
  a = MatrixInverse(a)
  Debug.Print MatrixDisplay(a)
  'error with various data type: 3 rows and 3 columns
  a = MatrixCreate(3, 1, "2", 3.3, True, "Text", 6, 7, 8, 9)
  a = MatrixInverse(a)
  Debug.Print MatrixDisplay(a)
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Get the Inverse Matrix of another Matrix given as argument.                                         '
' Returns:      The Inverse of the Matrix given in the argument.                                                    '
'               Nothing if the Inverse is not possible (Matrix is singular, has no Inverse).                        '
' Remarks:      Possible only for square matrices                                                                   '
'               https://www.mrexcel.com/board/threads/vba-matrix-inverse.937092/                                    '
'               It seams that VBA Gauss-Jordan inversion code is slower than Excel built-in function MInverse       '
'               A = Range("A2:B4") 'collecting matrix A                                                             '
'               B = Application.WorksheetFunction.MInverse(A)                                                       '
'               https://www.mrexcel.com/board/threads/matrix-worksheet-function-minverse-in-vba-with-array.1200885/ '
'               Operator in Matlab: .'   Function in Scilab: inv(A)                                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixInverse(Matrix As Variant) As Variant
  On Error Resume Next
  MatrixInverse = Application.WorksheetFunction.MInverse(Matrix)
  If Err.Number <> 0 Then MatrixInverse = Empty 'Error or Matrix is singular, has no inverse
End Function

Private Sub Rot_z_TestAll()
  Dim a
  a = Rot_z(1.0471975511966) '60 * Deg2Rad
  Debug.Print MatrixDisplay(a)
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Get Rotation matrix for rotations around Z-axis    '
' Returns:      Returns a 2D Array                                 '
' Remarks:      https://de.mathworks.com/help/phased/ref/rotz.html '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Rot_z(AngleInRadians As Double) As Variant
  Dim t(1 To 3, 1 To 3) As Variant
  t(1, 1) = Cos(AngleInRadians):  t(1, 2) = Sin(AngleInRadians): t(1, 3) = 0
  t(2, 1) = -Sin(AngleInRadians): t(2, 2) = Cos(AngleInRadians): t(2, 3) = 0
  t(3, 1) = 0:                    t(3, 2) = 0:                   t(3, 3) = 1
  Rot_z = t
End Function

Private Sub MatrixCreate3x3_TestAll()
  Dim a, b, c, d
  a = MatrixCreate(3, 1, 2, 3)
  Debug.Print MatrixDisplay(a)
  b = MatrixCreate(3, 4, 5, 6)
  c = MatrixCreate(3, 7, 8, 9)
  d = MatrixCreate3x3(a, b, c)
  Debug.Print MatrixDisplay(d)
  'first array has more elements
  a = MatrixCreate(3, 1, 2, 3, 4, 5)
  Debug.Print MatrixDisplay(a)
  d = MatrixCreate3x3(a, b, c)
  Debug.Print MatrixDisplay(d)
  'A is a simple 1D array, the rest normal 2D arrays
  a = Array(1, 2, 3)
  d = MatrixCreate3x3(a, b, c)
  Debug.Print MatrixDisplay(d)
  'all arrays are simple 1D arrays
  b = Array(4, 5, 6)
  c = Array(7, 8, 9)
  d = MatrixCreate3x3(a, b, c)
  Debug.Print MatrixDisplay(d)
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Assign values to a 2D Array in one command                          '
' Returns:      Returns a 2D Array with first dimension given in the first argument '
'               and all array values afterwards.                                    '
' Remarks:      http://www.cpearson.com/excel/OptionalArgumentsToProcedures.aspx    '
'               Use Array() built-in function to create a 1D Array                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixCreate3x3(ParamArray Matrices1x3() As Variant) As Variant
  If ArraySize(CVar(Matrices1x3)) >= 3 Then
    Dim MP: ReDim MP(1 To 3, 1 To 3)
    Dim Row As Long, Col As Long
    For Row = LBound(MP) To UBound(MP)
      Dim t As Variant: t = Matrices1x3(LBound(Matrices1x3) + Row - 1)
      Select Case ArraySize(t)
      Case Is < 3 'Matrix has at least 1 element on first dimension
        If ArraySize(t, 2) < 3 Then GoTo error 'Matrix has less than 3 elements on second dimension
        For Col = LBound(MP, 2) To UBound(MP, 2)
          MP(Row, Col) = t(LBound(t), LBound(t, 2) + Col - 1)
        Next
      Case Is >= 3 'Matrix has at least 3 elements on first dimension
        For Col = LBound(MP, 2) To UBound(MP, 2)
          MP(Row, Col) = Matrices1x3(LBound(Matrices1x3) + Row - 1)(LBound(t) + Col - 1)
        Next
      Case Else
        GoTo error
      End Select
    Next
    MatrixCreate3x3 = MP
    Exit Function
  End If
error:
  MatrixCreate3x3 = Empty
End Function

Private Sub MatrixTimes_TestAll()

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Multiplies coresponding elements in two 2D matrix   '
' Returns:      Returns a 2D array.                                 '
' Remarks:      https://de.mathworks.com/help/matlab/ref/times.html '
'               operator in Matlab: .*                              '
'               If an element is not a number, it will be skipped   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixTimes(Matrix1 As Variant, Matrix2 As Variant) As Variant
  If ArraySize(Matrix1) <> ArraySize(Matrix2) Then GoTo error
  If ArraySize(Matrix1, 2) <> ArraySize(Matrix2, 2) Then GoTo error
  Dim Row2Offset As Long: Row2Offset = LBound(Matrix2) - LBound(Matrix1)
  Dim Col2Offset As Long: Col2Offset = LBound(Matrix2, 2) - LBound(Matrix1, 2)
  Dim MP: ReDim MP(1 To ArraySize(Matrix1), 1 To ArraySize(Matrix1, 2))
  Dim Row As Long: For Row = LBound(Matrix1) To UBound(Matrix1)
    Dim Col As Long: For Col = LBound(Matrix1, 2) To UBound(Matrix1, 2)
      If IsNumeric(Matrix1(Row, Col)) And IsNumeric(Matrix2(Row + Row2Offset, Col + Col2Offset)) Then _
        MP(Row, Col) = Matrix1(Row, Col) * Matrix2(Row + Row2Offset, Col + Col2Offset)
    Next
  Next
  MatrixTimes = MP
  Exit Function
error:
  MatrixTimes = Empty
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Add two 2D matrix                                                                     '
' Returns:      Returns a 2D array containing the matrix sum of two matrices contained in 2D arrays.  '
' Remarks:      https://stackoverflow.com/questions/53508080/adding-arrays-together-in-vba-for-output '
'               8 times quicker as WorksheetFunction.MMult(MultArr, Array(Matrix1, Matrix2))          '
'               If an element is not a number, it will be skipped                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixSum(Matrix1 As Variant, Matrix2 As Variant) As Variant
  If ArraySize(Matrix1) <> ArraySize(Matrix2) Then GoTo error
  If ArraySize(Matrix1, 2) <> ArraySize(Matrix2, 2) Then GoTo error
  Dim Row2Offset As Long: Row2Offset = LBound(Matrix2) - LBound(Matrix1)
  Dim Col2Offset As Long: Col2Offset = LBound(Matrix2, 2) - LBound(Matrix1, 2)
  Dim Row1 As Long, Col2 As Long
  Dim MP: ReDim MP(1 To ArraySize(Matrix1), 1 To ArraySize(Matrix1, 2))
  Dim Row As Long: For Row = LBound(Matrix1) To UBound(Matrix1)
    Dim Col As Long: For Col = LBound(Matrix1, 2) To UBound(Matrix1, 2)
      If IsNumeric(Matrix1(Row, Col)) And IsNumeric(Matrix2(Row + Row2Offset, Col + Col2Offset)) Then _
        MP(Row, Col) = Matrix1(Row, Col) + Matrix2(Row + Row2Offset, Col + Col2Offset)
    Next
  Next
  MatrixSum = MP
  Exit Function
error:
  MatrixSum = Empty
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Substract second matrix from the first one        '
' Returns:      Returns a 2D array.                               '
' Remarks:      If an element is not a number, it will be skipped '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixSubstract(Matrix1 As Variant, Matrix2 As Variant) As Variant
  If ArraySize(Matrix1) <> ArraySize(Matrix2) Then GoTo error
  If ArraySize(Matrix1, 2) <> ArraySize(Matrix2, 2) Then GoTo error
  Dim Row2Offset As Long: Row2Offset = LBound(Matrix2) - LBound(Matrix1)
  Dim Col2Offset As Long: Col2Offset = LBound(Matrix2, 2) - LBound(Matrix1, 2)
  Dim Row1 As Long, Col2 As Long
  Dim MP: ReDim MP(1 To ArraySize(Matrix1), 1 To ArraySize(Matrix1, 2))
  Dim Row As Long: For Row = LBound(Matrix1) To UBound(Matrix1)
    Dim Col As Long: For Col = LBound(Matrix1, 2) To UBound(Matrix1, 2)
      If IsNumeric(Matrix1(Row, Col)) And IsNumeric(Matrix2(Row + Row2Offset, Col + Col2Offset)) Then _
        MP(Row, Col) = Matrix1(Row, Col) - Matrix2(Row + Row2Offset, Col + Col2Offset)
    Next
  Next
  MatrixSubstract = MP
  Exit Function
error:
  MatrixSubstract = Empty
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Multiply a number to a 0D, 1D, 2D or 3D matrix        '
' Returns:      Returns a matrix of same size, containing the product '
'               between the number and the matrix                     '
' Remarks:      https://stackoverflow.com/a/35671729                  '
'               If an element is not a number, it will be skipped     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixProductNumber(Number As Double, Matrix As Variant) As Variant
  Dim t As Long, max_sizes As Long
  
  On Error GoTo max_sizes_found
  max_sizes = 1
  Do
    t = UBound(Matrix, max_sizes)
    max_sizes = max_sizes + 1
  Loop While max_sizes < 3
max_sizes_found:
  On Error GoTo 0
  Dim MP, Row As Long, Col As Long, z1 As Long
  Select Case max_sizes - 1
    Case 0:
      If IsNumeric(Matrix) Then _
        MP = Matrix * Number
    Case 1:
      ReDim MP(LBound(Matrix) To UBound(Matrix))
      For Col = LBound(Matrix) To UBound(Matrix)
        If IsNumeric(Matrix(Col)) Then _
          MP(Col) = Matrix(Col) * Number
      Next
    Case 2:
      ReDim MP(LBound(Matrix) To UBound(Matrix), LBound(Matrix, 2) To UBound(Matrix, 2))
      For Row = LBound(Matrix) To UBound(Matrix)
        For Col = LBound(Matrix, 2) To UBound(Matrix, 2)
          If IsNumeric(Matrix(Row, Col)) Then _
            MP(Row, Col) = Matrix(Row, Col) * Number
        Next
      Next
    Case 3:
      ReDim MP(LBound(Matrix) To UBound(Matrix), LBound(Matrix, 2) To UBound(Matrix, 2), LBound(Matrix, 3) To UBound(Matrix, 3))
      For z1 = LBound(Matrix) To UBound(Matrix)
        For Row = LBound(Matrix, 2) To UBound(Matrix, 2)
          For Col = LBound(Matrix, 3) To UBound(Matrix, 3)
            If IsNumeric(Matrix(z1, Row, Col)) Then _
              MP(Row, Col, z1) = Matrix(z1, Row, Col) * Number
          Next
        Next
      Next
    Case Else:
      MP = Empty
  End Select
  MatrixProductNumber = MP
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Multiply two 2D matrices                                                     '
' Returns:      Returns a 2D array containing the matrix product                             '
'               of two matrices contained in 2D arrays.                                      '
'               If an element inside any matrix is not a number, Result will be empty.       '
' Remarks:      https://stackoverflow.com/questions/61834457/matrix-multiplication-using-vba '
'               10 times quicker as WorksheetFunction.MMult(Matrix1, Matrix2)                '
'               https://www.pls-fix-thx.com/post/vba-matrix-multiplication                   '
'               A = Range("A2:B4") 'collecting matrix A                                      '
'               B = Range("D2:D3") 'collecting matrix B                                      '
'               'Calculating Matrix C with .MMult syntax.                                    '
'               C = Application.MMult(A, B)                                                  '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixProduct(Matrix1 As Variant, Matrix2 As Variant) As Variant
  If ArraySize(Matrix1, 2) <> ArraySize(Matrix2) Then GoTo error
  Dim ColRowOffset As Long: ColRowOffset = LBound(Matrix2) - LBound(Matrix1, 2)
  Dim MP: ReDim MP(LBound(Matrix1) To UBound(Matrix1), LBound(Matrix2, 2) To UBound(Matrix2, 2))
  Dim Row As Long: For Row = LBound(Matrix1) To UBound(Matrix1)
    Dim Col As Long: For Col = LBound(Matrix2, 2) To UBound(Matrix2, 2)
      Dim ColRow As Long: For ColRow = LBound(Matrix1, 2) To UBound(Matrix1, 2)
        If Not (IsNumeric(Matrix1(Row, ColRow)) And IsNumeric(Matrix2(ColRow + ColRowOffset, Col))) Then GoTo error
        MP(Row, Col) = MP(Row, Col) + Matrix1(Row, ColRow) * Matrix2(ColRow + ColRowOffset, Col)
      Next
    Next
  Next
  If Col = 2 And Row = 2 Then MatrixProduct = MP(1, 1) Else MatrixProduct = MP
  Exit Function
error:
  MatrixProduct = Empty
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose:      Display a 0D, 1D, 2D or 3D matrix        '
' Returns:      Returns a string                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MatrixDisplay(Matrix As Variant) As String
  Dim t As Long, max_sizes As Long
  
  On Error GoTo max_sizes_found
  max_sizes = 1
  Do
    t = UBound(Matrix, max_sizes)
    max_sizes = max_sizes + 1
  Loop While max_sizes < 3
max_sizes_found:
  Dim s As String, Row1 As Long, Col1 As Long, z1 As Long
  Select Case max_sizes - 1
    Case 0:
      s = MyFormat(Matrix)
    Case 1: 'PASS
      s = "[" & LBound(Matrix) & " To " & UBound(Matrix) & "] =" & vbNewLine
      For Col1 = LBound(Matrix) To UBound(Matrix)
        s = s & MyFormat(Matrix(Col1)) & "  "
      Next Col1
    Case 2: 'PASS
      s = "[" & LBound(Matrix) & " To " & UBound(Matrix) & ", " & LBound(Matrix, 2) & " To " & UBound(Matrix, 2) & "] =" & vbNewLine
      For Row1 = LBound(Matrix) To UBound(Matrix)
        For Col1 = LBound(Matrix, 2) To UBound(Matrix, 2)
          s = s & MyFormat(Matrix(Row1, Col1)) & "  "
        Next Col1
        s = s & vbNewLine
      Next Row1
    Case 3:
      s = "[" & LBound(Matrix) & " To " & UBound(Matrix) & ", " & LBound(Matrix, 2) & " To " & UBound(Matrix, 2) & ", " & LBound(Matrix, 3) & " To " & UBound(Matrix, 3) & "] =" & vbNewLine
      For z1 = LBound(Matrix) To UBound(Matrix)
        s = vbNewLine & "Level " & z1 & ":" & vbNewLine
        For Row1 = LBound(Matrix, 2) To UBound(Matrix, 2)
          For Col1 = LBound(Matrix, 3) To UBound(Matrix, 3)
            s = s & MyFormat(Matrix(z1, Row1, Col1)) & "  "
          Next Col1
          s = s & vbNewLine
        Next Row1
      Next z1
    Case Else
      s = "Error: matrices size over 3 are not supported to be displayed."
  End Select
  MatrixDisplay = s
End Function

Public Function MyFormat(ByVal d As Variant, Optional Places As Long = 7, Optional DecimalPlaces As Long = 10) As String
  If IsNumeric(d) Then
    Const myDecSep As String = "."
    Const myThousSep As String = ""
    Dim DecSep As String, ThousSep As String
    
    Dim s1 As String: s1 = format(Fix(d), "#,##0.")
    DecSep = Right(s1, 1): s1 = Left(s1, Len(s1) - 1)
    If Abs(d) >= 1000# Then
      ThousSep = Mid(s1, Len(s1) - 3, 1)
      s1 = Replace(s1, ThousSep, myThousSep)
    ElseIf d < 0# Then
      If Left(s1, 1) <> "-" Then s1 = "-0"
    End If
    If Len(s1) < Places Then s1 = Space(Places - Len(s1)) & s1
    
    Dim s2 As String: s2 = format(d, "." & String(DecimalPlaces, "#"))
    s2 = Mid(s2, InStr(1, s2, DecSep, vbBinaryCompare) + 1)
    If Len(s2) < DecimalPlaces Then s2 = s2 & Space(DecimalPlaces - Len(s2))
  
    MyFormat = s1 & myDecSep & s2
  Else
    MyFormat = CStr(d)
    If Len(MyFormat) > Places + DecimalPlaces + 1 Then
      MyFormat = Left(MyFormat, Places + DecimalPlaces - 2) & "..."
    Else
      MyFormat = MyFormat & String(Places + DecimalPlaces + 1 - Len(MyFormat), " ")
    End If
  End If
End Function
