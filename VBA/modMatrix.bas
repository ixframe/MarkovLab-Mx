Attribute VB_Name = "modMatrix"
Option Explicit

'===============================================================================
' Proyecto : MarkovLab-Mx
' Modulo   : modMatrix
' Version  : 0.1.0
' Objetivo : Operaciones matriciales fundamentales
'===============================================================================

'-------------------------------------------------------------------------------
' Crea una matriz de ceros de dimensiones nRows x nCols.
'
' Entrada:
'   nRows = numero de filas
'   nCols = numero de columnas
'
' Salida:
'   Matriz bidimensional de tipo Double
'-------------------------------------------------------------------------------
Public Function MatrixZero(ByVal nRows As Long, _
                           ByVal nCols As Long) As Variant

    Dim A() As Double

    If nRows < 1 Or nCols < 1 Then
        Err.Raise vbObjectError + 1000, _
                  "MatrixZero", _
                  "Las dimensiones de la matriz deben ser mayores que cero."
    End If

    ReDim A(1 To nRows, 1 To nCols)

    MatrixZero = A

End Function

'-------------------------------------------------------------------------------
' Crea una matriz identidad cuadrada de orden n.
'
' Entrada:
'   n = orden de la matriz
'
' Salida:
'   Matriz identidad n x n de tipo Double
'-------------------------------------------------------------------------------
Public Function MatrixIdentity(ByVal n As Long) As Variant

    Dim A() As Double
    Dim i As Long

    If n < 1 Then
        Err.Raise vbObjectError + 1001, _
                  "MatrixIdentity", _
                  "El orden de la matriz debe ser mayor que cero."
    End If

    ReDim A(1 To n, 1 To n)

    For i = 1 To n
        A(i, i) = 1#
    Next i

    MatrixIdentity = A

End Function


'-------------------------------------------------------------------------------
' Crea una copia independiente de una matriz bidimensional.
'
' Entrada:
'   sourceMatrix = matriz original
'
' Salida:
'   Copia independiente de la matriz original
'-------------------------------------------------------------------------------
Public Function MatrixCopy(ByVal sourceMatrix As Variant) As Variant

    Dim resultMatrix() As Double
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long

    firstRow = LBound(sourceMatrix, 1)
    lastRow = UBound(sourceMatrix, 1)
    firstCol = LBound(sourceMatrix, 2)
    lastCol = UBound(sourceMatrix, 2)

    ReDim resultMatrix(firstRow To lastRow, _
                       firstCol To lastCol)

    For i = firstRow To lastRow
        For j = firstCol To lastCol
            resultMatrix(i, j) = CDbl(sourceMatrix(i, j))
        Next j
    Next i

    MatrixCopy = resultMatrix

End Function


'-------------------------------------------------------------------------------
' Multiplica dos matrices bidimensionales.
'
' Entrada:
'   matrixA = primera matriz
'   matrixB = segunda matriz
'
' Salida:
'   Producto matricial matrixA * matrixB
'
' Condicion:
'   El numero de columnas de matrixA debe ser igual
'   al numero de filas de matrixB.
'-------------------------------------------------------------------------------
Public Function MatrixMultiply(ByVal matrixA As Variant, _
                               ByVal matrixB As Variant) As Variant

    Dim resultMatrix() As Double

    Dim firstRowA As Long
    Dim lastRowA As Long
    Dim firstColA As Long
    Dim lastColA As Long

    Dim firstRowB As Long
    Dim lastRowB As Long
    Dim firstColB As Long
    Dim lastColB As Long

    Dim nRowsA As Long
    Dim nColsA As Long
    Dim nRowsB As Long
    Dim nColsB As Long

    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim sumValue As Double

    firstRowA = LBound(matrixA, 1)
    lastRowA = UBound(matrixA, 1)
    firstColA = LBound(matrixA, 2)
    lastColA = UBound(matrixA, 2)

    firstRowB = LBound(matrixB, 1)
    lastRowB = UBound(matrixB, 1)
    firstColB = LBound(matrixB, 2)
    lastColB = UBound(matrixB, 2)

    nRowsA = lastRowA - firstRowA + 1
    nColsA = lastColA - firstColA + 1

    nRowsB = lastRowB - firstRowB + 1
    nColsB = lastColB - firstColB + 1

    If nColsA <> nRowsB Then
        Err.Raise vbObjectError + 1002, _
                  "MatrixMultiply", _
                  "Las dimensiones de las matrices no son compatibles."
    End If

    ReDim resultMatrix(1 To nRowsA, 1 To nColsB)

    For i = 1 To nRowsA
        For j = 1 To nColsB

            sumValue = 0#

            For k = 1 To nColsA
                sumValue = sumValue + _
                    CDbl(matrixA(firstRowA + i - 1, _
                                 firstColA + k - 1)) * _
                    CDbl(matrixB(firstRowB + k - 1, _
                                 firstColB + j - 1))
            Next k

            resultMatrix(i, j) = sumValue

        Next j
    Next i

    MatrixMultiply = resultMatrix

End Function