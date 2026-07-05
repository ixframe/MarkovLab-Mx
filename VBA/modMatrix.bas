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