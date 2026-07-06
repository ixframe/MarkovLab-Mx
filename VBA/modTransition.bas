Attribute VB_Name = "modTransition"
Option Explicit

'===============================================================================
' Proyecto : MarkovLab-Mx
' Modulo   : modTransition
' Version  : 0.1.0
' Objetivo : Construccion de matrices de transicion para cadenas de Markov
'===============================================================================

'-------------------------------------------------------------------------------
' Construye una matriz de conteos de transicion a partir de una serie.
'
' Entrada:
'   stateSeries = matriz bidimensional de una sola columna
'   minState    = valor minimo permitido
'   maxState    = valor maximo permitido
'
' Salida:
'   Matriz cuadrada de conteos.
'
' Interpretacion:
'   counts(i, j) = numero de veces que el estado i
'                  fue seguido inmediatamente por el estado j.
'-------------------------------------------------------------------------------
Public Function BuildTransitionCounts(ByVal stateSeries As Variant, _
                                      ByVal minState As Long, _
                                      ByVal maxState As Long) As Variant

    Dim counts() As Double
    Dim firstRow As Long
    Dim lastRow As Long
    Dim nStates As Long

    Dim currentState As Long
    Dim nextState As Long

    Dim currentIndex As Long
    Dim nextIndex As Long

    Dim i As Long

    If maxState < minState Then
        Err.Raise vbObjectError + 2000, _
                  "BuildTransitionCounts", _
                  "maxState debe ser mayor o igual que minState."
    End If

    firstRow = LBound(stateSeries, 1)
    lastRow = UBound(stateSeries, 1)

    If lastRow <= firstRow Then
        Err.Raise vbObjectError + 2001, _
                  "BuildTransitionCounts", _
                  "La serie debe contener al menos dos observaciones."
    End If

    nStates = maxState - minState + 1
    ReDim counts(1 To nStates, 1 To nStates)

    For i = firstRow To lastRow - 1

        currentState = CLng(stateSeries(i, 1))
        nextState = CLng(stateSeries(i + 1, 1))

        If currentState < minState Or currentState > maxState Then
            Err.Raise vbObjectError + 2002, _
                      "BuildTransitionCounts", _
                      "Se encontro un estado actual fuera del rango permitido."
        End If

        If nextState < minState Or nextState > maxState Then
            Err.Raise vbObjectError + 2003, _
                      "BuildTransitionCounts", _
                      "Se encontro un estado siguiente fuera del rango permitido."
        End If

        currentIndex = currentState - minState + 1
        nextIndex = nextState - minState + 1

        counts(currentIndex, nextIndex) = _
            counts(currentIndex, nextIndex) + 1#

    Next i

    BuildTransitionCounts = counts

End Function