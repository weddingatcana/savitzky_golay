Attribute VB_Name = "modOptimization"
Option Explicit

Public Type tPeaks
    peaks_2D() As Double
    peaks_SG() As Double
End Type

Public Enum ePadding
    padding = 1
    no_padding = 0
End Enum

Public Function optPolyCoeff(ByRef A#(), _
                             ByVal polyOrder&) As Double()

    Dim rawX#(), _
        rawY#(), _
        Vm#(), _
        i_coeff#(), _
        f_coeff#()

        
        rawX = modMatrix.matVec(A, 1)
        rawY = modMatrix.matVec(A, 2)
        Vm = modMath.mathVandermonde(rawX, polyOrder)
        
        i_coeff = modMatrix.matPin(Vm)
        f_coeff = modMatrix.matMul(i_coeff, rawY)
        
        optPolyCoeff = f_coeff

End Function

Public Function optPolyFit(ByRef A#(), _
                           ByVal polyOrder&) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowCoeff&, _
        coeff#(), _
        i&, k&, _
        sum#, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        coeff = optPolyCoeff(A, polyOrder)
        rawMaxRowCoeff = UBound(coeff, 1)
        
        ReDim C(1 To rawMaxRowA, 1 To rawMaxColA)
        
        For i = 1 To rawMaxRowA
        
            sum = 0
            For k = 1 To rawMaxRowCoeff
            
                sum = sum + (coeff(k, 1) * (A(i, 1) ^ (k - 1)))
                    
            Next k
            
            C(i, 1) = A(i, 1)
            C(i, 2) = sum
            
        Next i
        
        optPolyFit = C

End Function

Public Function optPolyFit_seperate_coeff(ByRef A#(), _
                                          ByRef coeff#()) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        rawMaxRowCoeff&, _
        i&, k&, _
        sum#, _
        C#()
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        rawMaxRowCoeff = UBound(coeff, 1)
        
        ReDim C(1 To rawMaxRowA, 1 To (rawMaxColA + 1))
        
        For i = 1 To rawMaxRowA
        
            sum = 0
            For k = 1 To rawMaxRowCoeff
            
                sum = sum + (coeff(k, 1) * (A(i, 1) ^ (k - 1)))
                    
            Next k
            
            C(i, 1) = A(i, 1)
            C(i, 2) = sum
            
        Next i
        
        optPolyFit_seperate_coeff = C

End Function

Public Function optSavGol(ByRef A#(), _
                 Optional ByVal window& = 11, _
                 Optional ByVal polyOrder& = 2, _
                 Optional ByVal interp As ePadding = padding) As Double()

    Dim rawMaxRowA&, _
        rawMaxColA&, _
        moving_mid&, _
        i&, j&, _
        k&, p&, _
        length&, _
        buffer#(), _
        buffer_ends#(), _
        poly#(), _
        poly_ends#(), _
        mid&, _
        C#()

        If window Mod 2 <> 0 And _
           window >= polyOrder + 1 Then
        Else
            Exit Function
        End If

        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
        
        If rawMaxColA <> 2 Or _
           window > rawMaxRowA Then
            Exit Function
        End If
        
        If interp = padding Then
        
            length = rawMaxRowA - (window - 1)
            mid = (window \ 2) + 1
            
            ReDim C(1 To rawMaxRowA, 1 To rawMaxColA)
            ReDim buffer(1 To window, 1 To rawMaxColA)
            ReDim poly(1 To window, 1 To rawMaxColA)
            
            For i = 1 To rawMaxRowA
                
                j = i
                k = 1
                Do
                
                    If k > window Then
                        Exit Do
                    End If
            
                    buffer(k, 1) = A(j, 1)
                    buffer(k, 2) = A(j, 2)
                    
                    j = j + 1
                    k = k + 1
            
                Loop
                
                poly = modOptimization.optPolyFit(buffer, polyOrder)
                
                If i = 1 Then
                
                    For p = i To mid
                        C(p, 1) = A(p, 1)
                        C(p, 2) = poly(p, 2)
                    Next p
                                
                ElseIf i > 1 And i < length Then
                    
                    moving_mid = mid + i - 1
                    C(moving_mid, 1) = A(moving_mid, 1)
                    C(moving_mid, 2) = poly(mid, 2)
                
                ElseIf i = length Then
                
                    For p = (rawMaxRowA - mid + 1) To rawMaxRowA
                        C(p, 1) = A(p, 1)
                        C(p, 2) = poly(mid, 2)
                        mid = mid + 1
                    Next p
                    
                    optSavGol = C
                    Exit Function
                
                Else
                End If
            
            Next i
        
        Else
        
            length = rawMaxRowA - (window - 1)
            mid = (window \ 2) + 1
            moving_mid = mid
            
            ReDim C(1 To length, 1 To rawMaxColA)
            ReDim buffer(1 To window, 1 To rawMaxColA)
            ReDim poly(1 To window, 1 To rawMaxColA)
            
            For i = 1 To length
                
                j = i
                k = 1
                Do
                
                    If k > window Then
                        Exit Do
                    End If
            
                    buffer(k, 1) = A(j, 1)
                    buffer(k, 2) = A(j, 2)
                    
                    j = j + 1
                    k = k + 1
            
                Loop
                
                poly = modOptimization.optPolyFit(buffer, polyOrder)
                
                C(i, 1) = A(moving_mid, 1)
                C(i, 2) = poly(mid, 2)
                moving_mid = moving_mid + 1
            
            Next i
    
            optSavGol = C
            Exit Function
    
        End If
        
End Function

Public Function optSavGolPeaks(ByRef data_2D#(), _
                               ByRef data_SG#(), _
                               ByRef data_fD#()) As tPeaks
                            
    Dim rawMaxRow2D&, _
        rawMaxCol2D&, _
        rawMaxRowSG&, _
        rawMaxColSG&, _
        rawMaxRowfD&, _
        rawMaxColfD&, _
        first#, _
        second#, _
        third#, _
        fourth#, _
        SG_x#, _
        SG_y#, _
        twoD_x#, _
        twoD_y#, _
        twoD_yy#, _
        peaksSG_x#(), _
        peaksSG_y#(), _
        peaks2D_x#(), _
        peaks2D_y#(), _
        final#(), _
        index&, _
        m&, p&, dw&, _
        i&, j&, k&

        rawMaxRow2D = UBound(data_2D, 1)
        rawMaxCol2D = UBound(data_2D, 2)
        rawMaxRowSG = UBound(data_SG, 1)
        rawMaxColSG = UBound(data_SG, 2)
        rawMaxRowfD = UBound(data_fD, 1)
        'rawMaxColfD = UBound(data_fD, 2)
        
        j = 1
        p = 1
        For i = 1 To (rawMaxRowfD - 3)
        
             first = data_fD(i + 0)
            second = data_fD(i + 1)
             third = data_fD(i + 2)
            fourth = data_fD(i + 3)
            
            If first > 0 And _
              second > 0 And _
               third < 0 And _
              fourth < 0 Then
              
                ReDim Preserve peaksSG_x(1 To j)
                ReDim Preserve peaksSG_y(1 To j)
                ReDim Preserve peaks2D_x(1 To j)
                ReDim Preserve peaks2D_y(1 To j)
              
                peaksSG_x(j) = data_SG((i + 1) + 1, 1)
                peaksSG_y(j) = data_SG((i + 1) + 1, 2)
                
                SG_x = peaksSG_x(j)
                SG_y = peaksSG_y(j)
                j = j + 1
                
                For m = 1 To rawMaxRow2D
                    
                    twoD_x = data_2D(m, 1)
                
                    If SG_x = twoD_x Then
                        index = m
                        Exit For
                    End If
                
                Next m
                
                dw = 1
                For k = 1 To (rawMaxRow2D - 1)
                
                    twoD_x = data_2D(k, 1)
                    twoD_y = data_2D(k, 2)
                    twoD_yy = data_2D(k + 1, 2)
                    
                    If k >= index - dw And _
                       k <= index + dw Then
                       
                       If twoD_y >= twoD_yy Then
                           
                           peaks2D_x(p) = twoD_x
                           peaks2D_y(p) = twoD_y
                           
                           p = p + 1
                           Exit For
                       
                       End If
                    
                    End If
                
                Next k
            
            End If
            
        Next i
        
        optSavGolPeaks.peaks_2D = modMatrix.matJoin(modMatrix.matDim(peaks2D_x), _
                                                    modMatrix.matDim(peaks2D_y))
        
        optSavGolPeaks.peaks_SG = modMatrix.matJoin(modMatrix.matDim(peaksSG_x), _
                                                    modMatrix.matDim(peaksSG_y))

End Function

Public Function optfD(ByRef dataX#(), _
                      ByRef dataY#()) As Double()

    Dim rawMaxRowY&, _
        rawMaxColY&, _
        i&, j&, _
        fD#()
        
        rawMaxRowY = UBound(dataY, 1)
        rawMaxColY = UBound(dataY, 2)

        ReDim fD(1 To rawMaxRowY - 2)
        
        For i = 2 To (rawMaxRowY - 1)
            
            fD(i - 1) = (dataY(i + 1, 1) - dataY(i - 1, 1)) / _
                        (dataX(i + 1, 1) - dataX(i - 1, 1))
                        
        Next i
        
        optfD = fD

End Function
