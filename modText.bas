Attribute VB_Name = "modText"
Option Explicit

Public Type tData
    x As Double
    y As Double
End Type

Public Function csv2D(ByVal csvLine$) As tData

    Dim result As tData, _
        delimiter$, _
        position$, _
        length&

        delimiter = ","
        length = Len(csvLine)
        position = InStr(1, csvLine, delimiter)
        
        result.x = CDbl(Left(csvLine, (position - 1)))
        result.y = CDbl(Right(csvLine, (length - position)))
        
        csv2D = result

End Function

Public Function csvND(ByVal csvLine$) As Double()

    Dim char$(), _
        result#(), _
        rawMaxRow&, _
        delimiter$, _
        i&

        delimiter = ","
        char = Split(csvLine, delimiter)
        
        rawMaxRow = UBound(char)
        ReDim result(0 To rawMaxRow)
        
        For i = 0 To rawMaxRow
            result(i) = char(i)
        Next i
        
        csvND = result

End Function

Public Function csvFind$()

    Dim FDO As FileDialog, _
        SelectionChosen&
    
        Set FDO = Application.FileDialog(msoFileDialogFilePicker)
        SelectionChosen = -1
        
        With FDO
            .InitialFileName = "C:\"
            .Title = "Choose CSV"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Allowed File Extensions", "*.csv"
            
            If .Show = SelectionChosen Then
                csvFind = .SelectedItems(1)
            Else
            End If
            
        End With
    
        Set FDO = Nothing

End Function

Public Function csvParse(ByVal csvFilepath$) As Double()

    Dim fileObject As Object, _
        textObject As Object, _
        data2D As tData, _
        dataColumns&, _
        dataND#(), _
        dataLine$, _
        i&, j&, _
        C#()
        
        Set fileObject = CreateObject("Scripting.FileSystemObject")
        'Set textObject = CreateObject("Scripting.TextStream")
        Set textObject = fileObject.OpenTextFile(csvFilepath)
        
        With textObject
        
            dataLine = .readline
            dataND = modText.csvND(dataLine)
            'Split returns a zero based array, hence +1
            dataColumns = UBound(dataND) + 1
            
            'previously was i = 0; however, once I included the readline above to find columns, had to
            'start at 1 because the C array would have one row less than what was required.
            i = 1
            Do
                If .AtEndOfStream Then
                    Exit Do
                End If
            
                i = i + 1
                .SkipLine
            Loop
            
            .Close
            ReDim C(1 To i, 1 To dataColumns)
        
        End With
        
        Set textObject = fileObject.OpenTextFile(csvFilepath)
        
        With textObject
        
            i = 1
            Do
            
                If .AtEndOfStream Then
                    Exit Do
                End If
                
                dataLine = .readline
                
                If dataColumns = 2 Then
                    'Redundant, but kept in anyway
                    data2D = modText.csv2D(dataLine)
                    C(i, 1) = data2D.x
                    C(i, 2) = data2D.y
                    
                Else
                
                    dataND = modText.csvND(dataLine)
                    For j = 0 To (dataColumns - 1)
                        'Again, +1 because of Split
                        C(i, j + 1) = dataND(j)
                    Next j
                
                End If
                
                i = i + 1
            
            Loop
            .Close
        
        End With
        
        Set fileObject = Nothing
        Set textObject = Nothing
        
        csvParse = C
        
End Function

Public Function csvWrite(ByRef A#(), _
                         ByVal csvFilename$, _
                Optional ByVal csvDirectory$ = "C:\Users\qp\Desktop\") As Boolean

    Dim FSO As Object, _
        txtFile As Object, _
        rawMaxRowA&, _
        rawMaxColA&, _
        concatString$, _
        delimiter$, _
        i&, j&
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set txtFile = FSO.CreateTextFile(csvDirectory & csvFilename)
        delimiter = ","
        
        rawMaxRowA = UBound(A, 1)
        rawMaxColA = UBound(A, 2)
       
        For i = 1 To rawMaxRowA
            For j = 1 To rawMaxColA
            
                concatString = concatString & A(i, j) & delimiter
                
            Next j
            
            txtFile.Write concatString & vbCrLf
            concatString = ""
            
        Next i
        
        txtFile.Close
        csvWrite = True
        Set FSO = Nothing
        Set txtFile = Nothing

End Function
