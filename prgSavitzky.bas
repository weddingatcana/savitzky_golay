Attribute VB_Name = "prgSavitzky"
Option Explicit

Sub PolyFit()

    Dim csvStatus As Boolean, _
        csvFilepath$, _
        csvMatrix#(), _
 _
        dataX#(), _
        dataX_dS#(), _
        dataX_SG#(), _
        dataY#(), _
        dataY_dS#(), _
        dataY_SG#(), _
        data_dS#(), _
        data_2D#(), _
        data_SG#(), _
        data_fD#(), _
 _
        window&, _
        peaks As tPeaks

        'Choose csv file to use, multiple columns fine
        csvFilepath = modText.csvFind
        
        'If nothing chosen, close program
        If Len(csvFilepath) = 0 Then
            Exit Sub
        End If
        
        'Pull data from csv file into array, no assumption of columnar formatting
        csvMatrix = modText.csvParse(csvFilepath)
        
        'Separate x and y arrays from larger csvMatrix, if needed
        dataX = modMatrix.matVec(csvMatrix, 1)
        dataY = modMatrix.matVec(csvMatrix, 2)
        
        'Combine separate x and y arrays for use in optSavGol which takes 2D (x,y) array
        data_2D = modMatrix.matJoin(dataX, dataY)
   
        'Generate smoothed data by savitsky-golay filter
        data_SG = modOptimization.optSavGol(data_2D, 67, 2)
        
        dataX_SG = modMatrix.matVec(data_SG, 1)
        dataY_SG = modMatrix.matVec(data_SG, 2)
        
        'Calculate central first derivative of data_2D
        data_fD = modOptimization.optfD(dataX_SG, dataY_SG)
        
        'Find peaks of data_2D and data_SG
        peaks = modOptimization.optSavGolPeaks(data_2D, data_SG, data_fD)
        
        
        'Write the various results to written filepaths
        csvStatus = modText.csvWrite(data_2D, "xy.csv")
        csvStatus = modText.csvWrite(data_SG, "smoothed.csv")
        csvStatus = modText.csvWrite(peaks.peaks_2D, "xy_peaks.csv")
        csvStatus = modText.csvWrite(peaks.peaks_SG, "smoothed_peaks.csv")

End Sub
