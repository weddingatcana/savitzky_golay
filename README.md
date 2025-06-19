# Savitzky Golay Filter

When working on various projects, I tend to have the need to smooth data or evaluate trends. Since having completed work on polynomial fitting, employing the Savitzky-Golay filter followed naturally. There's a great wikipedia article detailing the specifics of the filter, [found here](https://en.wikipedia.org/wiki/Savitzky%E2%80%93Golay_filter), and for those curious enough, *and having journal access*, you can check out the seminal paper [also found here](https://pubs.acs.org/doi/abs/10.1021/ac60214a047). 

However, simply put, the filter will take a noisy signal and have you define a moving window from which the points inside will be fitted with a polynomial of order specified by you. Once fit, you take the center point from the fitted polynomial as the first point to your filtered data. Indexing the window by one will move you down the noisy signal and repeat the process until there is no more data to filter.

This repository consists of five modules, four with the **mod** prefix to denote separate libraries of functions and the fifth with the **prg** prefix to denote the main program where we will filter the noisy data.

## Getting Started

Again, just as discussed in the **polynomial_fitting** repository, the noisy input data will only be accepted as a **.csv** file. Just as before, it's important to note that you are not restricted to importing only two columns of data, nor any specific columnar formatting. However, this does not support headers, so the data should be numeric only. 

```VBA
'choose the .csv file you'd like
csvFilepath = modText.csvFind
```

Parsing the file, and loading it into an array:

```VBA
csvMatrix = modText.csvParse(csvFilepath)
```

As an optional step, assuming ***csvMatrix*** has more than two columns of data - (x,y), we can individually specify what our chosen x and y arrays will be. These individual arrays are dimensioned as Nx1, (rxc). So, for instance, if you had ***csvMatrix*** having three columns of data - (x, y1, y2), and we only wanted to filter (x,y2) then we would pull out individual vectors as such:

```VBA
'Separate x and y arrays from larger csvMatrix, if needed
dataX = modMatrix.matVec(csvMatrix, 1)
dataY = modMatrix.matVec(csvMatrix, 3)
```

We'd then recombine both vectors into a single 2D array with specific columnar formatting - (x,y2):

```VBA
data_2D = modMatrix.matJoin(dataX, dataY)
```

It's important to note that our Savitzky-Golay function assumes the input array is of the form - (x,y). Which is why we've specified the columnar formatting as seen above. With our data formatted appropriately let's turn our attention to the filter, ***optSavGol***.

```VBA
data_SG = modOptimization.optSavGol(data_2D, 11, 3)
```

The function above, ***optSavGol*** has three inputs - the 2D data, the size of the window (the number of points we'll fit at a time), and the order of the polynomial, respectively. For this example, we've chosen a window size of eleven points and a third order polynomial. Now, it should be noted, that to uniquely determine an nth order polynomial one will minimally need (n+1) points. Let's explain a bit more.

A straight line, or a first order polynomial is of the form:

y(x) = a\*x + b

We have two unknowns, a and b. To solve a linear system of equations with two unknowns we'll need, minimally, two equations (aka, two points). If less, the system is underdetermined, and if greater the system is overdetermined. Let's look at a second order polynomial which is of the form:

y(x) = a\*x^2 + b\*x + c

This time we have three unknowns, a, b and c. Just as before, with three unknowns we'll need three equations to solve the system uniquely - otherwise we have infinitely many solutions. Therefore, an nth order polynomial needs (n+1) points. There are some checks and balances within ***optSavGol*** on the window size and polynomial order, but just remember that keeping a healthy seperation between the two values will lead to better numerical stability.

Now, once the function has completed, we now seek to export the filtered data, ***data_SG***, and perhaps other fields. We'll define a boolean variable, ***csvStatus***, to display true/false if the exporting was successful. We can see this in practice below:

```VBA
'Write the various results to written filepaths
csvStatus = modText.csvWrite(csvMatrix, "raw.csv")
csvStatus = modText.csvWrite(data_2D, "xy.csv")
csvStatus = modText.csvWrite(data_SG, "filtered.csv")
```

## Extra Features

### Downsampling

As previously discussed in the **polynomial_fitting** repository, we allow for data to be downsampled in order to reduce the time required to calculate our filtered data. The function, ***mathDownSampling***, takes only one dimension of data at a time - x or y etc. More specifics can be found in the aforementioned repository, so for now we'll just focus on the implementation:

```VBA
'Separate x and y arrays from larger csvMatrix, just as before
dataX = modMatrix.matVec(csvMatrix, 1)
dataY = modMatrix.matVec(csvMatrix, 3)

'Perform downsampling, let's say every other point
dataX_dS = modMath.mathDownSampling(dataX, 2)
dataY_dS = modMath.mathDownSampling(dataY, 2)

'Recombining again
data_2D = modMatrix.matJoin(dataX_dS, dataY_dS)
````

### Peak Finding

Generally, when trying to smooth noisy data, you're looking for trends and points of inflection. It proves rather useful in analyses to find local/global peaks of said trends. Trying to find the decay rate for a damped sinusoid? Find the peaks and fit an exponential. Looking to find the cycle time for a noisy, complicated waveform? *P e a k s.*

The aptly named, ***optSavGolPeaks***, function will do just that. It takes three input arrays, respectively:
  
  1) The raw data set, ***data_2D***.
  2) The filtered data set, ***data_SG***.
  3) The derivative data set, ***data_fD***.

Since we're dealing with discrete data, we don't have the luxury of selecting peaks by simply finding when the first derivative is zero.  We'll need to compute the first derivative by way of a first order central finite difference then check when there is a sign change. If you imagine the slope of a single hill, the slope changes from positive to zero then negative. With our array of differences, ***data_fD***, we seek when a linear sequence of elements within the array are of the pattern: (+,+,-,-). If that occurs we know that our peak would be represented by the second element (second +).

The normal route of using this function is to calculate differences from the smoothed data set, ***data_SG***. From there, peaks will be found using the above methodology. Using these smoothed peaks we will then loop through the raw data set and evaluate a small window range of values near where the smoothed peaks were located, to try and find the analgous raw peaks.

We'd program this as such, starting from ***data_SG***:

```VBA
'Generate smoothed data by Savitsky-Golay filter
data_SG = modOptimization.optSavGol(data_2D, 11, 3)

'Separate x,y data for use with optfD
dataX_SG = modMatrix.matVec(data_SG, 1)
dataY_SG = modMatrix.matVec(data_SG, 2)

'Calculate finite difference array
data_fD = modOptimization.optfD(dataX_SG, dataY_SG)

'Find peaks of data_2D and data_SG
peaks = modOptimization.optSavGolPeaks(data_2D, data_SG, data_fD)

'Write the various results to written filepaths
csvStatus = modText.csvWrite(data_2D, "xy.csv")
csvStatus = modText.csvWrite(data_SG, "smoothed.csv")
csvStatus = modText.csvWrite(peaks.peaks_2D, "xy_peaks.csv")
csvStatus = modText.csvWrite(peaks.peaks_SG, "smoothed_peaks.csv")
````

You'll notice the variable **peaks** returns two arrays, this is because it is a user defined type, **tPeaks**. Which returns both the smoothed peaks and the raw peaks. 

It is important to note that you could also simply use the ***optSavGolPeaks*** function to just return peaks from data that hasn't been smoothed. Merely generate a difference array of the data that hasn't been smoothed, and have both 1) & 2) inputs of the function be equal. An example of this is below:

```VBA
'Separate x and y arrays from larger csvMatrix, if needed
dataX = modMatrix.matVec(csvMatrix, 1)
dataY = modMatrix.matVec(csvMatrix, 3)

'Combine x,y data into one array
data_2D = modMatrix.matJoin(dataX, dataY))

'Calculate finite difference array of non-smoothed data
data_fD = modOptimization.optfD(dataX, dataY)

'Find peaks of data_2D (raw)
peaks = modOptimization.optSavGolPeaks(data_2D, data_2D, data_fD)

'Write the various results to written filepaths
csvStatus = modText.csvWrite(data_2D, "xy.csv")
csvStatus = modText.csvWrite(peaks.peaks_2D, "xy_peaks_no_smooth.csv")
````

### Window Sizing & Polynomial Order


## Notes
