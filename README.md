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

The function above, ***optSavGol*** has three inputs - the 2D data, the size of the window (the number of points we'll fit at a time), and the order of the polynomial, respectively. For this example, we've chosen a window size of eleven points and a third order polynomial. When selecting the window size we must always choose an odd number. The rationale is that if we were to have an even window size, the middle point selected would yield an x data point that wouldnt correspond with any x data point within the raw ***data_2D*** data set. We'd be selecting a point in-between unique x data points. Now, it should be noted, that to uniquely determine an nth order polynomial one will minimally need (n+1) points. Let's explain a bit more.

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

Since we're dealing with discrete data, we don't have the luxury of selecting peaks by simply finding when the first derivative is zero.  We'll need to compute the first derivative by way of a first order central finite difference then check when there is a sign change. If you imagine the slope of a single hill, from left to right, the slope changes from positive to zero then negative. With our array of differences, ***data_fD***, we seek to find when a linear sequence of elements within the array are of the pattern: (+,+,-,-). If that occurs we know that our peak would be represented by the second element (second +).

The normal route of using this function is to calculate differences from the smoothed data set, ***data_SG***. From there, peaks will be found using the above methodology. Using these smoothed peaks we will then loop through the raw data set and evaluate a small window range of values near where the smoothed peaks were located, to try and find the analgous raw peaks.

We'd program this as such, starting from ***data_SG***:

```VBA
'Generate smoothed data from filter
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

You'll notice the variable **peaks** returns two arrays, this is because it is of a user defined type, **tPeaks**. Which returns both the smoothed peaks and the raw peaks. 

It is important to note that you could also simply use the ***optSavGolPeaks*** function to just return peaks from data that won't be smoothed. Merely generate a difference array of the data that hasn't been smoothed, and have both inputs, 1) & 2), of the function be equal. An example of this is below:

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

### Padding

There exists a fourth, optional, argument for the ***optSavGol*** function that allows for padding. To explain, when fitting a polynomial to a window of raw data points we will select the middle point of the window as our filtered data. However, as a consequence, points from start and end of any raw data set, will be excluded from the filtered data set. Here's a toy example, given starting x,y data:

**dataY** = (2,4,8,9,8,4,2)

**dataX** = (1,2,3,4,5,6,7)

So our starting data array's contain seven elements total. Let's say our window is three, so starting off the algorithm we'll evaluate the starting data within the window:

**dataY_window** = (2,4,8)

**dataX_window** = (1,2,3)

Let's say we make a fit through **dataY_window**, our fitted data would be, generally:

**dataY_window_fit** = (f1,f2,f3)

Then, selecting our middle point (x,y) we get (2,f2). Therefore, we have completely dropped the first data point of our starting data set. Similarly if we went through the motions we'd find that we'll also drop the final point of our starting data. So, to find how many points will be dropped from any starting data set, as a function of window size, we find:

**dropped_points**(window) = window - 1

To continue futher, the number of points dropped at either end of the starting data set is (**dropped_points**/2). So, with that explained, if we want to have the starting data and filtered data sets to be of the same length, we'll just include the points left out. The points included will be points corresponding to a polynomial fit, not the raw points. To show this implemented:

```VBA
data_SG = modOptimization.optSavGol(data_2D, 11, 3, padding)
```

The fourth argument is by default padded, however using the enumeration provided you can also select ***no_padding***:

```VBA
data_SG = modOptimization.optSavGol(data_2D, 11, 3, no_padding)
```

## Notes

The filter can, under certain conditions, generate nonsensical data. Essentially, by having a polynomial order that is too close to a given window size, you are forcing the filter to create a highly complex curve that is overly sensitive to noise, leading to the observed nonsensical data. To mitigate this, ensure your window size is significantly larger than your polynomial order.

Anecdotally, I have noticed that larger data sets are more sensitive to this overfitting, ill-conditioned phenomena. The results that I've witnessed always start out with the filtered data providing a good result, and then gradually starts going off the rails gradually outputting data well beyond the dependent axis bounds from raw data. It's as if the noise is being fitted and gradually compounding over the length of the data set.
