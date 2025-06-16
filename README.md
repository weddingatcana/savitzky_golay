# Savitzky Golay Filter

When working on various projects, I tend to have the need to smooth data or evaluate trends. Since having completed work on polynomial fitting, employing the Savitzky-Golay filter followed naturally. There's a great wikipedia article detailing the specifics of the filter, [found here](https://en.wikipedia.org/wiki/Savitzky%E2%80%93Golay_filter), and for those curious enough, and having journal access, you can check out the seminal paper [also found here](https://pubs.acs.org/doi/abs/10.1021/ac60214a047). 

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

As an optional step, assuming ***csvMatrix*** has more than two columns of data - (x,y), we can individually specify what our chosen x and y arrays will be. These individual arrays are dimensioned as Nx1, (rxc). So, for instance, if you had ***csvMatrix*** having three columns of data - (x, y1, y2), and we only wanted to perform a fit to (x,y2) then we would pull out individual vectors as such:

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

## Extra Features

### Downsampling

### Peak Finding

### Window Sizing & Polynomial Order


## Notes
