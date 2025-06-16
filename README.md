# Savitzky Golay Filter

When working on various projects, I tend to have the need to smooth data or evaluate trends. Since having completed work on polynomial fitting, employing the Savitzky-Golay filter followed naturally. There's a great wikipedia article detailing the specifics of the filter, [found here](https://en.wikipedia.org/wiki/Savitzky%E2%80%93Golay_filter), and for those curious enough, and having journal access, you can check out the seminal paper [also found here](https://pubs.acs.org/doi/abs/10.1021/ac60214a047). 

However, simply put, the filter will take a noisy signal and have you define a moving window from which the points inside will be fitted with a polynomial of order specified by you. Once fit, you take the center point from the fitted polynomial as the first point to your filtered data. Indexing the window by one will move you down the noisy signal and repeat the process until there is no more data to filter.

## Getting Started

## Extra Features

### Downsampling

### Peak Finding

### Window Sizing & Polynomial Order


## Notes
