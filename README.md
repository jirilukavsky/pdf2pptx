
# pdf2pptx

<!-- badges: start -->
<!-- badges: end -->

The package pdf2pptx is a simple tool for converting PDF presentations into PPTX format (based on images). 

It is implemented in R to easily extend [R Markdown](https://bookdown.org/yihui/rmarkdown/beamer-presentation.html) workflow. This way you can convert your presentation to PPTX and easily present on platforms like MS Teams, where PDF presentation is problematic.

Under the hood, the package uses `officer`, `magick`, `pdftools` packages.

## Installation

You can install the released version of pdf2pptx from [GitHub](https://github.com/jirilukavsky/pdf2pptx) with:

``` r
# Install devtools package if necessary
if(!"devtools" %in% rownames(installed.packages())) install.packages("devtools")

# Install the stable verion from GitHub
devtools::install_github("jirilukavsky/pdf2pptx")
```

## Example

To convert a PowerPoint file, simply call `pdf2pptx` command and specify the source and destination file names:

``` r
library(pdf2pptx)
pdf2pptx("your.pdf", "your.pptx")
```

