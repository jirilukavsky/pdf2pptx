# pdf2pptx

Simple R script for converting PDF presentations into PPTX format (based on images).

## Usage

Source the `pdf2pptx.R` file into your environment. This will require several packages (`officer`, `magick`, `pdftools`) and one function called `pdf2pptx`. 

To use:

```
pdf2pptx(pdf_filename, pptx_filename)
```

You can specify image density with optional parameter `density = 300`.