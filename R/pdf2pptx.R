
#' Convert PDF to PPTX
#'
#' @param pdf_filename File name of the source PDF file
#' @param pptx_filename File name of the destination PPTX file
#' @param dpi Optional parameter specifying image density for rendering images from PDF
#' @param path Directory where to put rendered images. If `NULL`, temporary folder is used and the images are deleted.
#'
#' @return Nothing
#' @export
#'
#' @examples
#' example_from_url <-
#'   "http://mirrors.ctan.org/macros/latex/contrib/beamer-contrib/themes/metropolis/demo/demo.pdf"
#' # conversion takes several seconds
#' pdf2pptx(example_from_url, "demo.pptx")
#' unlink("demo.pptx")
pdf2pptx <- function(
  pdf_filename,
  pptx_filename,
  dpi = 300,
  path = NULL
  ) {

  if (is.null(path)) {
    folder_for_files <- file.path(tempdir(), "pdf_images")
    dir.create(folder_for_files)
  } else {
    folder_for_files <- path
  }

  # turn pdf into png files
  img_pdf <- magick::image_read_pdf(pdf_filename, density = dpi)
  img_png <- magick::image_convert(img_pdf, format = "png")
  n_slides <- length(img_pdf)
  slide_filenames <-
    file.path(folder_for_files, sprintf("slide_%04d.png", 1:n_slides))
  for (i in 1:n_slides) {
    magick::image_write(
      img_png[i],
      slide_filenames[i]
    )
  }

  # insert png files into pptx
  pptx <- officer::read_pptx()
  # for info use layout_summary(pptx)
  for (i in 1:n_slides) {
    pptx <- officer::add_slide(pptx, layout = "Blank", master = "Office Theme")
    pptx <- officer::ph_with(
      pptx,
      officer::external_img(slide_filenames[i]),
      location = officer::ph_location_fullsize())
  }
  print(pptx, target = pptx_filename)

  if (is.null(path)) {
    unlink(folder_for_files, recursive = T)
  }
}
