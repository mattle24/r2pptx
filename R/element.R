#' @include generics.R
#' @include location.R
NULL


#' Element
#'
#' An S4 An class to represent text, a flextable, an image, a ggplot2, etc we
#' add to a slide.
#' @export
setClass(
  "R2PptxElement",
  slots = c(
    key = "R2PptxLocation",
    value = "ANY"
  )
)


#' New Element
#'
#' Make a new `R2PptxElement`. Element represent text, a flextable, an image, a
#' ggplot2, etc to add to a slide.
#' @param key character. Name of the placeholder label for this element.
#' @param value object. Object to put into a PowerPoint slide, eg text or a plot.
#' @export
new_element <- function(key, value) {
  if (missing(key)) {
    stop("`key` was missing. See `officer::layout_summary()` for layout options.")
  }

  new("R2PptxElement", key = asR2PptxLocation(key), value = value)
}
