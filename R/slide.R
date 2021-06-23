#' @include generics.R
#' @include element.R
NULL

setClass(
  "R2PptxSlideList",
  slots = c(
    slides = "list"
  ),
  prototype = list(
    slides = list()
  )
)


#' Slide
#'
#' An S4 class to represent a powerpoint slide
#' @slot layout character. Name of the powerpoint layout to use for this
#'   slide.
#' @slot elements list. List of `R2PptxElement` objects.
#' @export
setClass(
  "R2PptxSlide",
  slots = c(
    layout = "character",
    elements = "list"
  )
)

#' New slide
#'
#' Make a `R2PptxSlide` object representing a powerpoint slide.
#' @param ... `R2PptxElements` to initialize the slide with
#' @param layout character. Name of the powerpoint layout to use for this
#'   slide.
#' @export
new_slide <- function(..., layout) {
  if (missing(layout)) {
    rlang::abort("`layout` was missing. See `officer::plot_layout_properties()` for key options.")
  }
  elements <- list(...)
  new("R2PptxSlide", layout = layout, elements = elements)
}


# length ------------------------------------------------------------------

setMethod("length", "R2PptxSlide", function(x) length(x@elements))


# add element -------------------------------------------------------------

setMethod(
  "append_element",
  signature = signature(e1 = "R2PptxSlide", e2 = "R2PptxElement"),
  function(e1, e2) {
    e1@elements <- append(e1@elements, e2)
    e1
  }
)

setMethod(
  "+",
  signature = signature(e1 = "R2PptxSlide", e2 = "R2PptxElement"),
  function(e1, e2) {
    append_element(e1, e2)
  }
)
