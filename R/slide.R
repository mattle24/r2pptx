#' @include generics.R
#' @include element.R
NULL


#' Slide
#'
#' An S4 class to represent a powerpoint slide
#' @slot layout character. Name of the powerpoint layout to use for this
#'   slide.
#' @slot elements list. List of `R2PptxElement` objects.
#' @export
setClass(
  "R2PptxSlide",
  contains = "R2Pptx",
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


# slide list --------------------------------------------------------------

#' Slide list
#'
#' `R2PptxSlideList` is an S4 class to contain groups of `R2PptxSlide` objects
#' that are not part of a presentation. It is meant to be used to create lists
#' of slides and then be able to add the list easily to a presentation.
#' @slot slides list. A list of `R2PptxSlide` objects
setClass(
  Class = "R2PptxSlideList",
  contains = "R2Pptx",
  slots = c(
    slides = "list"
  ),
  prototype = list(
    slides = list()
  )
)


setValidity("R2PptxSlideList", function(object) {
  if (!all(purrr::map_lgl(object@slides, ~ inherits(., "R2PptxSlide")))) {
    "Each element in the list must be a `R2PptxSlide` object"
  } else {
    TRUE
  }
})

setGeneric("asSlideList", function(x) standardGeneric("asSlideList"))
setMethod(
  "asSlideList",
  "list",
  function(x) {
    new("R2PptxSlideList", slides = x)
  }
)
setMethod(
  "asSlideList",
  "R2PptxSlide",
  function(x) {
    asSlideList(list(x))
  }
)

# method to be able to append slides to each other in a slide list
# two slides
setMethod(
  "append_slide",
  signature(e1 = "R2PptxSlide", e2 = "R2PptxSlide"),
  function(e1, e2) {
    asSlideList(list(e1, e2))
  }
)

# slide list and a slide
setMethod(
  "append_slide",
  signature(e1 = "R2PptxSlideList", e2 = "R2PptxSlide"),
  function(e1, e2) {
    e1@slides <- append(e1@slides, e2)
    validObject(e1)
    e1
  }
)

# slide and a slide list
setMethod(
  "append_slide",
  signature(e1 = "R2PptxSlide", e2 = "R2PptxSlideList"),
  function(e1, e2) {
    e2@slides <- append(e1, e2@slides)
    validObject(e2)
    e2
  }
)

