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

# length

#' get slide length (elements)
#' @rdname length
setMethod("length", "R2PptxSlide", function(x) length(x@elements))

# show method
setMethod(
  "show",
  "R2PptxSlide",
  function(object) {
    element_types <- sapply(object@elements, function(x) class(x@value)[1])
    cat(glue::glue(
      "Slide with layout `{l}` and {n} elements:\n",
      paste("- ", element_types, collapse = "\n"),
      l = object@layout,
      n = length(object)
    ))
  }
)

#' New slide
#'
#' Make a `R2PptxSlide` object representing a powerpoint slide.
#' @param layout character. Name of the powerpoint layout to use for this
#'   slide.
#' @param elements list. List of `R2PptxElements` to initialize the slide with.
#'   Defaults to empty list.
#' @export
new_slide <- function(layout, elements = list()) {
  if (missing(layout)) {
    stop("`layout` was missing. See `officer::plot_layout_properties()` for key options.")
  }
  if (is(elements, "R2PptxElement")) {
    elements <- list(elements)
  }
  new("R2PptxSlide", layout = layout, elements = elements)
}


# length ------------------------------------------------------------------

#' Get slide length
#'
#' Returns the number of elements in a slide
#' @rdname length
#' @param x `R2PptxSlide` object
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

#' Add element to slide
#'
#' Add an `R2PptxElement` object to an `R2PptxSlide` object.
#' @param e1 `R2PptxSlide` object
#' @param e2 `R2PptxElement` object
#' @export
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
  if (!all(sapply(object@slides, function(x) inherits(x, "R2PptxSlide")))) {
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

setGeneric("get_slides", function(x) standardGeneric("get_slides"))
#' method to get slides
setMethod(
  "get_slides",
  "R2PptxSlideList",
  function(x) x@slides
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


#' Add slide to slidelist
#'
#' @param e1 `R2PptxSlide` object
#' @param e2 `R2PptxSlideList` object
#' @export
setMethod(
  "+",
  signature = signature(e1 = "R2PptxSlide", e2 = "R2PptxSlideList"),
  function(e1, e2) {
    append_slide(e1, e2)
  }
)
