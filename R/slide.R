#' @include generics.R
#' @include element.R
NULL


#' Slide
#'
#' An S4 class to represent a PowerPoint slide
#' @slot layout character. Name of the PowerPoint layout to use for this
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
#' Make a `R2PptxSlide` object representing a PowerPoint slide.
#' @param layout character. Name of the PowerPoint layout to use for this
#'   slide.
#' @param elements list. List of `R2PptxElements` to initialize the slide with.
#'   Defaults to empty list.
#' @export
#' @return An object of class \code{R2PptxSlide} representing a future
#'   PowerPoint slide.
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
#' @return Integer, number of elements in the \code{R2PptxSlide} object
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
#' Add an \code{R2PptxElement} object to an \code{R2PptxSlide} object.
#' @param e1 \code{R2PptxSlide} object
#' @param e2 \code{R2PptxElement} object
#' @export
#' @return An object of class \code{R2PptxSlide}, which is \code{e1} with an
#' additional element \code{e2}
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
#' \code{R2PptxSlideList} is an S4 class to contain groups of \code{R2PptxSlide}
#' objects that are not part of a presentation. It is meant to be used to create
#' lists of slides and then be able to add the list easily to a presentation.
#' @slot slides list. A list of \code{R2PptxSlide} objects
#' @export
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


#' New slide list
#'
#' Make a \code{R2PptxSlideList} object representing a list of PowerPoint slides
#' @param slides list. List of \code{R2PptxSlide} objects to initialize the list with.
#'   Defaults to empty list.
#' @export
#' @return An object of class \code{R2PptxSlideList} representing a list of
#'   \code{R2pptxSlide} object.
new_slidelist <- function(slides = list()) {
  if (is(slides, "R2PptxSlide")) {
    slides <- list(slides)
  }
  new("R2PptxSlideList", slides = slides)
}


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
#' @param x \code{R2PptxSlideList} object
#' @returns List, a list of \code{R2PptxSlide} objects.
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
#' @param e1 \code{PR2PptxSlide} object
#' @param e2 \code{R2PptxSlideList} object
#' @export
#' @return An object of class \code{R2PptxSlideList} which is the
#'   \code{R2PptxSlideList} list \code{e1} with an additional slide which is
#'   \code{e2}.
setMethod(
  "+",
  signature = signature(e1 = "R2PptxSlide", e2 = "R2PptxSlideList"),
  function(e1, e2) {
    append_slide(e1, e2)
  }
)


#' Add slidelist to slidelist
#'
#' @param e1 \code{R2PptxSlideList} object
#' @param e2 \code{R2PptxSlideList} object
#' @export
#' @return An object of class \code{R2PptxSlideList} which is \code{e1} with
#'   additional slides (all the slides in \code{e2}).
setMethod(
  "+",
  signature = signature(e1 = "R2PptxSlideList", e2 = "R2PptxSlideList"),
  function(e1, e2) {
    for (slide in get_slides(e2)) {
      e1 <- append_slide(e1, slide)
    }
    e1
  }
)
