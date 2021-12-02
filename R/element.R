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
    # TODO use `location` not key and deprecate `key`
    key = "R2PptxLocation",
    value = "ANY"
  )
)


#' New Element
#'
#' Make a new `R2PptxElement`. Element represent text, a flextable, an image, a
#' ggplot2, etc to add to a slide.
#' @export
#' @return An object of class \code{R2PptxElement} representing something to put
#'   on a slide.
setGeneric("new_element", function(key, value) standardGeneric("new_element"))

#' New Element
#'
#' Make a new `R2PptxElement`. Element represent text, a flextable, an image, a
#' ggplot2, etc to add to a slide.
#' @param key character. Name of the placeholder label for this element.
#' @param value object. Object to put into a PowerPoint slide, eg text or a plot.
#' @export
#' @return An object of class \code{R2PptxElement} representing something to put
#'   on a slide.
setMethod(
  "new_element",
  "character",
  function(key, value) {
    location <- new_location(officer::ph_location_label, ph_label = key)
    new("R2PptxElement", key = location, value = value)
  }
)

#' New Element
#'
#' Make a new `R2PptxElement`. Element represent text, a flextable, an image, a
#' ggplot2, etc to add to a slide.
#' @param key `R2PptxLocation` object created with \code{new_location()}
#' @param value object. Object to put into a PowerPoint slide, eg text or a plot.
#' @export
#' @return An object of class \code{R2PptxElement} representing something to put
#'   on a slide.
setMethod(
  "new_element",
  "R2PptxLocation",
  function(key, value) {
    new("R2PptxElement", key = key, value = value)
  }
)

#' @keywords internal
setGeneric("value", function(x) standardGeneric("value"))

#' Get value of element
#' @keywords internal
setMethod(
  "value",
  "R2PptxElement",
  function(x) x@value
)

#' @keywords internal
setGeneric("location", function(x) standardGeneric("location"))

#' Get location (key) of element
#' @keywords internal
setMethod(
  "location",
  "R2PptxElement",
  function(x) x@key
)


# TODO rename this generic / method
setGeneric("add_pptx", function(x, pptx_obj) standardGeneric("add_pptx"))
setMethod(
  "add_pptx",
  "R2PptxElement",
  function(x, pptx_obj) {
    location <- location(x)
    officer::ph_with(
      pptx_obj,
      value = value(x),
      location = do.call(ph_location_fn(location), ph_location_args(location))
    )
  }
)
