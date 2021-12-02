#' Location
#'
#' A Location is used to place an element on a slide.
#' @slot ph_location_fn function. Function in the \code{officer::ph_location*} family
#' @slot args list. Arguments to pass to \code{ph_location_fn}
#' @export
setClass(
  "R2PptxLocation",
  slots = c(
    ph_location_fn = "function",
    args = "list"
  )
)


#' New Location
#' @param ph_location_fn function. Function in the \code{officer::ph_location*} family
#' @param ... args to pass to \code{ph_location_fn}
#' @seealso \link[officer]{ph_location}
#' @examples
#' # create an element with some text offset 2 from the left and 2 from the top
#' # of the slide
#' element_location <- new_location(officer::ph_location, left = 2, top = 2)
#' element <- new_element(
#'   key = element_location,
#'   value = "Some text"
#' )
#' presentation <- new_presentation() +
#'   new_slide("Title Slide", elements = list(element))
#' path <- tempfile(fileext = ".pptx")
#' write_pptx(presentation, path)
#' if (interactive()) browseURL(path)
#' @export
new_location <- function(ph_location_fn, ...) {
  new("R2PptxLocation", ph_location_fn = ph_location_fn, args = list(...))
}


#' @keywords internal
setGeneric("ph_location_fn", function(x) standardGeneric("ph_location_fn"))

#' @keywords internal
setMethod(
  "ph_location_fn",
  "R2PptxLocation",
  function(x) x@ph_location_fn
)


#' @keywords internal
setGeneric("ph_location_args", function(x) standardGeneric("ph_location_args"))

#' @keywords internal
setMethod(
  "ph_location_args",
  "R2PptxLocation",
  function(x) x@args
)
