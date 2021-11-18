#' Location
#'
#' A Location is used to place an element on a slide.
setClass(
  "R2PptxLocation",
  slots = c(
    ph_location_fn = "function",
    args = "list"
  )
)


R2PptxLocation <- function(ph_location_fn, args) {
  new("R2PptxLocation", ph_location_fn = ph_location_fn, args = args)
}


#' @param ph_location_fn function. Function in the \code{officer::ph_location*} family
#' @param ... args to pass to \code{ph_location_fn}
new_location <- function(ph_location_fn, ...) {
  new("R2PptxLocation", ph_location_fn = ph_location_fn, args = list(...))
}
