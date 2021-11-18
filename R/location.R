#' Placeholder Location
#'
#' A Location is used to place an element on a slide. This class handles
#' locations that are the names of PowerPoint placeholders.
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


setGeneric("new_location", function(x, ...) standardGeneric("new_location"))
setMethod(
  "new_location",
  "character",
  function(x) R2PptxLocation(ph_location_fn = officer::ph_location_label, args = list(ph_label = x))
)

setMethod(
  "new_location",
  "function",
  function(x, ...) {
    args <- list(...)
    R2PptxLocation(ph_location_fn = x, args = args)
  }
)

