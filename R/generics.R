#' @import methods
NULL


# super class -----------------------------------------------------------

setClass("R2Pptx")

# write -------------------------------------------------------------------

#' Write pptx
#'
#' Write an object to a `.pptx` file.
#' @param x object
#' @param path character. File path to write to.
#' @param add_slide_numbers logical. Default TRUE. Adds dynamic slide numbers
#' @param start_slide integer. Default 1. First slide to add slide numbers to
#' @export
setGeneric("write_pptx", function(x, path, add_slide_numbers = TRUE, start_slide = 1) {
  standardGeneric("write_pptx")
})

# append slide ---------------------------------------------------------------

setGeneric("append_slide", function(e1, e2) standardGeneric("append_slide"))

# append element -------------------------------------------------------------

setGeneric("append_element", function(e1, e2) standardGeneric("append_element"))

