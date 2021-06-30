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
#' @export
setGeneric("write_pptx", function(x, path) standardGeneric("write_pptx"))

# append slide ---------------------------------------------------------------

setGeneric("append_slide", function(e1, e2) standardGeneric("append_slide"))

# append element -------------------------------------------------------------

setGeneric("append_element", function(e1, e2) standardGeneric("append_element"))

