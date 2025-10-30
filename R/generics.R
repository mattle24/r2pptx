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
#' @param add_slide_numbers logical. If TRUE (default), adds dynamic slide
#'   numbers to slides. Set to FALSE to disable.
#' @param start_slide integer. First slide to add numbers to (default 1).
#' @export
setGeneric("write_pptx", function(x, path, add_slide_numbers = TRUE, start_slide = 1) {
  standardGeneric("write_pptx")
})

# append slide ---------------------------------------------------------------

setGeneric("append_slide", function(e1, e2) standardGeneric("append_slide"))

# append element -------------------------------------------------------------

setGeneric("append_element", function(e1, e2) standardGeneric("append_element"))

