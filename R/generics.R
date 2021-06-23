#' @import methods
NULL

# write -------------------------------------------------------------------

setGeneric("write_pptx", function(x, path, ...) standardGeneric("write_pptx"))

# append slide ---------------------------------------------------------------

setGeneric("append_slide", function(e1, e2) standardGeneric("append_slide"))

# append element -------------------------------------------------------------

setGeneric("append_element", function(e1, e2) standardGeneric("append_element"))

