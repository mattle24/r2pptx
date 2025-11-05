#' Plot layout
#'
#' A thin wrapper around \code{officer::plot_layout_properties()} to plot
#' layouts for \code{R2PpptxPresentation} objects.
#' @param x \code{R2PpptxPresentation} object
#' @param layout character. Name of the layout to see properties for.
#' @export
#' @return No return value, called for side effects
plot_layout <- function(x, layout) {
  pptx_obj <- officer::read_pptx(template_path(x))
  officer::plot_layout_properties(pptx_obj, layout)
}


#' Get layouts
#'
#' A thin wrapper around \code{officer::layout_summary()} to get layouts for
#' \code{R2PpptxPresentation} objects.
#' @param x \code{R2PpptxPresentation} object
#' @export
#' @return An object of class \code{data.frame} with fields for the layout name
#'   and the name of the slide theme (master).
get_layouts <- function(x) {
  pptx_obj <- officer::read_pptx(template_path(x))
  officer::layout_summary(pptx_obj)
}


#' Get layout properties
#'
#' A thin wrapper around \code{officer::plot_layout_properties()} to get layouts
#' properties for \code{R2PpptxPresentation} objects.
#' @param x \code{R2PpptxPresentation} object
#' @param layout character. Name of the layout to see properties for.
#' @export
#' @return An object of class \code{data.frame} with fields for placeholder
#'   attributes and one row per placeholder element.
get_layout_properties <- function(x, layout) {
  pptx_obj <- officer::read_pptx(template_path(x))
  officer::layout_properties(pptx_obj, layout)
}

#' Get all layouts' properties
#'
#' Gets all the layout properties for all the layouts in a presentation
#'
#' @param x \code{R2PpptxPresentation} object
#' @export
#' @return An object of class \code{data.frame} with fields for placeholder
#'   attributes and one row per placeholder element per layout.
get_all_layout_properties <- function(x) {
  .officer_get_all_layout_properties(officer::read_pptx(template_path(x)))
}

#' Get all layouts' properties
#'
#' Gets all the layout properties for all the layouts in an officer presentation
#'
#' @param pptx_obj \code{officer::rpptx} object
#' @keywords internal
#' @return An object of class \code{data.frame} with fields for placeholder
#'   attributes and one row per placeholder element per layout.
.officer_get_all_layout_properties <- function(pptx_obj) {
  layouts <- officer::layout_summary(pptx_obj)
  x <- lapply(layouts$layout, \(layout) officer::layout_properties(pptx_obj, layout))
  Reduce(rbind, x)
}
