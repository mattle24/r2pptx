#' Plot layout
#'
#' A thin wrapper around `officer::plot_layout_properties()` to plot
#' layouts for `R2PpptxPresentation` objects.
#' @param x R2PpptxPresentation object
#' @param layout character. Name of the layout to see properties for.
#' @export
plot_layout <- function(x, layout) {
  pptx_obj <- officer::read_pptx(template_path(x))
  officer::plot_layout_properties(pptx_obj, layout)
}


#' Get layouts
#'
#' A thin wrapper around `officer::layout_summary()` to get layouts for
#' `R2PpptxPresentation` objects.
#' @param x R2PpptxPresentation object
#' @export
get_layouts <- function(x) {
  pptx_obj <- officer::read_pptx(template_path(x))
  officer::layout_summary(pptx_obj)
}


#' Gets layout properties
#'
#' A thin wrapper around `officer::plot_layout_properties()` to get layouts
#' properties for `R2PpptxPresentation` objects.
#' @param x R2PpptxPresentation object
#' @param layout character. Name of the layout to see properties for.
#' @export
get_layout_properties <- function(x, layout) {
  pptx_obj <- officer::read_pptx(template_path(x))
  officer::layout_properties(pptx_obj, layout)
}
