.onLoad <- function(libname, pkgname) {
  if (is.null(getOption("default_pptx_template"))) {
    options("default_pptx_template" = .DEFAULT_PPT_TEMPLATE)
  }
  invisible()
}
