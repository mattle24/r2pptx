.onLoad <- function(libname, pkgname) {
  if (is.null(getOption("default_pptx_template"))) {
    options("default_pptx_template" = system.file(package = "officer", "template/template.pptx"))
  }
  invisible()
}
