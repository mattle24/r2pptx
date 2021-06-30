#' @include generics.R
#' @include element.R
#' @include location.R
#' @include slide.R
#' @include presentation.R
NULL

#' Add R2Ppptx slide
#'
#' Add an `R2PptxSlide` object to something else compatible.
#' @export
setMethod(
  "+",
  signature = signature(e1 = "R2Pptx", e2 = "R2PptxSlide"),
  function(e1, e2) {
    append_slide(e1, e2)
  }
)
