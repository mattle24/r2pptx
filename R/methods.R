#' @include generics.R
#' @include element.R
#' @include location.R
#' @include slide.R
#' @include presentation.R
NULL


setMethod(
  "+",
  signature = signature(e1 = "ANY", e2 = "R2PptxSlide"),
  function(e1, e2) {
    append_slide(e1, e2)
  }
)
