#' @include generics.R
#' @include element.R
#' @include location.R
#' @include slide.R
#' @include presentation.R
NULL


setMethod(
  "+",
  signature = signature(e1 = "R2Pptx", e2 = "R2Pptx"),
  function(e1, e2) {
    append_slide(e1, e2)
  }
)
