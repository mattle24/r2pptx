#' @include generics.R
#' @include element.R
#' @include location.R
#' @include slide.R
#' @include presentation.R
NULL

#' Add R2Ppptx slide
#'
#' Add an `R2PptxSlide` object to something else compatible.
#' @param e1 An object that a slide can be added to, such as an
#'   `R2PptxPresentation`, `R2PptxSlide`, or `R2PptxSlideList`.
#' @param e2 `R2PptxSlide` object
#' @export
#' @return If \code{e1} is an object of class \code{R2PptxPresentation} then
#'   returns an object of class \code{R2PptxPresentation}. Otherwise returns an
#'   object of class \code{R2PptxSlideList}
setMethod(
  "+",
  signature = signature(e1 = "R2Pptx", e2 = "R2PptxSlide"),
  function(e1, e2) {
    append_slide(e1, e2)
  }
)
