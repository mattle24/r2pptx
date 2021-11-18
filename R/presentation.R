#' @include slide.R
#' @include generics.R
NULL

#' Presentation
#'
#' An S4 class to represent a PowerPoint presentation.
#' @slot slides list. List of `R2PptxSlide` objects.
#' @slot template_path character. File path to the PowerPoint template used
#' for this presentation.
#' @export
setClass(
  "R2PptxPresentation",
  contains = "R2Pptx",
  slots = c(
    slides = "list",
    template_path = "character"
  )
)


setValidity("R2PptxPresentation", function(object) {
  if (!all(sapply(object@slides, function(x) inherits(x, "R2PptxSlide")))) {
    "Each slide must be a `R2PptxSlide` object"
  } else if (!endsWith(template_path(object), ".pptx")) {
    "Template must be a `.pptx` file"
  } else if (!file.exists(template_path(object))) {
    glue::glue(
      "Template path must be a valid file. File `{f}` not found",
      f = template_path(object)
    )
  } else {
    TRUE
  }
})


#' New Presentation
#'
#' Make a new `R2PptxPresentation`. Presentations represent PowerPoint decks.
#' @param template_path character. Path of the file that has the PowerPoint
#'   template to use. Defaults to path set in `options("default_pptx_template")`
#' @param slides list. Optional. List of slides to initiate the presentation
#'   with.
#' @export
#' @return An object of class \code{R2PptxPresentation} representing a future
#'   PowerPoint presentation.
new_presentation <- function(
  template_path = getOption("default_pptx_template"),
  slides = list()
) {
  if (class(slides) != "list") {
    slides <- list(slides)
  }
  new(
    "R2PptxPresentation",
    template_path = template_path,
    slides = slides
  )
}


# show method
setMethod(
  "show",
  "R2PptxPresentation",
  function(object) {
    cat("Presentation with", length(object), "slides.")
  }
)


# add slide -------------------------------------------------------------

setMethod(
  "append_slide",
  signature(e1 = "R2PptxPresentation", e2 = "R2PptxSlide"),
  function(e1, e2) {
    e1@slides <- append(e1@slides, e2)
    validObject(e1)
    e1
  }
)


#' Add R2Ppptx slidelist
#'
#' Add an \code{R2PptxSlideList} object to a presentation.
#' @param e1 \code{R2PptxPresentation} object
#' @param e2 \code{R2PptxSlideList} object
#' @keywords internal
#' @return An object of class \code{R2PptxPresentation}, which is the
#'   \code{R2PptxPresentation} object \code{e1} with an additional slide, the
#'   \code{R2pptxSlide} object \code{e2}.
setMethod(
  "+",
  signature = signature(e1 = "R2PptxPresentation", e2 = "R2PptxSlideList"),
  function(e1, e2) {
    for (slide in get_slides(e2)) {
      e1 <- append_slide(e1, slide)
    }
    e1
  }
)



# write pptx --------------------------------------------------------------

#' @describeIn write_pptx Write a presentation to a `.pptx` file
#' @return Returns the \code{R2PptxPresentation} object given to the function.
setMethod(
  "write_pptx",
  "R2PptxPresentation",
  function(x, path) {
    pptx_obj <- officer::read_pptx(path = template_path(x))

    # TODO method to get slides
    for (slide in x@slides) {
      # TODO method to get layout
      pptx_obj <- officer::add_slide(pptx_obj,
                                     layout = slide@layout,
                                     master = pptx_obj$masterLayouts$names()[1])
      for (element in slide@elements) {
        pptx_obj <- add_pptx(element, pptx_obj)
      }
    }
    print(pptx_obj, target = path)
    invisible(x)
  }
)


# length ------------------------------------------------------------------

#' get presentation length (slides)
#' @rdname length
#' @return Integer, the number of slides in the presentation.
setMethod("length", "R2PptxPresentation", function(x) length(x@slides))

# template path -----------------------------------------------------------

#' Get template path
#' @param x object to get the template path for.
#' @export
setGeneric("template_path", function(x) standardGeneric("template_path"))


#' @describeIn template_path Get the template path of an
#'   \code{R2PptxPresentation} object.
#' @return Character, the file path this \code{R2PptxPresentation} points to.
setMethod("template_path", "R2PptxPresentation", function(x) {
  x@template_path
})


#' Set template path
#' @param x object to set the template path of.
#' @param value character. File path of the new template
#' @export
setGeneric("template_path<-", function(x, value) standardGeneric("template_path<-"))

#' @describeIn template_path-set Set the template path of an \code{R2PptxPresentation}
#'   object.
#' @return The \code{R2PptxPresentation} object \code{x} with the changed
#'   template path.
setMethod("template_path<-", "R2PptxPresentation", function(x, value) {
  x@template_path <- value
  validObject(x)
  x
})
