describe("validity works", {
  it("succeeeds when no slides are given", {
    p <- new_presentation(.DEFAULT_PPT_TEMPLATE)
    expect_true(is(p, "R2PptxPresentation"))
  })
  it("succeeeds when a valid slide is given", {
    slide <- new_slide(layout = "test")
    p <- new_presentation(.DEFAULT_PPT_TEMPLATE, slides = slide)
    expect_true(is(p, "R2PptxPresentation"))
  })
  it("fails when an invalid slide is given", {
    slide <- "fake slide"
    expect_error(
      new_presentation(.DEFAULT_PPT_TEMPLATE, slides = slide),
      regexp = "Each slide must be a `R2PptxSlide` object"
    )
  })
  it("fails when an non-pptx template path is given", {
    expect_error(new_presentation(template_path = "fake_slides.ppt"),
                 regexp = "Template must be a `.pptx` file")
  })
  it("fails when an invalid template path is given", {
    expect_error(new_presentation(template_path = "fake_slides.pptx"),
                 regexp = "File `fake_slides.pptx` not found")
  })
})

describe("add slide", {
  it("can add a slide", {
    p <- new_presentation(.DEFAULT_PPT_TEMPLATE)
    p <- p + new_slide("test")
    expect_s4_class(p, "R2PptxPresentation")
    expect_equal(length(p), 1)
  })
  it("can add a slideList", {
    p <- new_presentation(.DEFAULT_PPT_TEMPLATE)
    slide_list <- new_slide("test") + new_slide("test")
    expect_s4_class(slide_list, "R2PptxSlideList")
    p <- p + slide_list
    expect_s4_class(p, "R2PptxPresentation")
    expect_equal(length(p), 2)
  })
})

describe("slide numbers", {
    it("adds slide numbers when template has them", {
      skip_if_not(interactive())
      presentation <- new_presentation(.DEFAULT_PPT_TEMPLATE) +
        new_slide("Title Slide", elements = list()) +
        new_slide("Title Slide", elements = list())
      path <- tempfile(fileext = ".pptx")
      write_pptx(presentation, path)
      # developer should visually check slide numbers are present.
      # NOTE: could write something to parse the XML and see if there are slide numbers.
      # Would be better but a lot of work.
      browseURL(path)
    })
    it("handles templates without slide numbers", {
      presentation <- new_presentation(test_path("data/template_no_slide_numbers.pptx"))
      presentation <- presentation +
        new_slide("TITLE", list())
      path <- tempfile(fileext = ".pptx")
      # for this test is it sufficient that we can write the pptx file without error.
      # ie, we do not try to write elements that don't exist
      # there are warnings here because the template is weird.
      # okay to ignore
      suppressWarnings(write_pptx(presentation, path))
    })
  }
)

