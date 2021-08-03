describe("validity works", {
  it("succeeeds when no slides are given", {
    p <- new_presentation()
    expect_true(is(p, "R2PptxPresentation"))
  })
  it("succeeeds when a valid slide is given", {
    slide <- new_slide(layout = "test")
    p <- new_presentation(slides = slide)
    expect_true(is(p, "R2PptxPresentation"))
  })
  it("fails when an invalid slide is given", {
    slide <- "fake slide"
    expect_error(
      new_presentation(slides = slide),
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
    p <- new_presentation()
    p <- p + new_slide("test")
    expect_s4_class(p, "R2PptxPresentation")
    expect_equal(length(p), 1)
  })
  it("can add a slideList", {
    p <- new_presentation()
    slide_list <- new_slide("test") + new_slide("test")
    expect_s4_class(slide_list, "R2PptxSlideList")
    p <- p + slide_list
    expect_s4_class(p, "R2PptxPresentation")
    expect_equal(length(p), 2)
  })
})
