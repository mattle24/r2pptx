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
