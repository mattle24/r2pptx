describe("validity works", {
  it("succeeeds when no slides are given", {
    expect_silent(new_presentation("test"))
  })
  it("succeeeds when a valid slide is given", {
    slide <- new_slide(layout = "test")
    expect_silent(new_presentation("test", slides = slide))
  })
  it("fails when an invalid slide is given", {
    slide <- "fake slide"
    expect_error(new_presentation("test", slides = slide))
  })
  it("fails when an invalid template path is given", {
    expect_error(new_presentation("test", template_path = "fake_slides.pptx"))
  })
})
