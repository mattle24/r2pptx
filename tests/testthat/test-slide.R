expect_is_r2pptx_slide_list <- function(x) expect_true(is(x, "R2PptxSlideList"))

describe("R2PptxSlideList", {
  it("can be initialized", {
    x <- new("R2PptxSlideList")
    expect_is_r2pptx_slide_list(x)
  })
  it("succeeds when each element is a slide", {
    slide <- new_slide(layout = "test")
    x <- new("R2PptxSlideList", slides = list(slide))
    expect_is_r2pptx_slide_list(x)
  })
  it("fails when an element is not a slide", {
    expect_error(new("R2PptxSlideList", slides = list("x")))
  })
})

describe("asSlideList", {
  it("works when input is a list of slides", {
    slide <- new_slide(layout = "test")
    x <- asSlideList(list(slide, slide))
    expect_is_r2pptx_slide_list(x)
  })
  it("works when input is a slide", {
    slide <- new_slide(layout = "test")
    x <- asSlideList(slide)
    expect_is_r2pptx_slide_list(x)
  })
})

describe("append_slide", {
  it("works for two slides", {
    slide <- new_slide(layout = "test")
    x <- slide + slide
    expect_is_r2pptx_slide_list(x)
  })
  it("works for a slidelist and a slide", {
    slide <- new_slide(layout = "test")
    slidelist <- asSlideList(slide)
    x <- slidelist + slide
    expect_is_r2pptx_slide_list(x)
  })
  it("works for a slide and a slidelist", {
    slide <- new_slide(layout = "test")
    slidelist <- asSlideList(slide)
    x <- slide + slidelist
    expect_is_r2pptx_slide_list(x)
  })
})
