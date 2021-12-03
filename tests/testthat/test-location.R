test_that("location works on character input", {
  element <- new_element(
    key = "Title 1",
    value = "Some text"
  )
  presentation <- new_presentation() +
    new_slide("Title Slide", elements = list(element))
  path <- tempfile(fileext = ".pptx")
  write_pptx(presentation, path)
  expect_true(file.exists(path))
})


test_that("location works on R2PptxLocation input", {
  element_location <- new_location(officer::ph_location, left = 2, top = 2)
  element <- new_element(
    key = element_location,
    value = "Some text"
  )
  presentation <- new_presentation() +
    new_slide("Title Slide", elements = list(element))
  path <- tempfile(fileext = ".pptx")
  write_pptx(presentation, path)
  expect_true(file.exists(path))
})
