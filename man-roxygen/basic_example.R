# see https://mattle24.github.io/r2pptx/articles/basic_usage.html
# for more documentation
presentation <- new_presentation()
print(presentation)
print(template_path(presentation))

title_slide <- new_slide(
  layout = "Title Slide",
  elements = list(
    new_element(key = "Title 1", value = "The title"),
    new_element(key = "Subtitle 2", value = "The subtitle")
  )
)
print(title_slide)
presentation <- presentation + title_slide
print(presentation)
ppt_path <- tempfile(fileext = ".pptx")
write_pptx(presentation, ppt_path)
if (interactive()) system(paste("open", ppt_path))
