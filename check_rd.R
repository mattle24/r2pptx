files <- list.files("man", pattern = ".Rd$", full.names = TRUE)
function_files <- files[sapply(files, function(x) !grepl("-class.Rd$", x))]
contains_value <- sapply(function_files, function(f) {
  rd_contents <- readLines(f)
  any(sapply(rd_contents, function(x) grepl("value\\{", x)))
})
if (length(contains_value) > sum(contains_value)) {
  needs_value <- function_files[!contains_value]
  stop(paste(needs_value, collapse = ", "), "\nneed @return tags")
} else {
  n <- sum(contains_value)
  paste0(n, "/", n, " function documentation files have a \value tag")
}
