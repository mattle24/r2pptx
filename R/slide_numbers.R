#' Add dynamic slide numbers to slides
#'
#' Internal function to inject dynamic slide number field codes into PowerPoint
#' slides after they've been created. This ensures slide numbers appear and
#' update automatically when slides are reordered.
#'
#' @param pptx_path character. Path to the PPTX file to modify
#' @param slides list. List of R2PptxSlide objects (to get layout names)
#' @param template_path character. Path to the template (to get placeholder positions)
#' @param start_slide integer. First slide to add numbers to (default 1)
#' @keywords internal
#' @return character. Path to the modified PPTX file
add_dynamic_slide_numbers <- function(
  pptx_path,
  slides,
  template_path,
  start_slide = 1
) {
  # Check if xml2 and uuid packages are available
  if (!requireNamespace("xml2", quietly = TRUE)) {
    warning("xml2 package required for dynamic slide numbers. Skipping.")
    return(pptx_path)
  }
  if (!requireNamespace("uuid", quietly = TRUE)) {
    warning("uuid package required for dynamic slide numbers. Skipping.")
    return(pptx_path)
  }

  # Create temporary directory for unzipping
  temp_dir <- tempfile()
  dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))

  # Unzip the PPTX
  utils::unzip(pptx_path, exdir = temp_dir)

  # Read template to get layout properties
  template_pres <- officer::read_pptx(template_path)

  # Process each slide
  n_slides <- length(slides)
  for (slide_idx in start_slide:n_slides) {
    slide_file <- file.path(temp_dir, "ppt", "slides",
                           sprintf("slide%d.xml", slide_idx + 1))  # +1 because officer adds 1 for master

    if (!file.exists(slide_file)) {
      next
    }

    # Get layout name for this slide
    layout_name <- slides[[slide_idx]]@layout

    # Get slidenum placeholder position from layout
    layout_props <- tryCatch({
      officer::layout_properties(template_pres, layout = layout_name)
    }, error = function(e) {
      warning(sprintf("Could not get layout properties for layout '%s': %s",
                     layout_name, e$message))
      return(NULL)
    })

    if (is.null(layout_props)) {
      next
    }

    # Check if this layout has a slidenum placeholder
    slidenum_rows <- layout_props[
      !is.na(layout_props$fld_type) & layout_props$fld_type == "slidenum",
    ]

    if (nrow(slidenum_rows) == 0) {
      # No slidenum in this layout, skip
      next
    }

    # Get position from layout
    slidenum_pos <- slidenum_rows[1, ]

    # Read and parse slide XML
    slide_xml <- xml2::read_xml(slide_file)
    sp_tree <- xml2::xml_find_first(slide_xml, ".//p:spTree")

    if (inherits(sp_tree, "xml_missing")) {
      next
    }

    # Generate unique IDs
    shape_id <- slide_idx + 100
    field_uuid <- uuid::UUIDgenerate()

    # Convert position from inches to EMU (English Metric Units)
    # 1 inch = 914400 EMU
    x_emu <- as.integer(slidenum_pos$offx * 914400)
    y_emu <- as.integer(slidenum_pos$offy * 914400)
    cx_emu <- as.integer(slidenum_pos$cx * 914400)
    cy_emu <- as.integer(slidenum_pos$cy * 914400)

    # Create the slide number shape XML with proper field code
    slidenum_xml_string <- sprintf('
<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="%d" name="Slide Number %d"/>
    <p:cNvSpPr>
      <a:spLocks noGrp="1"/>
    </p:cNvSpPr>
    <p:nvPr>
      <p:ph type="sldNum" idx="12"/>
    </p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="%d" y="%d"/>
      <a:ext cx="%d" cy="%d"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="91425" tIns="91425" rIns="91425" bIns="91425"/>
    <a:lstStyle/>
    <a:p>
      <a:fld id="{%s}" type="slidenum">
        <a:rPr lang="en-US" sz="1000" baseline="0">
          <a:solidFill>
            <a:srgbClr val="50514F"/>
          </a:solidFill>
          <a:latin typeface="Calibri"/>
        </a:rPr>
        <a:t>%d</a:t>
      </a:fld>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </p:txBody>
</p:sp>',
      shape_id,
      slide_idx,
      x_emu,
      y_emu,
      cx_emu,
      cy_emu,
      field_uuid,
      slide_idx
    )

    # Parse and add to slide
    slidenum_node <- xml2::read_xml(slidenum_xml_string)
    xml2::xml_add_child(sp_tree, slidenum_node)

    # Write back
    xml2::write_xml(slide_xml, slide_file)
  }

  # Rezip the PPTX
  output_file <- tempfile(fileext = ".pptx")
  old_wd <- getwd()
  setwd(temp_dir)
  system(sprintf("zip -r -q '%s' .", output_file))
  setwd(old_wd)

  # Replace original file with modified version
  file.copy(output_file, pptx_path, overwrite = TRUE)

  return(pptx_path)
}
