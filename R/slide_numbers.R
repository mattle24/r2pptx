#' Add dynamic slide numbers to officer presentation object
#'
#' Adds dynamic slide number field codes to an officer rpptx object before
#' it's written to disk. Uses officer's slide APIs to add the field codes
#' directly to the in-memory XML, avoiding unzip/rezip cycle.
#'
#' @param pptx_obj officer rpptx object
#' @param slides list of R2PptxSlide objects
#' @param template_path character. Path to template for style extraction
#' @param start_slide integer. First slide to add numbers to (default 1)
#' @keywords internal
#' @return Modified officer rpptx object
add_dynamic_slide_numbers <- function(
  pptx_obj,
  slides,
  template_path,
  start_slide = 1
) {
  if (!requireNamespace("xml2", quietly = TRUE)) {
    warning("xml2 package required for dynamic slide numbers. Skipping.")
    return(pptx_obj)
  }
  if (!requireNamespace("uuid", quietly = TRUE)) {
    warning("uuid package required for dynamic slide numbers. Skipping.")
    return(pptx_obj)
  }

  style_cache <- list()
  n_slides <- length(slides)

  for (slide_idx in start_slide:n_slides) {
    layout_name <- slides[[slide_idx]]@layout

    # Get style for this layout (cached)
    if (is.null(style_cache[[layout_name]])) {
      style_cache[[layout_name]] <- extract_slidenum_style(
        template_path,
        layout_name
      )
    }
    style <- style_cache[[layout_name]]

    # Get layout properties using officer
    layout_props <- tryCatch({
      officer::layout_properties(pptx_obj, layout = layout_name)
    }, error = function(e) {
      warning(sprintf(
        "Could not get layout properties for layout '%s': %s",
        layout_name, e$message
      ))
      return(NULL)
    })

    if (is.null(layout_props)) next

    # Find slidenum placeholder
    slidenum_props <- layout_props[
      !is.na(layout_props$fld_type) & layout_props$fld_type == "slidenum",
    ]

    if (nrow(slidenum_props) == 0) next

    # Add slidenum field code using officer's slide API
    pptx_obj <- add_slidenum_to_slide(
      pptx_obj,
      slide_idx = slide_idx + 1,  # officer uses 1-based + 1 for actual slides
      style = style,
      slidenum_props = slidenum_props[1, ]
    )
  }

  pptx_obj
}


#' Add slide number field code to a single slide
#' @keywords internal
add_slidenum_to_slide <- function(pptx_obj, slide_idx, style, slidenum_props) {
  # Get the slide object using officer's API
  slide <- pptx_obj$slide$get_slide(slide_idx)

  # Create the slidenum XML node
  shape_id <- slide_idx + 100
  field_uuid <- uuid::UUIDgenerate()

  slidenum_xml_string <- create_slidenum_xml(
    slide_idx = slide_idx - 1,  # Convert back to user-facing numbering
    style = style,
    slidenum_props = slidenum_props,
    shape_id = shape_id,
    field_uuid = field_uuid
  )

  # Parse and add to slide XML
  node <- xml2::as_xml_document(slidenum_xml_string)
  slide_xml <- slide$get()
  sp_tree <- xml2::xml_find_first(slide_xml, ".//p:spTree")

  if (!inherits(sp_tree, "xml_missing")) {
    xml2::xml_add_child(sp_tree, node)
  }

  pptx_obj
}


#' Create slide number XML element
#' @keywords internal
create_slidenum_xml <- function(slide_idx, style, slidenum_props, shape_id, field_uuid) {
  # Convert inches to EMU (English Metric Units)
  x_emu <- as.integer(slidenum_props$offx * 914400)
  y_emu <- as.integer(slidenum_props$offy * 914400)
  cx_emu <- as.integer(slidenum_props$cx * 914400)
  cy_emu <- as.integer(slidenum_props$cy * 914400)

  # Apply defaults
  style <- apply_style_defaults(style)

  # Build anchor attribute only if present
  anchor_attr <- if (!is.null(style$anchor) && !is.na(style$anchor)) {
    sprintf(' anchor="%s" anchorCtr="0"', style$anchor)
  } else {
    ""
  }

  cat(sprintf(
    "  → Slide %d: color=%s, font=%s, size=%d, align=%s, anchor=%s\n",
    slide_idx, style$color, style$font, style$size, style$align,
    ifelse(is.null(style$anchor) || is.na(style$anchor), "NA", style$anchor)
  ))

  sprintf(
    '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="%d" name="Slide Number %d"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="sldNum" idx="12"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="%d" y="%d"/>
      <a:ext cx="%d" cy="%d"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr spcFirstLastPara="1" wrap="square" lIns="%s" tIns="%s" rIns="%s" bIns="%s"%s>
      <a:normAutofit/>
    </a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="%s"/>
      <a:fld id="{%s}" type="slidenum">
        <a:rPr lang="en-US" sz="%d" baseline="0">
          <a:solidFill><a:srgbClr val="%s"/></a:solidFill>
          <a:latin typeface="%s"/>
        </a:rPr>
        <a:t>%d</a:t>
      </a:fld>
      <a:endParaRPr lang="en-US"/>
    </a:p>
  </p:txBody>
</p:sp>',
    shape_id, slide_idx,
    x_emu, y_emu, cx_emu, cy_emu,
    style$lIns, style$tIns, style$rIns, style$bIns, anchor_attr,
    style$align, field_uuid, style$size, style$color, style$font, slide_idx
  )
}


#' Extract slide number styling from PowerPoint template
#'
#' Extracts color, font, size, alignment, and body properties for slide numbers
#' from a PowerPoint template layout.
#'
#' @param template_path character. Path to the PPTX template file
#' @param layout_name character. Name of the layout
#' @keywords internal
#' @return list with styling attributes
extract_slidenum_style <- function(template_path, layout_name) {
  temp_dir <- tempfile()
  dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))

  utils::unzip(template_path, exdir = temp_dir)

  layout_idx <- get_layout_index(template_path, layout_name)
  if (is.null(layout_idx)) return(default_slidenum_style())

  layout_file <- file.path(
    temp_dir, "ppt", "slideLayouts",
    sprintf("slideLayout%d.xml", layout_idx)
  )

  if (!file.exists(layout_file)) return(default_slidenum_style())

  layout_xml <- xml2::read_xml(layout_file)
  slidenum_shape <- xml2::xml_find_first(
    layout_xml,
    ".//p:sp[.//p:ph[@type='sldNum']]"
  )

  if (inherits(slidenum_shape, "xml_missing")) return(default_slidenum_style())

  # Extract styling
  text_style <- extract_text_style(slidenum_shape)
  body_props <- extract_body_properties(slidenum_shape)

  # Check master for missing values
  master_file <- file.path(temp_dir, "ppt", "slideMasters", "slideMaster1.xml")
  if (file.exists(master_file)) {
    text_style <- fill_from_master(text_style, master_file, "text")
    body_props <- fill_from_master(body_props, master_file, "body")
  }

  style <- c(text_style, body_props)

  cat(sprintf(
    "→ Extracted styling for layout '%s': color=%s, font=%s, size=%d, align=%s, anchor=%s\n",
    layout_name, style$color, style$font, style$size, style$align,
    ifelse(is.null(style$anchor) || is.na(style$anchor), "NA", style$anchor)
  ))

  style
}


#' Get layout index from template
#' @keywords internal
get_layout_index <- function(template_path, layout_name) {
  pptx_obj <- officer::read_pptx(template_path)
  layout_summary <- officer::layout_summary(pptx_obj)
  idx <- which(layout_summary$layout == layout_name)
  if (length(idx) == 0) {
    message(sprintf("Layout '%s' not found", layout_name))
    return(NULL)
  }
  idx
}


#' Default slide number styling
#' @keywords internal
default_slidenum_style <- function() {
  list(
    color = "50514F", font = "Calibri", size = 1000, align = "r",
    anchor = NULL, lIns = "91425", tIns = "91425", rIns = "91425", bIns = "91425"
  )
}


#' Extract text styling from slidenum shape
#' @keywords internal
extract_text_style <- function(slidenum_shape) {
  color <- xml2::xml_attr(
    xml2::xml_find_first(slidenum_shape, ".//a:lstStyle//a:lvl1pPr//a:defRPr//a:solidFill//a:srgbClr"),
    "val"
  )
  font <- xml2::xml_attr(
    xml2::xml_find_first(slidenum_shape, ".//a:lstStyle//a:lvl1pPr//a:defRPr//a:latin"),
    "typeface"
  )
  size <- as.integer(xml2::xml_attr(
    xml2::xml_find_first(slidenum_shape, ".//a:lstStyle//a:lvl1pPr//a:defRPr"),
    "sz"
  ))
  align <- xml2::xml_attr(
    xml2::xml_find_first(slidenum_shape, ".//a:lstStyle//a:lvl1pPr"),
    "algn"
  )
  list(color = color, font = font, size = size, align = align)
}


#' Extract body properties from slidenum shape
#' @keywords internal
extract_body_properties <- function(slidenum_shape) {
  body_pr <- xml2::xml_find_first(slidenum_shape, ".//p:txBody//a:bodyPr")
  if (inherits(body_pr, "xml_missing")) {
    return(list(anchor = NULL, lIns = NULL, tIns = NULL, rIns = NULL, bIns = NULL))
  }
  list(
    anchor = xml2::xml_attr(body_pr, "anchor"),
    lIns = xml2::xml_attr(body_pr, "lIns"),
    tIns = xml2::xml_attr(body_pr, "tIns"),
    rIns = xml2::xml_attr(body_pr, "rIns"),
    bIns = xml2::xml_attr(body_pr, "bIns")
  )
}


#' Fill missing style values from master slide
#' @keywords internal
fill_from_master <- function(style_list, master_file, type = c("text", "body")) {
  type <- match.arg(type)
  master_xml <- xml2::read_xml(master_file)
  master_slidenum <- xml2::xml_find_first(master_xml, ".//p:sp[.//p:ph[@type='sldNum']]")

  if (inherits(master_slidenum, "xml_missing")) return(style_list)

  master_style <- if (type == "text") {
    extract_text_style(master_slidenum)
  } else {
    extract_body_properties(master_slidenum)
  }

  for (name in names(style_list)) {
    if (is.null(style_list[[name]]) || is.na(style_list[[name]])) {
      style_list[[name]] <- master_style[[name]]
    }
  }
  style_list
}


#' Apply defaults to style
#' @keywords internal
apply_style_defaults <- function(style) {
  defaults <- default_slidenum_style()
  for (name in names(defaults)) {
    if (is.null(style[[name]]) || is.na(style[[name]])) {
      style[[name]] <- defaults[[name]]
    }
  }
  style
}
