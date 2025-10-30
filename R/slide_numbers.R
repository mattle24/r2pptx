extract_slidenum_style_from_template <- function(template_path, layout_name) {
  temp_dir <- tempfile()
  dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))

  utils::unzip(template_path, exdir = temp_dir)

  template_pres <- officer::read_pptx(template_path)
  layout_summary <- officer::layout_summary(template_pres)
  layout_idx <- which(layout_summary$layout == layout_name)

  if (length(layout_idx) == 0) {
    message(sprintf("Layout '%s' not found", layout_name))
    return(list(color = "50514F", font = "Calibri", size = 1000))
  }

  layout_file <- file.path(temp_dir, "ppt", "slideLayouts",
                           sprintf("slideLayout%d.xml", layout_idx))

  if (!file.exists(layout_file)) {
    message(sprintf("Layout file not found: %s", layout_file))
    return(list(color = "50514F", font = "Calibri", size = 1000))
  }

  layout_xml <- xml2::read_xml(layout_file)
  slidenum_shape <- xml2::xml_find_first(
    layout_xml,
    ".//p:sp[.//p:ph[@type='sldNum']]"
  )

  if (inherits(slidenum_shape, "xml_missing")) {
    message(sprintf("No slidenum placeholder found in layout '%s'", layout_name))
    return(list(color = "50514F", font = "Calibri", size = 1000))
  }

  color_node <- xml2::xml_find_first(
    slidenum_shape,
    ".//a:lstStyle//a:lvl1pPr//a:defRPr//a:solidFill//a:srgbClr"
  )
  color <- if (!inherits(color_node, "xml_missing")) {
    xml2::xml_attr(color_node, "val")
  } else {
    NULL
  }

  font_node <- xml2::xml_find_first(
    slidenum_shape,
    ".//a:lstStyle//a:lvl1pPr//a:defRPr//a:latin"
  )
  font <- if (!inherits(font_node, "xml_missing")) {
    xml2::xml_attr(font_node, "typeface")
  } else {
    NULL
  }

  size_node <- xml2::xml_find_first(
    slidenum_shape,
    ".//a:lstStyle//a:lvl1pPr//a:defRPr"
  )
  size <- if (!inherits(size_node, "xml_missing")) {
    as.integer(xml2::xml_attr(size_node, "sz"))
  } else {
    NULL
  }

  if (is.null(color) || is.null(font) || is.null(size)) {
    message(sprintf("Layout '%s' missing styling, checking master slide", layout_name))

    master_files <- list.files(
      file.path(temp_dir, "ppt", "slideMasters"),
      pattern = "^slideMaster[0-9]+\\.xml$",
      full.names = TRUE
    )

    if (length(master_files) > 0) {
      master_xml <- xml2::read_xml(master_files[1])
      master_slidenum <- xml2::xml_find_first(
        master_xml,
        ".//p:sp[.//p:ph[@type='sldNum']]"
      )

      if (!inherits(master_slidenum, "xml_missing")) {
        if (is.null(color)) {
          color_node <- xml2::xml_find_first(
            master_slidenum,
            ".//a:lstStyle//a:lvl1pPr//a:defRPr//a:solidFill//a:srgbClr"
          )
          if (!inherits(color_node, "xml_missing")) {
            color <- xml2::xml_attr(color_node, "val")
          }
        }

        if (is.null(font)) {
          font_node <- xml2::xml_find_first(
            master_slidenum,
            ".//a:lstStyle//a:lvl1pPr//a:defRPr//a:latin"
          )
          if (!inherits(font_node, "xml_missing")) {
            font <- xml2::xml_attr(font_node, "typeface")
          }
        }

        if (is.null(size)) {
          size_node <- xml2::xml_find_first(
            master_slidenum,
            ".//a:lstStyle//a:lvl1pPr//a:defRPr"
          )
          if (!inherits(size_node, "xml_missing")) {
            size <- as.integer(xml2::xml_attr(size_node, "sz"))
          }
        }
      }
    }
  }

  align_node <- xml2::xml_find_first(
    slidenum_shape,
    ".//a:lstStyle//a:lvl1pPr"
  )
  align <- if (!inherits(align_node, "xml_missing")) {
    xml2::xml_attr(align_node, "algn")
  } else {
    NULL
  }

  if (is.null(align) && !is.null(master_files) && length(master_files) > 0) {
    master_xml <- xml2::read_xml(master_files[1])
    master_slidenum <- xml2::xml_find_first(
      master_xml,
      ".//p:sp[.//p:ph[@type='sldNum']]"
    )
    if (!inherits(master_slidenum, "xml_missing")) {
      align_node <- xml2::xml_find_first(
        master_slidenum,
        ".//a:lstStyle//a:lvl1pPr"
      )
      if (!inherits(align_node, "xml_missing")) {
        align <- xml2::xml_attr(align_node, "algn")
      }
    }
  }

  body_pr_node <- xml2::xml_find_first(
    slidenum_shape,
    ".//p:txBody//a:bodyPr"
  )

  lIns <- tIns <- rIns <- bIns <- anchor <- NULL

  if (!inherits(body_pr_node, "xml_missing")) {
    lIns <- xml2::xml_attr(body_pr_node, "lIns")
    tIns <- xml2::xml_attr(body_pr_node, "tIns")
    rIns <- xml2::xml_attr(body_pr_node, "rIns")
    bIns <- xml2::xml_attr(body_pr_node, "bIns")
    anchor <- xml2::xml_attr(body_pr_node, "anchor")
  }

  if ((is.null(anchor) || is.na(anchor)) && !is.null(master_files) && length(master_files) > 0) {
    master_xml <- xml2::read_xml(master_files[1])
    master_slidenum <- xml2::xml_find_first(
      master_xml,
      ".//p:sp[.//p:ph[@type='sldNum']]"
    )
    if (!inherits(master_slidenum, "xml_missing")) {
      master_body_pr <- xml2::xml_find_first(
        master_slidenum,
        ".//p:txBody//a:bodyPr"
      )
      if (!inherits(master_body_pr, "xml_missing")) {
        if (is.null(lIns) || is.na(lIns)) lIns <- xml2::xml_attr(master_body_pr, "lIns")
        if (is.null(tIns) || is.na(tIns)) tIns <- xml2::xml_attr(master_body_pr, "tIns")
        if (is.null(rIns) || is.na(rIns)) rIns <- xml2::xml_attr(master_body_pr, "rIns")
        if (is.null(bIns) || is.na(bIns)) bIns <- xml2::xml_attr(master_body_pr, "bIns")
        if (is.null(anchor) || is.na(anchor)) anchor <- xml2::xml_attr(master_body_pr, "anchor")
      }
    }
  }

  color <- if (is.null(color)) "50514F" else color
  font <- if (is.null(font)) "Calibri" else font
  size <- if (is.null(size)) 1000 else size
  align <- if (is.null(align)) "r" else align

  lIns <- if (is.null(lIns) || is.na(lIns)) "91425" else lIns
  tIns <- if (is.null(tIns) || is.na(tIns)) "91425" else tIns
  rIns <- if (is.null(rIns) || is.na(rIns)) "91425" else rIns
  bIns <- if (is.null(bIns) || is.na(bIns)) "91425" else bIns

  result <- list(
    color = color,
    font = font,
    size = size,
    align = align,
    anchor = anchor,
    lIns = lIns,
    tIns = tIns,
    rIns = rIns,
    bIns = bIns
  )

  cat(sprintf(
    "→ Extracted styling for layout '%s': color=%s, font=%s, size=%d, align=%s, anchor=%s\n",
    layout_name, color, font, size, align, ifelse(is.null(anchor) || is.na(anchor), "NA", anchor)
  ))

  result
}

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
  if (!requireNamespace("xml2", quietly = TRUE)) {
    warning("xml2 package required for dynamic slide numbers. Skipping.")
    return(pptx_path)
  }
  if (!requireNamespace("uuid", quietly = TRUE)) {
    warning("uuid package required for dynamic slide numbers. Skipping.")
    return(pptx_path)
  }

  temp_dir <- tempfile()
  dir.create(temp_dir)
  on.exit(unlink(temp_dir, recursive = TRUE))

  utils::unzip(pptx_path, exdir = temp_dir)
  template_pres <- officer::read_pptx(template_path)

  style_cache <- list()

  n_slides <- length(slides)
  for (slide_idx in start_slide:n_slides) {
    slide_file <- file.path(temp_dir, "ppt", "slides",
                           sprintf("slide%d.xml", slide_idx + 1))

    if (!file.exists(slide_file)) {
      next
    }

    layout_name <- slides[[slide_idx]]@layout

    if (is.null(style_cache[[layout_name]])) {
      style_cache[[layout_name]] <- extract_slidenum_style_from_template(
        template_path,
        layout_name
      )
    }
    style <- style_cache[[layout_name]]

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

    slidenum_rows <- layout_props[
      !is.na(layout_props$fld_type) & layout_props$fld_type == "slidenum",
    ]

    if (nrow(slidenum_rows) == 0) {
      next
    }

    slidenum_pos <- slidenum_rows[1, ]

    slide_xml <- xml2::read_xml(slide_file)
    sp_tree <- xml2::xml_find_first(slide_xml, ".//p:spTree")

    if (inherits(sp_tree, "xml_missing")) {
      next
    }

    shape_id <- slide_idx + 100
    field_uuid <- uuid::UUIDgenerate()

    x_emu <- as.integer(slidenum_pos$offx * 914400)
    y_emu <- as.integer(slidenum_pos$offy * 914400)
    cx_emu <- as.integer(slidenum_pos$cx * 914400)
    cy_emu <- as.integer(slidenum_pos$cy * 914400)

    cat(sprintf(
      "  → Slide %d: Applying style color=%s, font=%s, size=%d, align=%s, anchor=%s\n",
      slide_idx, style$color, style$font, style$size, style$align,
      ifelse(is.null(style$anchor) || is.na(style$anchor), "NA", style$anchor)
    ))

    anchor_attr <- if (!is.null(style$anchor) && !is.na(style$anchor)) {
      sprintf(' anchor="%s" anchorCtr="0"', style$anchor)
    } else {
      ""
    }

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
    <a:bodyPr spcFirstLastPara="1" wrap="square" lIns="%s" tIns="%s" rIns="%s" bIns="%s"%s>
      <a:normAutofit/>
    </a:bodyPr>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="%s"/>
      <a:fld id="{%s}" type="slidenum">
        <a:rPr lang="en-US" sz="%d" baseline="0">
          <a:solidFill>
            <a:srgbClr val="%s"/>
          </a:solidFill>
          <a:latin typeface="%s"/>
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
      style$lIns,
      style$tIns,
      style$rIns,
      style$bIns,
      anchor_attr,
      style$align,
      field_uuid,
      style$size,
      style$color,
      style$font,
      slide_idx
    )

    slidenum_node <- xml2::read_xml(slidenum_xml_string)
    xml2::xml_add_child(sp_tree, slidenum_node)
    xml2::write_xml(slide_xml, slide_file)
  }

  output_file <- tempfile(fileext = ".pptx")
  old_wd <- getwd()
  setwd(temp_dir)
  system(sprintf("zip -r -q '%s' .", output_file))
  setwd(old_wd)

  file.copy(output_file, pptx_path, overwrite = TRUE)

  return(pptx_path)
}
