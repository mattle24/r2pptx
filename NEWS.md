# r2pptx 0.1.0.9002 (development)

* Added dynamic slide numbers feature (#TBD)

  * `write_pptx()` now automatically adds dynamic slide number field codes to slides

  * New parameter `add_slide_numbers` (default TRUE) to enable/disable slide numbers

  * New parameter `start_slide` (default 1) to control which slide to start numbering from

  * Slide numbers use PowerPoint field codes, so they update automatically when slides are reordered

  * Works with Google Slides after upload/conversion

# r2pptx 0.1.0.9001

* Allow custom locations by using `new_location()` when calling `new_element()`. The `R2PptxLocation` class is now exported. The `R2PptxLocation` class has slots `ph_location_fn` (function) and `args` (list) instead of the location `character` slot   (#8).


# r2pptx 0.1.0

* `R2PptxSlideList` improvements (#3)

  * Added `+` method to add `R2PptxSlideList` objects to `R2PptxSlideList` objects
  
  * Added `new_slidelist()` to create new slide list objects

