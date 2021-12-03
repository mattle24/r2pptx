# r2pptx 0.1.0.9001

* Allow custom locations by using `new_location()` when calling `new_element()`. The `R2PptxLocation` class is now exported. The `R2PptxLocation` class has slots `ph_location_fn` (function) and `args` (list) instead of the location `character` slot   (#8).


# r2pptx 0.1.0

* `R2PptxSlideList` improvements (#3)

  * Added `+` method to add `R2PptxSlideList` objects to `R2PptxSlideList` objects
  
  * Added `new_slidelist()` to create new slide list objects

