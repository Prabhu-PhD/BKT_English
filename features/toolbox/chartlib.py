# -*- coding: utf-8 -*-
'''
Created on 04.05.2016

@author: rdebeerst
'''

import bkt


chartlib_button = bkt.ribbon.DynamicMenu(
    id='menu-add-chart',
    label="Insert template slide",
    show_label=False,
    screentip="Insert slide from slide library",
    supertip="A template from the stored slide templates can be inserted as a new slide.",
    image_mso="BibliographyGallery",
    # image_mso="SlideMasterInsertLayout",
    #image_mso="CreateFormBlankForm",
    get_content = bkt.CallbackLazy("toolbox.models.chartlib", "charts", "get_root_menu")
)
shapelib_button = bkt.ribbon.DynamicMenu(
    id='menu-add-shape',
    label="Personal Shape Library",
    show_label=False,
    screentip="Insert shape from shape library",
    supertip="A shape from the stored shape templates can be inserted on the current slide.",
    image_mso="ActionInsert",
    #image_mso="ShapesInsertGallery",
    #image_mso="OfficeExtensionsGallery",
    get_content = bkt.CallbackLazy("toolbox.models.chartlib", "shapes", "get_root_menu")
)

# chartlibgroup = bkt.ribbon.Group(
#     label="chartlib",
#     children=[ chartlib_button, shapelib_button]
# )

# bkt.powerpoint.add_tab(
#     bkt.ribbon.Tab(
#         label="chartlib",
#         children = [
#             chartlibgroup
#         ]
#     )
# )





