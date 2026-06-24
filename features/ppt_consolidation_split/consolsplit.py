# -*- coding: utf-8 -*-
'''
Created on 2017-11-09
@author: Florian Stallmann
'''

import bkt

MODEL_MODULE = __package__ + ".consolsplit_model"
MODEL_CONTAINER = "ConsolSplit"

# consolsplit_gruppe = bkt.ribbon.Group(
#     label='Consolidate & split',
#     image_mso='ThemeBrowseForThemes',
#     children = [
#         bkt.ribbon.Button(
#             id = 'consolidate_ppt_slides',
#             label="Append slides from files",
#             show_label=True,
#             size="large",
#             image_mso='ThemeBrowseForThemes',
#             supertip="Append all slides from multiple PowerPoint files to this presentation.",
#             on_action=bkt.Callback(ConsolSplit.consolidate_ppt_slides, application=True, presentation=True),
#         ),
#         bkt.ribbon.Menu(
#             id = 'split_slides_to_ppt',
#             image_mso='ThemeSaveCurrent',
#             label="Save slides individually",
#             supertip="Save slides into individual PowerPoint files.",
#             show_label=True,
#             size="large",
#             item_size="large",
#             children=[
#                 bkt.ribbon.Button(
#                     # id = 'split_slides_to_ppt',
#                     label="Save slides individually",
#                     image_mso='ThemeSaveCurrent',
#                     description="Save each slide as an individual file in the selected folder. The files are numbered with the slide number and named after the slide title.",
#                     on_action=bkt.Callback(ConsolSplit.split_slides_to_ppt, application=True, presentation=True, slides=True),
#                 ),
#                 bkt.ribbon.Button(
#                     # id = 'split_slides_to_ppt',
#                     label="Save sections individually",
#                     image_mso='ThemeSaveCurrent',
#                     description="Save each section as an individual file in the selected folder. The files are numbered and named after the section title.",
#                     on_action=bkt.Callback(ConsolSplit.split_sections_to_ppt, application=True, presentation=True, slides=True),
#                 ),
#             ]
#         )
#     ]
# )

# bkt.powerpoint.add_tab(bkt.ribbon.Tab(
#     id="bkt_powerpoint_toolbox_extensions",
#     insert_before_mso="TabHome",
#     label=u'Toolbox 3/3',
#     # get_visible defaults to False during async-startup
#     get_visible=bkt.Callback(lambda:True),
#     children = [
#         consolsplit_gruppe,
#     ]
# ), extend=True)



bkt.powerpoint.add_backstage_control(
    bkt.ribbon.Tab(
        label="Consol./split",
        title="BKT - Consolidate & split files",
        insertAfterMso="TabPublish", #http://youpresent.co.uk/customising-powerpoint-2016-backstage-view/
        columnWidthPercent="50",
        children=[
            bkt.ribbon.FirstColumn(children=[
                bkt.ribbon.Group(id="bkt_consolsplit_consolidate_group", label="Append slides from files", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Button(
                            label="Append slides from files",
                            supertip="Append all slides from multiple PowerPoint files to this presentation",
                            image_mso='ThemeBrowseForThemes',
                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "consolidate_ppt_slides", application=True, presentation=True),
                            is_definitive=True,
                        ),
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Label(label="Append all slides from multiple PowerPoint files to this presentation."),
                        bkt.ribbon.Label(label="This operation can take some time for large files and many slides!"),
                    ]),
                ]),
                bkt.ribbon.Group(id="bkt_consolsplit_pic2slides_group", label="Create slides from images", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Menu(
                            label="Create slides from images",
                            supertip="Insert multiple image files (jpg, png, emf) each onto its own slide",
                            image_mso='PhotoGalleryProperties',
                            children=[
                                bkt.ribbon.MenuGroup(
                                    item_size="large",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Select image files",
                                            image_mso='PhotoGalleryProperties',
                                            description="Select all image files to insert individually.",
                                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "pictures_to_slides", context=True),
                                            is_definitive=True,
                                        ),
                                        bkt.ribbon.Button(
                                            label="Select folder with images",
                                            image_mso='OpenFolder',
                                            description="Select a folder with image files to insert.",
                                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "folder_to_slides", context=True),
                                            is_definitive=True,
                                        ),
                                    ]
                                ),
                            ]
                        ),
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Label(label="Insert multiple image files (jpg, png, emf) each onto its own slide."),
                        bkt.ribbon.Label(label="The selected slide is duplicated for each image file and the file name is set as the title."),
                    ]),
                ]),
                bkt.ribbon.Group(id="bkt_consolsplit_split_group", label="Save slides individually", children=[
                    bkt.ribbon.PrimaryItem(children=[
                        bkt.ribbon.Menu(
                            label="Save slides individually",
                            supertip="Save all slides as individual PowerPoint files in the selected folder",
                            image_mso='ThemeSaveCurrent',
                            children=[
                                bkt.ribbon.MenuGroup(
                                    item_size="large",
                                    children=[
                                        bkt.ribbon.Button(
                                            label="Save slides individually",
                                            image_mso='ThemeSaveCurrent',
                                            description="Save each slide individually. The files are numbered with the slide number and named after the slide title.",
                                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "split_slides_to_ppt", context=True, slides=True),
                                            is_definitive=True,
                                        ),
                                        bkt.ribbon.Button(
                                            label="Save sections individually",
                                            image_mso='SectionAdd',
                                            description="Save each section individually. The files are numbered and named after the section title.",
                                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "split_sections_to_ppt", context=True, slides=True),
                                            is_definitive=True,
                                        ),
                                    ]
                                ),
                            ]
                        ),
                    ]),
                    bkt.ribbon.TopItems(children=[
                        bkt.ribbon.Label(label="Save all slides as individual PowerPoint files in the selected folder."),
                        bkt.ribbon.Label(label="This operation can take some time for large files and many slides!"),
                    ]),
                ]),
            ])
        ]
    )
)