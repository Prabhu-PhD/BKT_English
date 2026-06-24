# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''

import bkt


clipboard_group = bkt.ribbon.Group(
    id="bkt_clipboard_group",
    label='Storage',
    image_mso='ObjectsMultiSelect',
    children=[
        bkt.ribbon.SplitButtonFixed(
            show_label=False,
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Paste"), context=True),
            children=[
                bkt.mso.button.PasteSpecialDialog,
                bkt.ribbon.Menu(
                    label="Insert menu",
                    supertip="Menu with various insert operations",
                    children=[
                        bkt.mso.button.PasteSpecialDialog,
                        bkt.mso.button.PasteAsPicture,
                        bkt.ribbon.MenuSeparator(title="Paste special"),
                        bkt.ribbon.Button(
                            id='paste_to_slides',
                            label="Insert on selected slides",
                            supertip="Insert the clipboard on all selected slides simultaneously.",
                            image_mso='PasteDuplicate',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "paste_to_slides", slides=True),
                        ),
                        bkt.ribbon.Button(
                            id='paste_as_link',
                            label="Insert as link",
                            supertip="Insert the clipboard as a linked element (e.g. image, OLE object).",
                            image_mso='PasteLink',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "paste_as_link", slide=True),
                        ),
                        bkt.ribbon.Button(
                            id='paste_and_replace',
                            label="Replace with clipboard",
                            supertip="Replace selected shapes with the clipboard content while keeping size and position.",
                            image_mso='PasteSingleCellExcelTableDestinationFormatting',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "paste_and_replace_shapes", slide=True, shapes=True),
                            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
                        ),
                        bkt.ribbon.Button(
                            id='paste_and_distribute',
                            label="Distribute text over selection",
                            supertip="Distribute each paragraph (or cell) from the clipboard individually onto the selected shapes (from left to right, and from top to bottom). Superfluous paragraphs are discarded.",
                            image_mso='PasteMergeList',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "paste_and_distribute", slide=True, shapes=True),
                            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.mso.button.ShowClipboard,
                    ]
                )
            ]
        ),
        bkt.ribbon.SplitButtonFixed(
            show_label=False,
            get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Copy"), context=True),
            children=[
                bkt.mso.button.Copy,
                bkt.ribbon.Menu(
                    label="Copy menu",
                    supertip="Menu with various copy operations",
                    children=[
                        bkt.mso.button.Copy,
                        bkt.mso.button.PasteDuplicate,
                        bkt.ribbon.Button(
                            id="copy_texts",
                            label="Copy shape text",
                            supertip="Copies the text of all selected shapes to the clipboard.",
                            image_mso='DrawTextBox',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "copy_texts", shapes=True),
                            get_enabled=bkt.get_enabled_auto
                        ),
                        bkt.ribbon.Button(
                            id="copy_slide_hq",
                            label="Copy slide as HQ image",
                            supertip="Copies the current slide to the clipboard in high quality.",
                            image_mso='CopyPicture',
                            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "SlidesMore", "copy_in_highquality", slide=True),
                            get_enabled=bkt.get_enabled_auto
                        ),
                    ]
                )
            ]
        ),
        #bkt.mso.control.PasteSpecialDialog,
        #bkt.mso.control.Cut,
        #bkt.mso.control.CopySplitButton,
        
        bkt.ribbon.DynamicMenu(
            label='Selection',
            screentip='Selection of shapes',
            supertip='Selection of shapes that resemble the current shape in type/background/border',
            show_label=False,
            image_mso='ObjectsMultiSelect',
            get_content=bkt.CallbackLazy("toolbox.models.shape_selection", "selection_menu"),
        ),
        
        bkt.mso.control.PasteApplyStyle,
        bkt.mso.control.PickUpStyle,
        bkt.ribbon.Button(
            id="select_by_fill",
            image_mso = 'ColorBlue',
            label='Selection of shapes with the same background',
            show_label=False,
            on_action=bkt.CallbackLazy("toolbox.models.shape_selection", "ShapeSelector", "selectByFill", context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Select shape objects with the same background",
            supertip="Select all shapes on the current slide that have the same background (color) as one of the selected shapes",
        ),


        bkt.mso.control.FormatPainter,
        bkt.ribbon.Button(
            id="format_syncer",
            label="Format Syncer",
            supertip="Format all shapes like the first selected shape",
            image_mso="ShapeFillEffectMoreTexturesDialogClassic",
            show_label=False,
            get_enabled=bkt.apps.ppt_shapes_min2_selected,
            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "FormatPainter", "sync_shapes", shapes=True)
        ),
        bkt.ribbon.Button(
            id="select_by_border",
            image_mso = 'ColorWhite',
            label='Selection of shapes with the same border',
            show_label=False,
            on_action=bkt.CallbackLazy("toolbox.models.shape_selection", "ShapeSelector", "selectByLine", context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            screentip="Select shape objects with the same border",
            supertip="Select all shapes on the current slide that have the same border (color, line style) as one of the selected shapes",
        ),

        #dirty hack to show only one of the following two buttons:
        # bkt.ribbon.Box(get_visible=bkt.Callback(FormatPainter.fp_visible, context=True), children=[
        #     bkt.mso.control.FormatPainter
        # ]),
        
    ]
)


