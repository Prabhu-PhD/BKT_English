# -*- coding: utf-8 -*-
'''
Created on 2017-07-24
@author: Florian Stallmann
'''

import bkt
import bkt.library.powerpoint as pplib

import bkt.dotnet as dotnet
Forms = dotnet.import_forms() #required to read clipboard and open file dialogs


MODEL_MODULE = __package__ + ".thumbnails_model"
MODEL_CONTAINER = "Thumbnailer"


BKT_THUMBNAIL = "BKT_THUMBNAIL"


class ThumbnailerUi(object):

    @classmethod
    def has_clipboard_data(cls):
        return Forms.Clipboard.ContainsData(BKT_THUMBNAIL) or (Forms.Clipboard.ContainsData("PowerPoint 12.0 Internal Slides") and Forms.Clipboard.ContainsData("Link Source")) #"PowerPoint 14.0 Slides Package"
        # return Forms.Clipboard.ContainsData(BKT_THUMBNAIL)

    @classmethod
    def enabled_paste(cls):
        return cls.has_clipboard_data()
        #return Forms.Clipboard.ContainsImage()

    @classmethod
    def is_thumbnail(cls, shape):
        return pplib.TagHelper.has_tag(shape, BKT_THUMBNAIL)



thumbnail_gruppe = bkt.ribbon.Group(
    id="bkt_slidethumbnails_group",
    label='Slide thumbnails',
    supertip="Enables inserting updatable slide thumbnails. The `ppt_thumbnails` feature must be installed.",
    image_mso='PasteAsPicture',
    children = [
        bkt.ribbon.Button(
            id = 'slide_copy',
            label="Copy slide(s) as thumbnail",
            show_label=True,
            image_mso='Copy',
            supertip="Copy current slide to create an updatable slide thumbnail.",
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slides_copy", presentation=True, slides=True),
        ),
        # bkt.ribbon.Button(
        #     id = 'shape_copy',
        #     label="Copy shape as thumbnail",
        #     show_label=True,
        #     image_mso='Copy',
        #     supertip="Copy current slide to create an updatable slide thumbnail.",
        #     on_action=bkt.Callback(Thumbnailer.shape_copy, presentation=True, slide=True, shape=True),
        # ),
        bkt.ribbon.SplitButton(
            get_enabled = bkt.Callback(ThumbnailerUi.enabled_paste),
            children=[
                bkt.ribbon.Button(
                    id = 'slide_paste',
                    label="Insert slide thumbnail",
                    show_label=True,
                    image_mso='PasteAsPicture',
                    supertip="Insert the copied slide as an updatable thumbnail with a reference to the original slide.\n\nIf the original slide is from another file in the same directory, only the file name is stored, otherwise the absolute path is stored and the original file must not be moved.",
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste", application=True),
                    # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                ),
                bkt.ribbon.Menu(label="Insert menu", supertip="Insert options for updatable slide thumbnails", children=[
                    bkt.ribbon.Button(
                        id = 'slide_paste_png',
                        label="Insert slide thumbnail as PNG",
                        show_label=True,
                        #image_mso='PasteAsPicture',
                        supertip="Insert the copied slide as an updatable thumbnail in PNG format with a reference to the original slide.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_png", application=True),
                        # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                    ),
                    bkt.ribbon.Button(
                        id = 'slide_paste_btm',
                        label="Insert slide thumbnail as bitmap",
                        show_label=True,
                        image_mso='PasteAsPicture',
                        supertip="Insert the copied slide as an updatable thumbnail in bitmap format with a reference to the original slide.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_btm", application=True),
                        # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                    ),
                    bkt.ribbon.Button(
                        id = 'slide_paste__emf',
                        label="Insert slide thumbnail as vector (EMF)",
                        show_label=True,
                        #image_mso='PasteAsPicture',
                        supertip="Insert the copied slide as an updatable thumbnail in vector format (Enhanced Metafile) with a reference to the original slide.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_emf", application=True),
                        # get_enabled = bkt.Callback(Thumbnailer.enabled_paste),
                    ),
                ])
            ]
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.Button(
                    id = 'slide_refresh',
                    label="Update all thumbnails",
                    show_label=True,
                    image_mso='PictureChange',
                    supertip="Update all slide thumbnails on the selected slides. The thumbnail must have been inserted beforehand with this function. If the slide comes from another file, it is automatically opened briefly.",
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_refresh", application=True, slides=True),
                ),
                bkt.ribbon.Menu(label="Update menu", supertip="Update of slide thumbnails on this slide or in the entire presentation", item_size="large", children=[
                    bkt.ribbon.Button(
                        id = 'slide_refresh2',
                        label="Update thumbnails on slide(s)",
                        description="Update all thumbnails on the current or selected slide(s)",
                        # show_label=True,
                        image_mso='PictureChange',
                        supertip="Update all slide thumbnails on the selected slides. The thumbnail must have been inserted beforehand with this function. If the slide comes from another file, it is automatically opened briefly.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_refresh", application=True, slides=True),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'presentation_refresh',
                        label="Update thumbnails in presentation",
                        description="Update all thumbnails in the entire presentation",
                        # show_label=True,
                        #image_mso='PictureChange',
                        supertip="Update all slide thumbnails in the presentation. The thumbnail must have been inserted beforehand with this function. If the slide comes from another file, it is automatically opened briefly.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "presentation_refresh", application=True, presentation=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'presentation_unset',
                        label="Convert thumbnails in presentation",
                        description="Delete the slide reference of all thumbnails in the presentation and convert thumbnails into images",
                        # show_label=True,
                        #image_mso='PictureChange',
                        supertip="Convert all slide thumbnails in the presentation into normal images that can no longer be updated.",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "presentation_unset", presentation=True),
                    ),
                ])
            ]
        ),
    ]
)


bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    #id_q="nsBKT:powerpoint_toolbox_extensions",
    #insert_after_q="nsBKT:powerpoint_toolbox_advanced",
    insert_before_mso="TabHome",
    label='Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        thumbnail_gruppe,
    ]
), extend=True)


bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuPicture', children=[
        bkt.ribbon.Button(
            id='context-thumbnail-refresh',
            label="Update thumbnail",
            supertip="Update selected slide thumbnail",
            insertBeforeMso='Cut',
            image_mso='PictureChange',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "shape_refresh", shape=True, application=True),
            get_visible=bkt.Callback(ThumbnailerUi.is_thumbnail, shape=True),
        ),
        bkt.ribbon.DynamicMenu(
            id='context-thumbnail-settings',
            label="Thumbnail settings",
            supertip="Change settings of the selected slide thumbnail",
            image_mso='PictureSharpenSoftenGallery',
            insertBeforeMso='Cut',
            get_visible=bkt.Callback(ThumbnailerUi.is_thumbnail, shape=True),
            get_content=bkt.CallbackLazy(MODEL_MODULE, "context_settings")
        ),
        bkt.ribbon.DynamicMenu(
            id='context-thumbnail-reference',
            label="Slide reference",
            supertip="Open or change the reference of the selected slide thumbnail",
            image_mso='PictureInsertFromFile',
            insertBeforeMso='Cut',
            get_visible=bkt.Callback(ThumbnailerUi.is_thumbnail, shape=True),
            get_content=bkt.CallbackLazy(MODEL_MODULE, "context_reference")
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut')
    ])
)


bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuThumbnail', children=[
        bkt.ribbon.Button(
            id='context-thumbnail-slide-copy',
            label="Copy as slide thumbnail",
            supertip="Copy selected slide as an updatable thumbnail",
            insertAfterMso='Copy',
            image_mso='Copy',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slides_copy", presentation=True, slides=True),
            #get_visible=bkt.Callback(Thumbnailer.is_thumbnail, shape=True),
        ),
    ])
)

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuFrame', children=[
        bkt.ribbon.SplitButton(
            insertAfterMso='PasteGalleryMini',
            get_enabled=bkt.Callback(ThumbnailerUi.enabled_paste),
            children=[
                bkt.ribbon.Button(
                    id='context-thumbnail-slide-paste',
                    label="Insert as slide thumbnail",
                    supertip="Insert as an updatable slide thumbnail in PNG format",
                    image_mso='PasteAsPicture',
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste", application=True),
                    #get_visible=bkt.Callback(Thumbnailer.is_thumbnail, shape=True),
                    get_enabled=bkt.Callback(ThumbnailerUi.enabled_paste),
                ),
                bkt.ribbon.Menu(label="Insert as slide thumbnail menu", supertip="Select format for inserting the thumbnail", children=[
                    bkt.ribbon.Button(
                        label="Insert as PNG (default)",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_png", application=True),
                    ),
                    bkt.ribbon.Button(
                        label="Insert as bitmap",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_btm", application=True),
                    ),
                    bkt.ribbon.Button(
                        label="Insert as vector (EMF)",
                        on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "slide_paste_emf", application=True),
                    ),
                ])
            ]
        ),
    ])
)


# register dialog
bkt.powerpoint.context_dialogs.register_dialog(
    bkt.contextdialogs.ContextDialog(
        id=BKT_THUMBNAIL,
        module="ppt_thumbnails.thumbnails_popup"
    )
)