# -*- coding: utf-8 -*-
'''
Created on 06.02.2018

@author: rdebeerst
'''



import bkt

#FIXME: would be nice to have less dependencies and more lazy loading of modules on callback
# from . import text
# from . import arrange
from .. import harvey
# from .. import shapes as mod_shapes
# from . import shape_selection
from .. import info
from .. import agenda
from .. import linkshapes
# from . import processshapes
# from . import language
# from . import slides


# =========================================
# = LOGIC ONLY REQUIRED FOR CONTEXT MENUS =
# =========================================

# Context menu if multiple connectors are selected
class CtxVerbinder(object):
    @staticmethod
    def ctx_connectors_reroute_enabled(shapes):
        return all(shape.Connector == -1 and shape.ConnectorFormat.BeginConnected == -1 and shape.ConnectorFormat.EndConnected == -1 for shape in shapes)

    @staticmethod
    def ctx_connectors_visible(shapes):
        return all(shape.Connector == -1 for shape in shapes)

    @staticmethod
    def set_connector_type(shapes, con_type):
        for shape in shapes:
            if shape.Connector == -1: #msoTrue
                shape.ConnectorFormat.Type = con_type

    @staticmethod
    def reroute_connectors(shapes):
        for shape in shapes:
            if shape.Connector == -1 and shape.ConnectorFormat.BeginConnected == -1 and shape.ConnectorFormat.EndConnected == -1: #msoTrue
                shape.RerouteConnections()

    @staticmethod
    def invert_direction(shapes):
        for shape in shapes:
            if shape.Connector == -1: #msoTrue
                #Swap begin and end styles
                shape.Line.BeginArrowheadLength, shape.Line.EndArrowheadLength = shape.Line.EndArrowheadLength, shape.Line.BeginArrowheadLength
                shape.Line.BeginArrowheadStyle, shape.Line.EndArrowheadStyle = shape.Line.EndArrowheadStyle, shape.Line.BeginArrowheadStyle
                shape.Line.BeginArrowheadWidth, shape.Line.EndArrowheadWidth = shape.Line.EndArrowheadWidth, shape.Line.BeginArrowheadWidth


class PictureFormat(object):
    @staticmethod
    def make_img_transparent(slide, shapes, transparency=0.5):
        if not bkt.message.confirmation("The existing image will be replaced. Continue?"):
            return

        import tempfile, os
        import bkt.library.powerpoint as pplib
        filename = os.path.join(tempfile.gettempdir(), "bktimgtransp.png")

        for shape in shapes:
            if shape.Type != pplib.MsoShapeType["msoPicture"]:
                continue

            shape.Export(filename, 2) #2=ppShapeFormatPNG

            pic_shape = slide.Shapes.AddShape(
                shape.AutoShapeType,
                shape.Left, shape.Top,
                shape.Width, shape.Height
                )
            pic_shape.LockAspectRatio = -1
            pic_shape.Rotation = shape.Rotation
            pplib.set_shape_zorder(pic_shape, value=shape.ZOrderPosition)
            shape.PickUp()
            pic_shape.Apply()
            pic_shape.line.visible = shape.line.visible # line is not properly transferred by pickup-apply

            pic_shape.fill.UserPicture(filename)
            pic_shape.fill.transparency = transparency
            pic_shape.Select(replace=False)

            shape.Delete()
            os.remove(filename)



# =================
# = CONTEXT MENUS =
# =================

class ContextMenuRecurring(object):
    ''' Collection of functions/buttons that appear in multiple context menus '''

    ### Paste replace ###
    cb_pastereplace_enabled = bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("Paste"), context=True)
    cb_pastereplace_action = bkt.CallbackLazy("toolbox.shape_selection", "SlidesMore","paste_and_replace", slide=True, shape=True)

    @classmethod
    def paste_replace_button(cls, prefix, **kwargs):
        return bkt.ribbon.Button(
            id=prefix+'-paste-and-replace',
            label="Replace with clipboard",
            supertip="Replace selected shape with the clipboard content while keeping size and position.",
            insertAfterMso='PasteGalleryMini',
            image_mso='PasteSingleCellExcelTableDestinationFormatting',
            on_action=cls.cb_pastereplace_action,
            # on_action=bkt.Callback(shape_selection.SlidesMore.paste_and_replace, slide=True, shape=True),
            get_enabled=cls.cb_pastereplace_enabled,
            get_visible=bkt.apps.ppt_shapes_exactly1_selected,
            **kwargs
        )
    
    ### Change language ###
    cb_lang_change = bkt.CallbackLazy("toolbox.language", "LangSetter", "get_dynamicmenu_content")
    @classmethod
    def change_lang_menu(cls, prefix, **kwargs):
        if "insertAfterMso" not in kwargs:
            kwargs["insertAfterMso"] = "ObjectFormatDialog"
        return bkt.ribbon.DynamicMenu(
                id=prefix+"-lang-change",
                label="Change language",
                supertip="Adjust the spell-check language for the selected shape(s)/text",
                image_mso="GroupLanguage",
                get_content=cls.cb_lang_change,
                **kwargs
            )



### Context menu for multiple shapes or grouped shape

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuObjectsGroup', children=[
        ### Lazy called BKT functions
        bkt.ribbon.DynamicMenu(
            label="BKT functions",
            supertip="Various BKT functions that are loaded dynamically for the selected shapes.",
            insertBeforeMso='Cut',
            image="bkt_logo",
            get_content=bkt.CallbackLazy("toolbox.contextmenus.dynamic", "ObjectsGroup", "get_children", shapes=True),
        ),
        ### Any shapes format sync
        bkt.ribbon.Button(id='context-format-sync', label="Match format", insertBeforeMso='Cut', image_mso="FormatPainter",
            supertip="Format all shapes like the shape under the cursor when the context menu is opened",
            on_action=bkt.CallbackLazy("toolbox.models.copy_paste_format", "FormatPainter", "cm_sync_shapes", shapes=True, context=True),
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut'),
        ### Connector functions (basically "copy" standard functions to multi-selection of connectors)
        bkt.ribbon.Menu(
            id='context-connectors-type',
            label="Connection types",
            supertip="Change connection type for all selected connectors",
            image_mso='ShapeConnectorStyleMenu',
            insertBeforeMso='ObjectsGroupMenu',
            get_visible=bkt.Callback(CtxVerbinder.ctx_connectors_visible, shapes=True),
            children=[
                bkt.ribbon.Button(
                    id='context-connectors-type-straight',
                    label="Straight connector",
                    image_mso='ShapeConnectorStyleStraight',
                    on_action=bkt.Callback(lambda shapes: CtxVerbinder.set_connector_type(shapes, 1), shapes=True),
                ),
                bkt.ribbon.Button(
                    id='context-connectors-type-elbow',
                    label="Angled connection",
                    image_mso='ShapeConnectorStyleElbow',
                    on_action=bkt.Callback(lambda shapes: CtxVerbinder.set_connector_type(shapes, 2), shapes=True),
                ),
                bkt.ribbon.Button(
                    id='context-connectors-type-curved',
                    label="Curved connection",
                    image_mso='ShapeConnectorStyleCurved',
                    on_action=bkt.Callback(lambda shapes: CtxVerbinder.set_connector_type(shapes, 3), shapes=True),
                ),
            ]
        ),
        bkt.ribbon.Button(
            id='context-connectors-reroute',
            label="Recreate connections",
            supertip="Recreate all selected connectors",
            insertBeforeMso='ObjectsGroupMenu',
            image_mso='ShapeRerouteConnectors',
            on_action=bkt.Callback(CtxVerbinder.reroute_connectors, shapes=True),
            get_visible=bkt.Callback(CtxVerbinder.ctx_connectors_visible, shapes=True),
            get_enabled=bkt.Callback(CtxVerbinder.ctx_connectors_reroute_enabled, shapes=True),
        ),
        bkt.ribbon.Button(
            id='context-connectors-invert',
            label="Reverse arrow direction",
            supertip="Reverse the arrow direction of the connector",
            insertBeforeMso='ObjectsGroupMenu',
            image_mso='ArrowStyleGallery',
            on_action=bkt.Callback(CtxVerbinder.invert_direction, shapes=True),
            get_visible=bkt.Callback(CtxVerbinder.ctx_connectors_visible, shapes=True),
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='ObjectsGroupMenu'),
        ### Language setting
        ContextMenuRecurring.change_lang_menu('ctx-shapes'),
        ### Clipboard operations
        ContextMenuRecurring.paste_replace_button('ctx-shapes'),
    ])
)


### Context menu for freeform shape type

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShapeFreeform', children=[
        ### Lazy called BKT functions
        bkt.ribbon.DynamicMenu(
            label="BKT functions",
            supertip="Various BKT functions that are loaded dynamically for the selected shapes.",
            insertBeforeMso='Cut',
            image="bkt_logo",
            get_content=bkt.CallbackLazy("toolbox.contextmenus.dynamic", "ShapeFreeform", "get_children", shape=True),
        ),
        bkt.ribbon.MenuSeparator(insertBeforeMso='Cut'),
        ### Clipboard operations
        ContextMenuRecurring.paste_replace_button('ctx-freeform'),
    ])
)


### Context menu for single "standard" shape

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShape', children=[
        ### Lazy called BKT functions
        bkt.ribbon.DynamicMenu(
            label="BKT functions",
            supertip="Various BKT functions that are loaded dynamically for the selected shapes.",
            insertBeforeMso='Cut',
            image="bkt_logo",
            get_content=bkt.CallbackLazy("toolbox.contextmenus.dynamic", "Shape", "get_children", shape=True),
        ),
        ### Language setting
        ContextMenuRecurring.change_lang_menu('ctx-shp'),
        ### Clipboard operations
        ContextMenuRecurring.paste_replace_button('ctx-shp'),
    ])
)


### Conext menu for text selection

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuTextEdit', children=[
        ### Language setting
        ContextMenuRecurring.change_lang_menu('ctx-txt'),
    ])
)


### Context menu for spell correction

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuSpell', children=[
        ### Language setting
        ContextMenuRecurring.change_lang_menu('ctx-spell'),
    ])
)


### Context menu for picture

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuPicture', children=[
        ### Lazy called BKT functions
        bkt.ribbon.DynamicMenu(
            label="BKT functions",
            supertip="Various BKT functions that are loaded dynamically for the selected shapes.",
            insertBeforeMso='Cut',
            image="bkt_logo",
            get_content=bkt.CallbackLazy("toolbox.contextmenus.dynamic", "Picture", "get_children", shape=True),
        ),
        ### Clipboard operations
        ContextMenuRecurring.paste_replace_button('ctx-pic'),
    ])
)


### Context menu for connector

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuShapeConnector', children=[
        ### Connector functions
        bkt.ribbon.Button(
            id='context-connector-invert',
            label="Reverse arrow direction",
            supertip="Reverse the arrow direction of the connector",
            insertAfterMso='ShapeRerouteConnectors',
            image_mso='ArrowStyleGallery',
            on_action=bkt.Callback(CtxVerbinder.invert_direction, shapes=True),
            #get_visible=bkt.Callback(CtxVerbinder.ctx_connectors_visible, shapes=True),
        ),
    ])
)


### Context menu for empty slide area

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuFrame', children=[
        ### nothing so far
    ])
)


### Context menu for slide thumbnails

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuThumbnail', children=[
        ### Copy to slides
        bkt.ribbon.Button(
            id='context-paste-to-slides',
            label="Insert on selected slides",
            supertip="Insert clipboard on all selected slides",
            insertAfterMso='PasteGalleryMini',
            image_mso='Paste',
            on_action=ContextMenuRecurring.cb_pastereplace_action,
            get_enabled=ContextMenuRecurring.cb_pastereplace_enabled,
        ),
        ### Language setting
        ContextMenuRecurring.change_lang_menu('ctx-slides', insertAfterMso='SlideBackgroundFormatDialog'),
        ### Export (send, save) selected slides
        bkt.ribbon.DynamicMenu(
            id="context-export-slides",
            label="Export selected slides",
            supertip="Export selected slides as a separate presentation or send as e-mail",
            image_mso="SaveSelectionToTextBoxGallery",
            insertAfterMso='SlideBackgroundFormatDialog',
            get_content=bkt.CallbackLazy("toolbox.contextmenus.slides", "ContextSlides", "get_children"),
        ),
        bkt.ribbon.MenuSeparator(insertAfterMso='SlideBackgroundFormatDialog'),
    ])
)


### Context menu for slide thumbnails in sort view

bkt.powerpoint.add_context_menu(
    bkt.ribbon.ContextMenu(id_mso='ContextMenuSlideSorter', children=[
        ### Export (send, save) selected slides
        bkt.ribbon.DynamicMenu(
            id="context-export2-slides",
            label="Export slides",
            supertip="Export selected slides as a separate presentation or send as e-mail",
            image_mso="SaveSelectionToTextBoxGallery",
            insertAfterMso='SlideBackgroundFormatDialog',
            get_content=bkt.CallbackLazy("toolbox.contextmenus.slides", "ContextSlides", "get_children"),
        ),
        bkt.ribbon.MenuSeparator(insertAfterMso='SlideBackgroundFormatDialog'),
    ])
)




# ==================
# = CONTEXTUAL TAB =
# ==================

picture_format_tab = bkt.ribbon.Tab(
    idMso = "TabPictureToolsFormat",
    children = [
        bkt.ribbon.Group(
            id="bkt_pictureformat_group",
            label="Format",
            insert_after_mso="GroupPictureTools",
            children = [
                bkt.ribbon.Button(
                    id = 'make_img_transparent',
                    label="Enable transparency",
                    supertip="Replaces the image with a shape with image fill that can be made natively transparent. The existing image is exported and then deleted, i.e. any cropped areas are lost and image formats cannot be undone.",
                    size="large",
                    show_label=True,
                    image_mso='PictureRecolorWashout',
                    on_action=bkt.Callback(PictureFormat.make_img_transparent),
                    # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
            ]
        )
    ]
)


# bkt.powerpoint.add_contextual_tab(
#     "TabSetDrawingTools",
#     harvey.harvey_ball_tab
# )
#use standard tab instead of contextual tab as contextual tab is not reliably shown (e.g. if PPT Format tab is hidden)
bkt.powerpoint.add_tab(harvey.harvey_ball_tab)

bkt.powerpoint.add_tab(agenda.agenda_tab)

bkt.powerpoint.add_tab(linkshapes.linkshapes_tab)

bkt.powerpoint.add_contextual_tab(
    "TabSetPictureTools",
    picture_format_tab
)

bkt.powerpoint.add_contextual_tab(
    "TabSetDrawingTools",
    info.context_format_tab
)





# ==========
# = POPUPS =
# ==========


bkt.powerpoint.context_dialogs.register("BKT_DIALOG_AMPEL3", "toolbox.popups.traffic_light") #traffic light
bkt.powerpoint.context_dialogs.register("BKT_DIALOG_STATESHAPE", "toolbox.popups.stateshapes") #stateshapes, e.g. likert scale

bkt.powerpoint.context_dialogs.register("BKT_PROCESS_CHEVRONS", "toolbox.popups.processshapes") #process chevrons - processshapes.ProcessChevrons.BKT_DIALOG_TAG
bkt.powerpoint.context_dialogs.register("BKT_SHAPE_HARVEY", "toolbox.popups.harvey") #harvey balls - harvey.HarveyBalls.BKT_HARVEY_DIALOG_TAG
bkt.powerpoint.context_dialogs.register("BKT_LINK_UUID", "toolbox.popups.linkshapes") #linked shapes - linkshapes.BKT_LINK_UUID
bkt.powerpoint.context_dialogs.register("TOOLBOX-AGENDA-POPUP", "toolbox.popups.agenda") #agenda textbox - agenda.TOOLBOX_AGENDA_POPUP