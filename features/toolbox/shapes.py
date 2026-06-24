# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''



import logging
# import locale

# from System import Array

import bkt
import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt
get_ambiguity_tuple = bkt.helpers.get_ambiguity_tuple

# from bkt.library.algorithms import get_bounding_nodes, mid_point

# from bkt import dotnet
# Drawing = dotnet.import_drawing()
# office = dotnet.import_officecore()

# other toolbox modules
from .chartlib import shapelib_button
# from .agenda import ToolboxAgenda
from . import text
# from . import harvey
# from . import stateshapes



class PositionSize(object):
    use_visual_pos  = bkt.settings.get("toolbox.possize.use_visual_pos", False)
    use_visual_size = bkt.settings.get("toolbox.possize.use_visual_size", False)

    @classmethod
    def toggle_use_visual_pos(cls):
        cls.use_visual_pos = not cls.use_visual_pos
        bkt.settings["toolbox.possize.use_visual_pos"] = cls.use_visual_pos

    @classmethod
    def get_image_use_visual_pos(cls):
        return bkt.ribbon.Gallery.get_check_image(cls.use_visual_pos)

    @classmethod
    def toggle_use_visual_size(cls):
        cls.use_visual_size = not cls.use_visual_size
        bkt.settings["toolbox.possize.use_visual_size"] = cls.use_visual_size

    @classmethod
    def get_image_use_visual_size(cls):
        return bkt.ribbon.Gallery.get_check_image(cls.use_visual_size)

    @classmethod
    def set_top(cls, shapes, value):
        attr = 'visual_top' if cls.use_visual_pos else 'top'
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, attr, value), 
            lambda shape: getattr(shape, attr), 
            shapes, value)
    
    @classmethod
    def get_top(cls, shapes):
        if not cls.use_visual_pos:
            return get_ambiguity_tuple(shape.top for shape in shapes) #shapes[0].top
        else:
            return get_ambiguity_tuple(shape.visual_top for shape in shapes)
    
    
    @classmethod
    def set_left(cls, shapes, value):
        attr = 'visual_left' if cls.use_visual_pos else 'left'
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, attr, value), 
            lambda shape: getattr(shape, attr), 
            shapes, value)

    @classmethod
    def get_left(cls, shapes):
        if not cls.use_visual_pos:
            return get_ambiguity_tuple(shape.left for shape in shapes) #shapes[0].left
        else:
            return get_ambiguity_tuple(shape.visual_left for shape in shapes)


    @classmethod
    def set_height(cls, shapes, value):
        attr = 'visual_height' if cls.use_visual_size else 'height'
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, attr, value), 
            lambda shape: getattr(shape, attr), 
            shapes, value)

    @classmethod
    def get_height(cls, shapes):
        if not cls.use_visual_size:
            return get_ambiguity_tuple(shape.height for shape in shapes) #shapes[0].height
        else:
            return get_ambiguity_tuple(shape.visual_height for shape in shapes)
    
    
    @classmethod
    def set_width(cls, shapes, value):
        attr = 'visual_width' if cls.use_visual_size else 'width'
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, attr, value), 
            lambda shape: getattr(shape, attr), 
            shapes, value)

    @classmethod
    def get_width(cls, shapes):
        if not cls.use_visual_size:
            return get_ambiguity_tuple(shape.width for shape in shapes) #shapes[0].width
        else:
            return get_ambiguity_tuple(shape.visual_width for shape in shapes)


    @staticmethod
    def set_zorder(shapes, value):
        delta = int(value) - shapes[0].ZOrderPosition
        shapes = sorted(shapes, key=lambda shape: shape.ZOrderPosition, reverse=True if delta > 0 else False)
        for shape in shapes:
            pplib.set_shape_zorder(shape, delta=delta)
        # Normal behavior too confusing for users:
        # bkt.apply_delta_on_ALT_key(
        #     PositionSize._set_shape_zorder, 
        #     lambda shape: shape.ZOrderPosition, 
        #     shapes, int(value))

    @staticmethod
    def get_zorder(shapes):
        if len(shapes) == 1:
            return shapes[0].ZOrderPosition
        else:
            return (True, shapes[0].ZOrderPosition) #force ambiguous mode

    @staticmethod
    def front_to_back(shapes):
        shapes = sorted(shapes, key=lambda shape: shape.ZOrderPosition, reverse=True)
        target_zorder = shapes.pop(-1).ZOrderPosition
        for shape in shapes:
            pplib.set_shape_zorder(shape, value=target_zorder)

    @staticmethod
    def back_to_front(shapes):
        shapes = sorted(shapes, key=lambda shape: shape.ZOrderPosition, reverse=False)
        target_zorder = shapes.pop(-1).ZOrderPosition
        for shape in shapes:
            pplib.set_shape_zorder(shape, value=target_zorder)

    @staticmethod
    def zorder_top2bottom(shapes, reverse=False):
        shapes = sorted(shapes, key=lambda shape: shape.Top, reverse=reverse)
        start = shapes[0].ZOrderPosition
        for shape in shapes:
            pplib.set_shape_zorder(shape, value=start)
            start += 1
        #update selection
        pplib.shapes_to_range(shapes).select()

    @classmethod
    def zorder_bottom2top(cls, shapes):
        cls.zorder_top2bottom(shapes, True)

    @staticmethod
    def zorder_left2right(shapes, reverse=False):
        shapes = sorted(shapes, key=lambda shape: shape.Left, reverse=reverse)
        start = shapes[0].ZOrderPosition
        for shape in shapes:
            pplib.set_shape_zorder(shape, value=start)
            start += 1
        #update selection
        pplib.shapes_to_range(shapes).select()

    @classmethod
    def zorder_right2left(cls, shapes):
        cls.zorder_left2right(shapes, True)
    
    @staticmethod
    def set_height_to_width(shapes):
        for shape in shapes:
            shape.square(w2h=False)
    
    @staticmethod
    def set_width_to_height(shapes):
        for shape in shapes:
            shape.square(w2h=True)
    
    @staticmethod
    def swap_width_and_height(shapes):
        for shape in shapes:
            shape.transpose()
    
    @staticmethod
    def set_top_to_left(shapes):
        for shape in shapes:
            shape.top = shape.left
    
    @staticmethod
    def set_left_to_top(shapes):
        for shape in shapes:
            shape.left = shape.top
    
    @staticmethod
    def swap_left_and_top(shapes):
        for shape in shapes:
            shape.top, shape.left = shape.left, shape.top


class AspectRatio(object):
    types_scale = (
        pplib.MsoShapeType["msoPicture"],
        pplib.MsoShapeType["msoLinkedPicture"],
        pplib.MsoShapeType["msoFreeform"],
        pplib.MsoShapeType["msoEmbeddedOLEObject"],
        pplib.MsoShapeType["msoLinkedOLEObject"],
        pplib.MsoShapeType["msoMedia"],
    )
    types_in_db = (
        pplib.MsoShapeType["msoAutoShape"],
        pplib.MsoShapeType["msoCallout"],
    )

    aspect_ratios = [
        (1,1),
        (3,2),
        (4,3),
        (13,9),
        (15,10),
        (16,9),
    ]

    @classmethod
    def reset(cls, shapes):
        for shape in shapes:
            try:
                shape_type = shape.Type
                #TODO: placeholder support
                if shape_type in cls.types_scale:
                    height = shape.Height
                    shape.ScaleHeight(1, True)
                    shape.ScaleWidth(1, True)
                    #reapply ratio (only required if LockAspectRatio=0)
                    ratio = shape.Width/shape.Height
                    shape.Height = height
                    shape.Width = ratio*height
                elif shape_type in cls.types_in_db:
                    try:
                        shape_db = pplib.GlobalShapeDb.get_by_shape(shape)
                        ratio = shape_db["ratio"]
                    except:
                        logging.exception("shape not found in db")
                        ratio = 1
                    # landscape = shape.width > shape.height
                    shape.force_aspect_ratio(ratio)
            except:
                continue
    
    @staticmethod
    def swap(shapes):
        for shape in shapes:
            shape.transpose()
    
    @classmethod
    def set_aspect_ratio(cls, shapes, current_control):
        index = int(current_control["tag"])
        r1,r2 = cls.aspect_ratios[index]
        value = r1/r2
        for shape in shapes:
            # landscape = shape.width > shape.height
            shape.force_aspect_ratio(value)
            shape.lock_aspect_ratio = True
    
    @staticmethod
    def get_aspect_ratio(shape):
        return shape.width/shape.height
    
    @classmethod
    def get_aspect_ratio_label(cls, context):
        try:
            return "Aktuelles Seiteverhältnis: {:.4n}".format(cls.get_aspect_ratio(context.shape))
        except:
            return "Aktuelles Seiteverhältnis: -"

    @staticmethod
    def lock_aspect_ratio(shapes, pressed):
        for shape in shapes:
            shape.LockAspectRatio = -1 if pressed else 0
    
    @staticmethod
    def get_aspect_ratio_locked(shapes):
        return shapes[0].LockAspectRatio == -1


spinner_top = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_top",
    image_mso='ObjectNudgeDown',
    label="Position from top",
    show_label=False,
    screentip="Position from top",
    supertip="Change of position from the top.\n\nWith the CTRL key held, change by 0.1 cm instead of 0.2 cm.\n\nWith the ALT key held, change relative per shape (if multiple shapes are selected).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_top, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_top, shapes=True, wrap_shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ObjectNudgeDown', children=[
        bkt.ribbon.Button(
            label="Visual position",
            get_image=bkt.Callback(PositionSize.get_image_use_visual_pos),
            supertip="Use visual position taking rotation into account",
            on_action=bkt.Callback(PositionSize.toggle_use_visual_pos)
        ),
        bkt.ribbon.Button(
            label="Top = Left",
            image="possize_t2l",
            screentip="Set top = left",
            supertip="Sets the top edge equal to the left edge, taking the fixed point into account",
            on_action=bkt.Callback(PositionSize.set_top_to_left, shapes=True, wrap_shapes=True)
        ),
        bkt.ribbon.Button(
            label="Top ⇄ Left",
            image="possize_swap_tl",
            screentip="Swap top and left",
            supertip="Swaps the top edge with the left edge, taking the fixed point into account",
            on_action=bkt.Callback(PositionSize.swap_left_and_top, shapes=True, wrap_shapes=True)
        ),
    ])
)

spinner_left = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_left",
    image_mso='ObjectNudgeRight',
    label="Position from left",
    show_label=False,
    screentip="Position from left",
    supertip="Change of position from the left.\n\nWith the CTRL key held, change by 0.1 cm instead of 0.2 cm.\n\nWith the ALT key held, change relative per shape (if multiple shapes are selected).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_left, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_left, shapes=True, wrap_shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ObjectNudgeRight', children=[
        bkt.ribbon.Button(
            label="Visual position",
            get_image=bkt.Callback(PositionSize.get_image_use_visual_pos),
            supertip="Use visual position taking rotation into account",
            on_action=bkt.Callback(PositionSize.toggle_use_visual_pos)
        ),
        bkt.ribbon.Button(
            label="Left = Top",
            image="possize_l2t",
            screentip="Set left = top",
            supertip="Sets the left edge equal to the top edge, taking the fixed point into account",
            on_action=bkt.Callback(PositionSize.set_left_to_top, shapes=True, wrap_shapes=True)
        ),
        bkt.ribbon.Button(
            label="Left ⇄ Top",
            image="possize_swap_tl",
            screentip="Swap left and top",
            supertip="Swaps the left edge with the top edge, taking the fixed point into account",
            on_action=bkt.Callback(PositionSize.swap_left_and_top, shapes=True, wrap_shapes=True)
        ),
    ])
)

spinner_height = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_height",
    image_mso='ShapeHeight',
    label="Height",
    show_label=False,
    screentip="Height",
    supertip="Change of height.\n\nWith the CTRL key held, change by 0.1 cm instead of 0.2 cm.\n\nWith the ALT key held, change relative per shape (if multiple shapes are selected).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_height, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_height, shapes=True, wrap_shapes=True),
    get_enabled=bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ShapeHeight', children=[
        bkt.ribbon.Button(
            label="Visual size",
            get_image=bkt.Callback(PositionSize.get_image_use_visual_size),
            supertip="Use visual size taking rotation into account",
            on_action=bkt.Callback(PositionSize.toggle_use_visual_size)
        ),
        bkt.ribbon.Button(
            label="Height = Width",
            image="possize_h2w",
            screentip="Set height = width",
            supertip="Sets the height equal to the width, taking the fixed point into account. If the aspect ratio is locked, this is temporarily released.",
            on_action=bkt.Callback(PositionSize.set_height_to_width, shapes=True, wrap_shapes=True)
        ),
        bkt.ribbon.Button(
            label="Height ⇄ Width",
            image="possize_swap_hw",
            screentip="Swap height and width",
            supertip="Swaps the height with the width, taking the fixed point into account",
            on_action=bkt.Callback(PositionSize.swap_width_and_height, shapes=True, wrap_shapes=True)
        ),
    ])
)

spinner_width = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_width",
    image_mso='ShapeWidth',
    label="Width",
    show_label=False,
    screentip="Width",
    supertip="Change of width.\n\nWith the CTRL key held, change by 0.1 cm instead of 0.2 cm.\n\nWith the ALT key held, change relative per shape (if multiple shapes are selected).",
    round_cm=True,
    on_change=bkt.Callback(PositionSize.set_width, shapes=True, wrap_shapes=True),
    get_text=bkt.Callback(PositionSize.get_width, shapes=True, wrap_shapes=True),
    get_enabled=bkt.apps.ppt_shapes_or_text_selected,
    convert="pt_to_cm",
    image_element=pplib.LocpinGallery(image_mso='ShapeWidth', children=[
        bkt.ribbon.Button(
            label="Visual size",
            get_image=bkt.Callback(PositionSize.get_image_use_visual_size),
            supertip="Use visual size taking rotation into account",
            on_action=bkt.Callback(PositionSize.toggle_use_visual_size)
        ),
        bkt.ribbon.Button(
            label="Width = Height",
            image="possize_w2h",
            screentip="Set width = height",
            supertip="Sets the width equal to the height, taking the fixed point into account. If the aspect ratio is locked, this is temporarily released.",
            on_action=bkt.Callback(PositionSize.set_width_to_height, shapes=True, wrap_shapes=True)
        ),
        bkt.ribbon.Button(
            label="Width ⇄ Height",
            image="possize_swap_hw",
            screentip="Swap width and height",
            supertip="Swaps the width with the height, taking the fixed point into account",
            on_action=bkt.Callback(PositionSize.swap_width_and_height, shapes=True, wrap_shapes=True)
        ),
    ])
)

spinner_zorder = bkt.ribbon.RoundingSpinnerBox(
    id="pos_size_spinner_zorder",
    image_mso='ObjectBringForward',
    label="Z-Order",
    show_label=False,
    screentip="Z-Order",
    supertip="Change of the Z-order, i.e. the order of the shapes on the slide.",
    on_change=bkt.Callback(PositionSize.set_zorder, shapes=True),
    get_text=bkt.Callback(PositionSize.get_zorder, shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
    round_int=True,
    small_step=1,
    big_step=1,
    image_element=bkt.ribbon.Menu(
        children=[
            bkt.mso.control.ObjectBringToFront,
            bkt.mso.control.ObjectSendToBack,
            bkt.ribbon.MenuSeparator(title="Adjust"),
            bkt.ribbon.Button(
                label="Front to back",
                supertip="Brings all front shapes exactly behind the backmost shape",
                image="zorder_front_to_back",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.front_to_back, shapes=True),
            ),
            bkt.ribbon.Button(
                label="Back to front",
                supertip="Brings all back shapes exactly in front of the frontmost shape",
                image="zorder_back_to_front",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.back_to_front, shapes=True),
            ),
            bkt.ribbon.MenuSeparator(title="Sort"),
            bkt.ribbon.Button(
                label="Top to bottom",
                supertip="Sorts the Z-order from top to bottom so that the bottom shape becomes the frontmost",
                image="zorder_top_to_bottom",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.zorder_top2bottom, shapes=True),
            ),
            bkt.ribbon.Button(
                label="Bottom to top",
                supertip="Sorts the Z-order from bottom to top so that the top shape becomes the frontmost",
                image="zorder_bottom_to_top",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.zorder_bottom2top, shapes=True),
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                label="Left to right",
                supertip="Sorts the Z-order from left to right so that the right shape becomes the frontmost",
                image="zorder_left_to_right",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.zorder_left2right, shapes=True),
            ),
            bkt.ribbon.Button(
                label="Right to left",
                supertip="Sorts the Z-order from right to left so that the left shape becomes the frontmost",
                image="zorder_right_to_left",
                get_enabled=bkt.apps.ppt_shapes_min2_selected,
                on_action=bkt.Callback(PositionSize.zorder_right2left, shapes=True),
            ),
        ],
    ),
)

#button_lock_aspect_ratio = bkt.ribbon.CheckBox(
button_lock_aspect_ratio = dict(
    #id = 'shape_lock_aspect_ratio',
    # label="Aspect ratio",
    screentip="Lock aspect ratio",
    supertip="When the Lock aspect ratio checkbox is enabled, the height and width settings change in proportion to each other.",
    on_toggle_action = bkt.Callback(AspectRatio.lock_aspect_ratio, shapes=True),
    get_pressed = bkt.Callback(AspectRatio.get_aspect_ratio_locked, shapes=True),
    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
)

menu_lock_aspect_ratio = bkt.ribbon.Box(
    box_style="horizontal",
    children=[
        bkt.ribbon.Menu(
            label="Lock",
            show_label=False,
            image_mso="AutoSizePage",
            children=[
                bkt.ribbon.ToggleButton(id="shape_lock_aspect_ratio3", label="Toggle lock aspect ratio", image_mso="Lock", **button_lock_aspect_ratio),
                bkt.ribbon.Button(
                    get_label=bkt.Callback(AspectRatio.get_aspect_ratio_label, context=True),
                    enabled=False,
                ),
                bkt.ribbon.MenuSeparator(),
            ] + [
                bkt.ribbon.Button(
                    label="Set to {}:{} ({:.4n})".format(r[0], r[1], r[0]/r[1]),
                    tag=str(i),
                    on_action=bkt.Callback(AspectRatio.set_aspect_ratio, shapes=True, current_control=True, wrap_shapes=True),
                ) for i, r in enumerate(AspectRatio.aspect_ratios)
            ] + [
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    label="Swap",
                    screentip="Swap aspect ratio",
                    supertip="Swaps width and height and thereby reverses the aspect ratio.",
                    image_mso="PageScaleToFitOptionsDialog",
                    on_action = bkt.Callback(AspectRatio.swap, shapes=True, wrap_shapes=True),
                ),
                bkt.ribbon.Button(
                    label="Reset",
                    screentip="Reset aspect ratio",
                    supertip="Resets the aspect ratio to its original state.",
                    image_mso="ResetCurrentView",
                    on_action = bkt.Callback(AspectRatio.reset, shapes=True, wrap_shapes=True),
                ),
                # bkt.mso.control.PictureResetAndSize,
            ]
        ),
        bkt.ribbon.CheckBox(id="shape_lock_aspect_ratio2", label="Aspect r.", **button_lock_aspect_ratio),
        # bkt.ribbon.ToggleButton(
        #     label="Locked",
        #     # show_label=False,
        #     image_mso="Lock",
        # ),
        # bkt.ribbon.ToggleButton(
        #     label="Open",
        #     show_label=False,
        #     image_mso="Lock",
        # ),
    ]
)

size_group = bkt.ribbon.Group(
    id="bkt_size_group",
    label='Size',
    image_mso='GroupSizeAndPosition',
    children =[
        #spinner_height,
        #spinner_width,
        bkt.mso.control.ShapeHeight(show_label=False),
        bkt.mso.control.ShapeWidth(show_label=False),
        bkt.ribbon.CheckBox(id="shape_lock_aspect_ratio1", label="Aspect ratio", **button_lock_aspect_ratio),
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectSizeAndPositionDialog')
    ]
)

# pos_group = bkt.ribbon.Group(
#     label='Position',
#     image_mso='GroupSizeAndPosition',
#     children =[
#         spinner_top,
#         spinner_left,
#         spinner_zorder,
#         bkt.ribbon.DialogBoxLauncher(idMso='ObjectSizeAndPositionDialog')
#     ]
# )

pos_size_group = bkt.ribbon.Group(
    id="bkt_possize_group",
    label='Position/size',
    image_mso='GroupSizeAndPosition',
    children =[
        spinner_height,
        spinner_width,
        menu_lock_aspect_ratio,
        # bkt.ribbon.CheckBox(id="shape_lock_aspect_ratio2", **button_lock_aspect_ratio),
        spinner_top,
        spinner_left,
        spinner_zorder,
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectSizeAndPositionDialog')
    ]
)



class ShapeFormats(object):
    transparencies = list(range(0, 110, 10))

    @classmethod
    def _attr_setter(cls, shape, value, shp_object, attribute):
        try:
            if attribute == "Transparency":
                value = min(max(0, value/100),100)
            else:
                value = max(0, value)
            shp_object = getattr(shape, shp_object)
            setattr(shp_object, "visible", -1)
            setattr(shp_object, attribute, value)
        except:
            logging.exception("Setting %s attribute %s to value %s failed!", shp_object, attribute, value)
    @classmethod
    def _attr_getter(cls, shape, shp_object, attribute):
        try:
            shp_object = getattr(shape, shp_object)
            value = max(0, getattr(shp_object, attribute))
            if attribute == "Transparency":
                value = value*100
            return value
        except:
            logging.exception("Getting %s attribute %s failed!", shp_object, attribute)
            return 0

    ### Fill properties ###
    @classmethod
    def get_fill_enabled(cls, context):
        #TESTME: is fill implemented for all shape types? (see also problem with line)
        # shape = next(pplib.iterate_shape_subshapes(shapes))
        # return shape.Fill.visible == -1
        
        # copy enabled status of fill-button
        return context.app.commandbars.GetEnabledMso("ShapeFillColorPicker")

    @classmethod
    def get_fill_transparency(cls, shapes):
        shapes = pplib.iterate_shape_subshapes(shapes)
        for shape in shapes:
            try:
                return max(0, round(float(shape.fill.transparency)*100))
            except:
                continue
        return None
    
    @classmethod
    def set_fill_transparency(cls, shapes, value):
        value = min(max(0, value),100) #min=0, max=100
        shapes = list(pplib.iterate_shape_subshapes(shapes))
        bkt.apply_delta_on_ALT_key(
            # lambda shape, value: setattr(shape.Fill, 'Transparency', min(max(0, value/100),100)), 
            cls._attr_setter,
            cls._attr_getter,
            shapes, value, shp_object="Fill", attribute="Transparency")

    ### Line properties ###
    @classmethod
    def get_line_enabled(cls, context):
        # return len(cls._line_filter(shapes)) > 0
        # shape = next(pplib.iterate_shape_subshapes(shapes))
        # try:
        #     return hasattr(shape.line, "visible")
        # except ValueError:
        #     return False

        # copy enabled status of line-button
        return context.app.commandbars.GetEnabledMso("ShapeOutlineColorPicker")

    @classmethod
    def get_line_transparency(cls, shapes):
        shapes = pplib.iterate_shape_subshapes(shapes, exclude=[pplib.MsoShapeType['msoTable']])
        #IMPORTANT: if tables are not excluded, Powerpoint will crash if a table is selected and this function is executed
        for shape in shapes:
            try:
                return max(0, round(float(shape.line.transparency)*100))
            except:
                continue
        return None
    
    @classmethod
    def set_line_transparency(cls, shapes, value):
        value = min(max(0, value),100) #min=0, max=100
        shapes = pplib.iterate_shape_subshapes(shapes)
        bkt.apply_delta_on_ALT_key(
            # lambda shape, value: setattr(shape.Line, 'Transparency', min(max(0, value/100),100)), 
            cls._attr_setter,
            cls._attr_getter,
            shapes, value, shp_object="Line", attribute="Transparency")

    @classmethod
    def get_line_weight(cls, shapes):
        shapes = pplib.iterate_shape_subshapes(shapes, exclude=[pplib.MsoShapeType['msoTable']])
        #IMPORTANT: if tables are not excluded, Powerpoint will crash if a table is selected and this function is executed
        for shape in shapes:
            try:
                return max(0, shape.line.weight)
            except:
                continue
        return None
    
    @classmethod
    def set_line_weight(cls, shapes, value):
        value = max(0, value)
        shapes = list(pplib.iterate_shape_subshapes(shapes))
        bkt.apply_delta_on_ALT_key(
            # lambda shape, value: setattr(shape.Line, 'weight', max(0, value)), 
            cls._attr_setter,
            cls._attr_getter,
            shapes, value, shp_object="Line", attribute="weight")

    ### GALLERY ###
    @classmethod
    def get_item_count(cls):
        return len(cls.transparencies)
    
    @classmethod
    def get_item_label(cls, index):
        return "%s%%" % cls.transparencies[index]
    
    @classmethod
    def get_item_image(cls, index, context):
        return cls._get_image_for_transp(cls.transparencies[index], context)
    
    @classmethod
    def _get_image_for_transp(cls, transp, context):
        return context.python_addin.load_image( "transp_%s" % int(round(transp/10.0)*10) )


    @classmethod
    def fill_on_action_indexed(cls, selected_item, index, shapes):
        value = float(cls.transparencies[index])
        cls.set_fill_transparency(shapes, value)
    
    @classmethod
    def fill_get_selected_item_index(cls, context):
        try:
            return cls.transparencies.index(cls.get_fill_transparency(context.shapes))
        except:
            return -1

    @classmethod
    def line_on_action_indexed(cls, selected_item, index, shapes):
        value = float(cls.transparencies[index])
        cls.set_line_transparency(shapes, value)
    
    @classmethod
    def line_get_selected_item_index(cls, context):
        try:
            return cls.transparencies.index(cls.get_line_transparency(context.shapes))
        except:
            return -1


format_group = bkt.ribbon.Group(
    id="bkt_format_group",
    label="Format",
    image_mso='BehindText',
    children=[
        bkt.ribbon.RoundingSpinnerBox(
            id = 'fill_transparency',
            label="Transparency background",
            supertip="Change the transparency of the background",
            show_label=False,
            round_int = True,
            image="fill_transparency",
            on_change = bkt.Callback(ShapeFormats.set_fill_transparency, shapes=True),
            get_text  = bkt.Callback(ShapeFormats.get_fill_transparency, shapes=True),
            get_enabled = bkt.Callback(ShapeFormats.get_fill_enabled, context=True),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'line_transparency',
            label="Transparency line/border",
            supertip="Change the transparency of the border or line",
            show_label=False,
            round_int = True,
            image="line_transparency",
            on_change = bkt.Callback(ShapeFormats.set_line_transparency, shapes=True),
            get_text  = bkt.Callback(ShapeFormats.get_line_transparency, shapes=True),
            get_enabled = bkt.Callback(ShapeFormats.get_line_enabled, context=True),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'line_weight',
            label="Thick line/border",
            supertip="Change the thickness of the border or line",
            show_label=False,
            round_pt = True,
            rounding_factor=0.25,
            huge_step=1,
            big_step=0.5,
            small_step=0.25,
            image_mso="LineThickness",
            on_change = bkt.Callback(ShapeFormats.set_line_weight, shapes=True),
            get_text  = bkt.Callback(ShapeFormats.get_line_weight, shapes=True),
            get_enabled = bkt.Callback(ShapeFormats.get_line_enabled, context=True),
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectFormatDialog')
    ]
)

fill_transparency_gallery = bkt.ribbon.Gallery(
    id="bkt_fill_transparency_menu",
    label="Transparency background",
    supertip="Sets the background transparency to the chosen value.",
    show_label=False,
    show_item_label=True,
    image="fill_transparency",
    columns="1",
    get_enabled = bkt.Callback(ShapeFormats.get_fill_enabled, context=True),
    on_action_indexed = bkt.Callback(ShapeFormats.fill_on_action_indexed, shapes=True),
    get_selected_item_index = bkt.Callback(ShapeFormats.fill_get_selected_item_index, context=True),
    item_height="16",
    get_item_count=bkt.Callback(ShapeFormats.get_item_count),
    get_item_label=bkt.Callback(ShapeFormats.get_item_label),
    get_item_image=bkt.Callback(ShapeFormats.get_item_image, context=True),
    ### static definition of children has disadvantage that get_selected_item_index is called even if nothing 
    ### is selected, leading to an error message on ppt startup if UI error are cative.
    # children=[
    #     bkt.ribbon.Item(label="%s%%" % transp, image="transp_%s" % transp)
    #     for transp in ShapeFormats.transparencies
    # ]
)

line_transparency_gallery = bkt.ribbon.Gallery(
    id="bkt_line_transparency_menu",
    label="Transparency line/border",
    supertip="Sets the line transparency to the chosen value.",
    show_label=False,
    show_item_label=True,
    image="line_transparency",
    columns="1",
    get_enabled = bkt.Callback(ShapeFormats.get_line_enabled, context=True),
    on_action_indexed = bkt.Callback(ShapeFormats.line_on_action_indexed, shapes=True),
    get_selected_item_index = bkt.Callback(ShapeFormats.line_get_selected_item_index, context=True),
    item_height="16",
    get_item_count=bkt.Callback(ShapeFormats.get_item_count),
    get_item_label=bkt.Callback(ShapeFormats.get_item_label),
    get_item_image=bkt.Callback(ShapeFormats.get_item_image, context=True),
    ### static definition of children has disadvantage that get_selected_item_index is called even if nothing 
    ### is selected, leading to an error message on ppt startup if UI error are cative.
    # children=[
    #     bkt.ribbon.Item(label="%s%%" % transp, image="transp_%s" % transp)
    #     for transp in ShapeFormats.transparencies
    # ]
)


# default ui for shape styling
styles_group = bkt.ribbon.Group(
    id="bkt_style_group",
    label='Styles',
    image_mso='ShapeFillColorPicker',
    children = [
        bkt.mso.splitbutton.ShapeFillColorPicker,
        bkt.mso.splitbutton.ShapeOutlineColorPicker,
        bkt.mso.control.ShapeEffectsMenu,
        bkt.mso.splitbutton.TextFillColorPicker,
        bkt.mso.splitbutton.TextOutlineColorPicker,
        bkt.mso.control.TextEffectsMenu,
        bkt.mso.control.OutlineWeightGallery,
        bkt.mso.control.OutlineDashesGallery,
        bkt.mso.control.ArrowStyleGallery,
        fill_transparency_gallery,
        line_transparency_gallery,
        bkt.mso.control.ShapeQuickStylesHome, #if ppt_customformats is active, this button is replaced
        bkt.ribbon.DialogBoxLauncher(idMso='ObjectFormatDialog')
    ]
)


shapes_group = bkt.ribbon.Group(
    id="bkt_shapes_group",
    label='Shapes',
    image_mso='ShapesInsertGallery',
    children = [
        bkt.mso.control.ShapesInsertGallery,
        text.text_splitbutton,
        bkt.ribbon.DynamicMenu(
            image_mso='TableInsertGallery',
            label="Insert table",
            show_label=False,
            supertip="Insert standard or shape tables",
            # item_size="large", #not supported by dynamic menu
            get_content=bkt.CallbackLazy("toolbox.models.shapes_menu", "shapes_table_menu"),
        ),
        
        #bkt.mso.control.PictureInsertFromFilePowerPoint,
        shapelib_button,
        text.symbol_insert_splitbutton,
        bkt.ribbon.DynamicMenu(
            label='Special shapes',
            show_label=False,
            image_mso='SmartArtInsert',
            screentip="Special and interactive shapes",
            supertip="Insert interactive BKT shapes and special composite shapes that are otherwise cumbersome to create.",
            get_content=bkt.CallbackLazy("toolbox.models.shapes_menu", "shapes_interactive_menu"),
        ),
        bkt.mso.control.ShapeChangeShapeGallery,
        bkt.ribbon.DynamicMenu(
            image_mso='CombineShapesMenu',
            label="Modify shape",
            supertip="Functions to manipulate shape points, duplicate shapes, and convert text into symbol/graphic",
            show_label=False,
            get_content=bkt.CallbackLazy("toolbox.models.shapes_menu", "shapes_change_menu"),
        ),
        bkt.ribbon.DynamicMenu(
            label='More',
            show_label=False,
            image_mso='TableDesign',
            screentip="More functions",
            supertip="Insert standard objects (images, SmartArt, etc.), hide and show shapes again, adjust header and footer",
            get_content=bkt.CallbackLazy("toolbox.models.shapes_menu", "shapes_more_menu"),
        ),
    ]
)
