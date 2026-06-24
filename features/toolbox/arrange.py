# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

import logging

import math
# from heapq import nsmallest, nlargest

import bkt

import bkt.library.algorithms as algos
import bkt.library.powerpoint as pplib
pt_to_cm = pplib.pt_to_cm
cm_to_pt = pplib.cm_to_pt

# from .linkshapes import LinkedShapes
# from .shapes import PositionSize


class MasterShapeHandler(object):
    def __init__(self):
        #instance variables:
        
        #save preference for master shape (first or last selected) into class variables
        self.fallback_first_last = bkt.settings.get("arrange_advanced.default", "LAST")
        self.master = self.fallback_first_last
        #save preference to show master shape indicator into class variables
        self.master_indicator = bkt.settings.get("arrange_advanced.master_indicator", True)
        if self.master_indicator:
            self._register_dialog()


    def set_master_indicator(self, pressed):
        ''' callback: set whether master shape indicator is shown '''
        if not pressed:
            self.master_indicator = False
            bkt.settings["arrange_advanced.master_indicator"] = False
            self._unregister_dialog()
        else:
            self.master_indicator = True
            bkt.settings["arrange_advanced.master_indicator"] = True
            self._register_dialog()

    def get_master_indicator(self):
        ''' returns whether master shape indicator is shown '''
        return self.master_indicator

    def get_master_for_indicator(self, shapes):
        if self.master=="FIRST":
            return shapes[0]
        elif self.master=="LAST":
            return shapes[-1]
        else:
            return None

    def _register_dialog(self):
        from .popups.master import MasterShapeDialog
        # register dialog
        bkt.powerpoint.context_dialogs.register_dialog(
            MasterShapeDialog(self)
        )

    def _unregister_dialog(self):
        # unregister dialog
        bkt.powerpoint.context_dialogs.unregister("MASTER")


GlobalMasterShape = MasterShapeHandler()


class LocPinFactory(object):
    def __init__(self):
        self.locpins = dict()

    @property
    def swap(self):
        return self.locpins.setdefault("swap", pplib.LocPin(settings_key="toolbox.swap_locpin"))
    
    @property
    def dis1(self):
        return self.locpins.setdefault("dis1", pplib.LocPin(4))
    @property
    def dis2(self):
        return self.locpins.setdefault("dis2", pplib.LocPin(4))

    @property
    def rotation(self):
        return self.locpins.setdefault("rotation", pplib.LocPin(4, "toolbox.rotation_locpin")) #center point as initial locpin)


LocPinCollection = LocPinFactory()



swap_button = bkt.ribbon.SplitButtonFixed(
    show_label=False,
    get_enabled=bkt.apps.ppt_shapes_min2_selected,
    children=[
        bkt.ribbon.Button(
            id = 'swap',
            label="Swap",
            image_mso='CircularReferences',
            screentip="Swap shape position",
            supertip="Swap the position (left/top) of the selected shapes.",
            on_action=bkt.CallbackLazy("toolbox.models.arrange", "Swap", "multi_swap", shapes=True, shapes_min=2),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Menu(label="Swap menu", supertip="Functions to swap shape position, size or formatting", children=[
            bkt.ribbon.MenuSeparator(title="Swap selected shapes"),
            bkt.ribbon.Button(
                id = 'swap2',
                label="Swap position",
                image_mso='CircularReferences', #'MailMergeMatchFields'
                screentip="Swap shape position",
                supertip="Swap the position (left/top) of the selected shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "Swap", "multi_swap", shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'swap_pos_and_size',
                label="Swap position and size [Shift]",
                show_label=True,
                #image_mso='MailMergeMatchFields',
                screentip="Swap shape position",
                supertip="Swap the position (left/top) and size of the selected shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "Swap", "multi_swap_pos_size", shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            pplib.LocpinGallery(
                label="Anchor point when swapping",
                screentip="Set anchor point when swapping",
                supertip="Sets the point that should be fixed when swapping the shapes.",
                locpin=LocPinCollection.swap,
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                id = 'swap_left',
                label="Swap x position",
                show_label=True,
                screentip="Swap x position",
                supertip="Swap the x position (distance from the left slide edge) of the selected shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "Swap", "multi_swap_left", shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'swap_top',
                label="Swap y position",
                show_label=True,
                screentip="Swap y position",
                supertip="Swap the y position (distance from the top slide edge) of the selected shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "Swap", "multi_swap_top", shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'swap_zorder',
                label="Swap Z-order",
                show_label=True,
                screentip="Swap Z-order position",
                supertip="Swap the Z-order position of the selected shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "Swap", "multi_swap_zorder", shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                id = 'swap_format',
                label="Swap formatting",
                show_label=True,
                #image_mso='MailMergeMatchFields',
                screentip="Swap shape formatting",
                supertip="Swap the formatting (color, border, font, ...) of the selected shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "Swap", "multi_swap_format", shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'replace_keep_size',
                label="Replace and keep size",
                show_label=True,
                #size='large',
                image='replace_keep_size',
                supertip="The first selected shape replaces all other selected shapes, keeping the position and size of the replaced shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "Swap", "replace_keep_size", shapes=True, shapes_min=2),
                get_enabled = bkt.get_enabled_auto,
            ),
        ])
    ]
)

equal_height_button = bkt.ribbon.SplitButtonFixed(
    show_label=False,
    get_enabled=bkt.apps.ppt_shapes_min2_selected,
    children=[
        bkt.ribbon.Button(
            id = 'equal_height',
            label="Equal height",
            image_mso='SizeToTallest',
            screentip="Equal height as tallest shape",
            supertip="Normalize the height of the selected shapes to match the height of the tallest shape.\nWith the Shift key held, to the shortest shape.\nWith the Ctrl key held, to the last selected shape.",
            on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_height", shapes=True, shapes_min=2),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Menu(label="Equal height menu", supertip="Set multiple shapes to the same height in various ways", children=[
            bkt.ribbon.MenuSeparator(title="Align to shape selection"),
            bkt.ribbon.Button(
                id = 'equal_height2',
                label="Equal height as tallest shape",
                image_mso='SizeToTallest',
                supertip="Normalize the height of the selected shapes to match the height of the tallest shape.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_height_control", shapes=True, shapes_min=2, current_control=True),
                tag="max",
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                id = 'equal_height_min',
                label="Equal height as shortest shape [Shift]",
                show_label=True,
                image_mso='SizeToShortest',
                supertip="Normalize the height of the selected shapes to match the height of the shortest shape.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_height_control", shapes=True, shapes_min=2, current_control=True),
                tag="min",
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_height_median',
                label="Equal height as the median of the shape heights",
                show_label=True,
                #image_mso='ShapeHeight',
                supertip="Normalize the height of the selected shapes to match the median of the heights of the shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_height_control", shapes=True, shapes_min=2, current_control=True),
                tag="median",
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_height_mean',
                label="Equal height as the average of the shape heights",
                show_label=True,
                #image_mso='ShapeHeight',
                supertip="Normalize the height of the selected shapes to match the average of the heights of the shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_height_control", shapes=True, shapes_min=2, current_control=True),
                tag="mean",
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_height_last_sel',
                label="Equal height as reference shape [Ctrl]",
                show_label=True,
                #image_mso='ShapeWidth',
                supertip="Normalize the height of the selected shapes to match the height of the reference shape (i.e. the last or first selected shape).",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_height_master", shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
        ])
    ]
)

equal_width_button = bkt.ribbon.SplitButtonFixed(
    show_label=False,
    get_enabled=bkt.apps.ppt_shapes_min2_selected,
    children=[
        bkt.ribbon.Button(
            id = 'equal_width',
            label="Equal width",
            image_mso='SizeToWidest',
            screentip="Equal width as widest shape",
            supertip="Normalize the width of the selected shapes to match the width of the widest shape.\nWith the Shift key held, to the narrowest shape.\nWith the Ctrl key held, to the last selected shape.",
            on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_width", shapes=True, shapes_min=2),
            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Menu(label="Equal width menu", supertip="Set multiple shapes to the same width in various ways", children=[
            bkt.ribbon.MenuSeparator(title="Align to shape selection"),
            bkt.ribbon.Button(
                id = 'equal_width2',
                label="Equal width as widest shape",
                image_mso='SizeToWidest',
                supertip="Normalize the width of the selected shapes to match the width of the widest shape.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_width_control", shapes=True, shapes_min=2, current_control=True),
                tag="max",
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                id = 'equal_width_min',
                label="Equal width as narrowest shape [Shift]",
                show_label=True,
                image_mso='SizeToNarrowest',
                supertip="Normalize the width of the selected shapes to match the width of the narrowest shape.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_width_control", shapes=True, shapes_min=2, current_control=True),
                tag="min",
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_width_median',
                label="Equal width as the median of the shape widths",
                show_label=True,
                #image_mso='ShapeWidth',
                supertip="Normalize the width of the selected shapes to match the median of the widths of the shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_width_control", shapes=True, shapes_min=2, current_control=True),
                tag="median",
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_width_mean',
                label="Equal width as the average of the shape widths",
                show_label=True,
                #image_mso='ShapeWidth',
                supertip="Normalize the width of the selected shapes to match the average of the widths of the shapes.",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_width_control", shapes=True, shapes_min=2, current_control=True),
                tag="mean",
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
            bkt.ribbon.Button(
                id = 'equal_width_last_sel',
                label="Equal width as reference shape [Ctrl]",
                show_label=True,
                #image_mso='ShapeWidth',
                supertip="Normalize the width of the selected shapes to match the width of the reference shape (i.e. the last or first selected shape).",
                on_action=bkt.CallbackLazy("toolbox.models.arrange", "EqualSize", "equal_width_master", shapes=True, shapes_min=2),
                # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
            ),
        ])
    ]
)



class ShapeDistance(object):
    default_sep = 0.2
    vertical_edges   = bkt.settings.get("toolbox.shapedis.vertical_edges",   "distance") #other options: visual, top, center, bottom
    horizontal_edges = bkt.settings.get("toolbox.shapedis.horizontal_edges", "distance") #other options: visual, left, center, right
    vertical_fix   = bkt.settings.get("toolbox.shapedis.vertical_fix",   "top") #other options: bottom
    horizontal_fix = bkt.settings.get("toolbox.shapedis.horizontal_fix", "left") #other options: right

    @classmethod
    def change_settings(cls, name, value):
        # change setting and save value
        setattr(cls, name, value)
        bkt.settings["toolbox.shapedis."+name] = value

    @classmethod
    def _get_locpin(cls):
        locpin = pplib.LocPin()
        fix_v, fix_h = 1, 1 #visual and distance will use this locpin, but they will use x/x1/y/y1 properties anyway

        if cls.vertical_edges == "center":
            fix_v = 2
        elif cls.vertical_edges == "bottom":
            fix_v = 3
        
        if cls.horizontal_edges == "center":
            fix_h = 2
        elif cls.horizontal_edges == "right":
            fix_h = 3
        
        locpin.fixation = (fix_v, fix_h)
        return locpin

    @classmethod
    def get_enabled_min2_group(cls, shapes):
        if len(shapes) > 1:
            return True
        try:
            #test if shape is a group
            shapes[0].GroupItems
            return True
        except:
            return False
    
    @classmethod
    def get_wrapped_shapes(cls, shapes):
        # if only 1 shape is provided, it must be a group (secured by get_enabled)
        if len(shapes) == 1:
            try:
                shapes = list(iter(shapes[0].GroupItems))
            except:
                raise bkt.context.InappropriateContextError
        return pplib.wrap_shapes(shapes, cls._get_locpin())
    
    @classmethod
    def get_vertical_fix(cls):
        if bkt.get_key_state(bkt.KeyCodes.ALT):
            if cls.vertical_fix == "bottom":
                return "top"
            else:
                return "bottom"
        else:
            return cls.vertical_fix
        
    @classmethod
    def set_shape_sep_vertical(cls, shapes, value):
        try:
            shapes = cls.get_wrapped_shapes(shapes)
        except:
            return

        vertical_fix = cls.get_vertical_fix()
        if vertical_fix=="bottom":
            shapes.sort(key=lambda shape: (shape.y1, shape.x1), reverse=True)
        else:
            shapes.sort(key=lambda shape: (shape.y, shape.x))

        if cls.vertical_edges == "visual" and vertical_fix=="bottom":
            cur_y = shapes[0].visual_y1
        elif cls.vertical_edges == "visual": #vertical_fix=top
            cur_y = shapes[0].visual_y
        
        elif cls.vertical_edges == "distance" and vertical_fix=="bottom":
            cur_y = shapes[0].y1
        elif cls.vertical_edges == "distance": #vertical_fix=top
            cur_y = shapes[0].y

        else:
            cur_y = shapes[0].top
        
        for shape in shapes:
            if cls.vertical_edges == "visual":
                if vertical_fix=="bottom":
                    shape.visual_y1 = cur_y
                else:
                    shape.visual_y = cur_y
                delta = shape.visual_height + value
            elif cls.vertical_edges == "distance":
                if vertical_fix=="bottom":
                    shape.y1 = cur_y
                else:
                    shape.y = cur_y
                delta = shape.height + value
            else:
                shape.top = cur_y
                delta = value
            if vertical_fix=="bottom":
                cur_y -= delta
            else:
                cur_y += delta

    @classmethod
    def get_shape_sep_vertical(cls, shapes):
        from heapq import nsmallest, nlargest
        try:
            shapes = cls.get_wrapped_shapes(shapes)
        except:
            return
        
        if cls.vertical_fix=="bottom":
            # shapes.sort(key=lambda shape: (shape.y1, shape.x1))
            # shapes = shapes[-2:]
            shapes = nlargest(2, shapes, key=lambda shape: (shape.y1, shape.x1))
            shapes.reverse()
        else:
            # shapes.sort(key=lambda shape: (shape.y, shape.x))
            # shapes = shapes[:2]
            shapes = nsmallest(2, shapes, key=lambda shape: (shape.y, shape.x))
        
        if cls.vertical_edges == "distance":
            dis = shapes[1].y-shapes[0].y1
        elif cls.vertical_edges == "visual":
            dis = shapes[1].visual_y-shapes[0].visual_y1
        else:
            dis = shapes[1].top-shapes[0].top
        
        return dis

    @classmethod
    def get_horizontal_fix(cls):
        if bkt.get_key_state(bkt.KeyCodes.ALT):
            if cls.horizontal_fix == "right":
                return "left"
            else:
                return "right"
        else:
            return cls.horizontal_fix

    @classmethod
    def set_shape_sep_horizontal(cls, shapes, value):
        try:
            shapes = cls.get_wrapped_shapes(shapes)
        except:
            return

        horizontal_fix = cls.get_horizontal_fix()
        if horizontal_fix=="right":
            shapes.sort(key=lambda shape: (shape.x1, shape.y1), reverse=True)
        else:
            shapes.sort(key=lambda shape: (shape.x, shape.y))

        if cls.horizontal_edges == "visual" and horizontal_fix=="right":
            cur_x = shapes[0].visual_x1
        elif cls.horizontal_edges == "visual": #horizontal_fix=left
            cur_x = shapes[0].visual_x
        
        elif cls.horizontal_edges == "distance" and horizontal_fix=="right":
            cur_x = shapes[0].x1
        elif cls.horizontal_edges == "distance": #horizontal_fix=left
            cur_x = shapes[0].x

        else:
            cur_x = shapes[0].left
        
        for shape in shapes:
            if cls.horizontal_edges == "visual":
                if horizontal_fix=="right":
                    shape.visual_x1 = cur_x
                else:
                    shape.visual_x = cur_x
                delta = shape.visual_width + value
            elif cls.horizontal_edges == "distance":
                if horizontal_fix=="right":
                    shape.x1 = cur_x
                else:
                    shape.x = cur_x
                delta = shape.width + value
            else:
                shape.left = cur_x
                delta = value
            if horizontal_fix=="right":
                cur_x -= delta
            else:
                cur_x += delta

    @classmethod
    def get_shape_sep_horizontal(cls, shapes):
        from heapq import nsmallest, nlargest
        try:
            shapes = cls.get_wrapped_shapes(shapes)
        except:
            return
        
        if cls.horizontal_fix=="right":
            # shapes.sort(key=lambda shape: (shape.x1, shape.y1))
            # shapes = shapes[-2:]
            shapes = nlargest(2, shapes, key=lambda shape: (shape.x1, shape.y1))
            shapes.reverse()
        else:
            # shapes.sort(key=lambda shape: (shape.x, shape.y))
            # shapes = shapes[:2]
            shapes = nsmallest(2, shapes, key=lambda shape: (shape.x, shape.y))

        if cls.horizontal_edges == "distance":
            dis = shapes[1].x-shapes[0].x1
        elif cls.horizontal_edges == "visual":
            dis = shapes[1].visual_x-shapes[0].visual_x1
        else:
            dis = shapes[1].left-shapes[0].left
        
        return dis


    @classmethod
    def set_shape_sep_vertical_zero(cls, shapes):
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            cls.set_shape_sep_vertical(shapes, cm_to_pt(cls.default_sep))
        elif bkt.get_key_state(bkt.KeyCodes.CTRL):
            cls.set_shape_sep_vertical(shapes, cls.get_shape_sep_vertical(shapes))
        else:
            cls.set_shape_sep_vertical(shapes, 0)

    @classmethod
    def set_shape_sep_horizontal_zero(cls, shapes):
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            cls.set_shape_sep_horizontal(shapes, cm_to_pt(cls.default_sep))
        elif bkt.get_key_state(bkt.KeyCodes.CTRL):
            cls.set_shape_sep_horizontal(shapes, cls.get_shape_sep_horizontal(shapes))
        else:
            cls.set_shape_sep_horizontal(shapes, 0)


    ### Euclidian distance and angle methods ###

    @classmethod
    def is_mode_centric(cls):
        return cls.euclid_multi_shape_mode == "centric"

    @classmethod
    def is_mode_delta(cls):
        alt = bkt.get_key_state(bkt.KeyCodes.ALT)
        return alt or cls.euclid_multi_shape_mode == "delta"

    @classmethod
    def is_mode_distribute(cls):
        return cls.euclid_multi_shape_mode == "distribute"


class ShapeRotation(object):
    locpin = LocPinCollection.rotation

    @staticmethod
    def get_enabled(selection):
        try:
            return selection.Type in [2,3] and not selection.ShapeRange.HasTable in [-2, -1]
            #tables do not support rotation
        except:
            return False

    @classmethod
    def set_rotation(cls, shapes, value):
        shapes = pplib.wrap_shapes(shapes, cls.locpin)
        bkt.apply_delta_on_ALT_key(
            lambda shape, value: setattr(shape, 'rotation', value), 
            lambda shape: shape.rotation, 
            shapes, value)

    @classmethod
    def get_rotation(cls, shapes):
        shape = pplib.wrap_shape(shapes[0], cls.locpin)
        return shape.rotation

    @classmethod
    def set_rotation_zero(cls, shapes):
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            cls.set_rotation(shapes, 180)
        elif bkt.get_key_state(bkt.KeyCodes.CTRL):
            cls.set_rotation(shapes, cls.get_rotation(shapes))
        else:
            cls.set_rotation(shapes, 0)

    @staticmethod
    def get_pressed_flipv(shapes):
        return shapes[0].VerticalFlip == -1

    @staticmethod
    def set_flipv(pressed, shapes):
        pressed = -1 if pressed else 0
        for shape in shapes:
            if shape.VerticalFlip != pressed:
                shape.Flip(1) #msoFlipVertical

    @staticmethod
    def get_pressed_fliph(shapes):
        return shapes[0].HorizontalFlip == -1

    @staticmethod
    def set_fliph(pressed, shapes):
        pressed = -1 if pressed else 0
        for shape in shapes:
            if shape.HorizontalFlip != pressed:
                shape.Flip(0) #msoFlipHorizontal


class ShapeEuclid(object):
    ### only for euclid distance and angle:
    shape1_index  = 0 #center-shape-index is either 0 for first selected shape or -1 for last selected shape
    # shape1_locpin = pplib.LocPin(4) #center point as initial locpin
    # shape2_locpin = pplib.LocPin(4) #center point as initial locpin
    shape_rotate_with_angle = False #rotate shape if angle is changed
    euclid_multi_shape_mode = "centric" #Options: centric, delta, distribute

    @classmethod
    def get_shape_sep_euclid(cls, shapes):
        '''
        get euclidian distance from center shape to second shape
        '''
        shape1 = pplib.wrap_shape(shapes[cls.shape1_index], LocPinCollection.dis1)
        shape2 = pplib.wrap_shape(shapes[cls.shape1_index+1], LocPinCollection.dis2)
        shape1_x, shape1_y = shape1.left, shape1.top
        shape2_x, shape2_y = shape2.left, shape2.top

        # return math.sqrt( (shape2_y-shape1_y)**2 + (shape2_x-shape1_x)**2 )
        return math.hypot(shape2_x-shape1_x, shape2_y-shape1_y)
    
    @classmethod
    def set_shape_sep_euclid(cls, shapes, value):
        '''
        set euclidian distance from center shape to all other shapes
        '''
        def _get_current_distance(shape1, shape2):
            vector = [shape2.left-shape1.left, shape2.top-shape1.top]
            return math.hypot( *vector )

        def _get_new_shape_coords(shape1, shape2, delta_distance):
            vector = [shape2.left-shape1.left, shape2.top-shape1.top]
            cur_dis = math.hypot( *vector )
            if cur_dis == 0:
                raise ValueError("current distance is 0")
            uni_vector = [vector[0]/cur_dis, vector[1]/cur_dis]
            distance = cur_dis + delta_distance
            new_vector = [uni_vector[0]*distance, uni_vector[1]*distance]
            shape_rotation = (360-round(-180/math.pi * math.atan2(new_vector[1], new_vector[0]), 1)) % 360
            return new_vector, shape_rotation


        shape1 = pplib.wrap_shape(shapes[cls.shape1_index], LocPinCollection.dis1)
        shape1_x, shape1_y = shape1.left, shape1.top

        shapes = pplib.wrap_shapes(shapes[cls.shape1_index+1:], LocPinCollection.dis2)
        # shape2_x, shape2_y = shapes[0].left, shapes[0].top

        # alt = bkt.get_key_state(bkt.KeyCodes.ALT)

        # if cls.is_mode_centric() or cls.is_mode_distribute():
        if not cls.is_mode_delta():
            for i, shape in enumerate(shapes):
                try:
                    if cls.is_mode_centric():
                        delta_distance = value-_get_current_distance(shape1, shape)
                    else: #is_mode_distribute
                        delta_distance = (i+1)*value-_get_current_distance(shape1, shape)
                    new_vector, shape_rotation = _get_new_shape_coords(shape1, shape, delta_distance)
                except ValueError:
                    continue
                shape.left = shape1_x+new_vector[0]
                shape.top  = shape1_y+new_vector[1]
                # set shape rotation (without using wrapper function)
                if cls.shape_rotate_with_angle:
                    shape.shape.rotation = shape_rotation
        
        else: #is_mode_delta
            delta_distance = value-_get_current_distance(shape1, shapes[0])
            for shape in shapes:
                try:
                    new_vector, shape_rotation = _get_new_shape_coords(shape1, shape, delta_distance)
                except ValueError:
                    continue
                shape.left = shape1_x+new_vector[0]
                shape.top  = shape1_y+new_vector[1]
                # set shape rotation (without using wrapper function)
                if cls.shape_rotate_with_angle:
                    shape.shape.rotation = shape_rotation

    @classmethod
    def get_shape_angle(cls, shapes):
        '''
        get euclidian angle from center shape to second shaope
        '''
        shape1 = pplib.wrap_shape(shapes[cls.shape1_index], LocPinCollection.dis1)
        shape2 = pplib.wrap_shape(shapes[cls.shape1_index+1], LocPinCollection.dis2)
        shape1_x, shape1_y = shape1.left, shape1.top
        shape2_x, shape2_y = shape2.left, shape2.top

        vector = [shape2_x-shape1_x, shape2_y-shape1_y]
        return round(-180/math.pi * math.atan2(vector[1], vector[0]),1)

    @classmethod
    def set_shape_angle(cls, shapes, value):
        '''
        set euclidian angle from center shape to all other shapes
        '''
        def _get_current_angle(shape1, shape2):
            vector = [shape2.left-shape1.left, shape2.top-shape1.top]
            return round(-180/math.pi * math.atan2(vector[1], vector[0]), 1)

        def _get_new_shape_coords(shape1, shape2, delta_angle):
            vector = [shape2.left-shape1.left, shape2.top-shape1.top]
            new_vector = algos.rotate_point(vector[0], vector[1], 0, 0, delta_angle)
            shape_rotation = (360-round(-180/math.pi * math.atan2(new_vector[1], new_vector[0]), 1)) % 360
            return new_vector, shape_rotation


        shape1 = pplib.wrap_shape(shapes[cls.shape1_index], LocPinCollection.dis1)
        shape1_x, shape1_y = shape1.left, shape1.top

        shapes = pplib.wrap_shapes(shapes[cls.shape1_index+1:], LocPinCollection.dis2)
        # shape_rotation = (360-value) % 360

        # alt = bkt.get_key_state(bkt.KeyCodes.ALT)

        # if cls.is_mode_centric() or cls.is_mode_distribute():
        if not cls.is_mode_delta():
            for i, shape in enumerate(shapes):
                if cls.is_mode_centric():
                    delta_angle = -(_get_current_angle(shape1, shape) - value)
                else: #is_mode_distribute
                    delta_angle = -(_get_current_angle(shape1, shape) - value*(i+1))
                new_vector, shape_rotation = _get_new_shape_coords(shape1, shape, delta_angle)
                # shape2_x, shape2_y = shape.left, shape.top

                # vector = [shape2_x-shape1_x, shape2_y-shape1_y]
                # cur_angle = round(-180/math.pi * math.atan2(vector[1], vector[0]), 1)
                # new_vector = algos.rotate_point(vector[0], vector[1], 0, 0, -(cur_angle-value))

                shape.left = shape1_x+new_vector[0]
                shape.top  = shape1_y+new_vector[1]
                # set shape rotation (without using wrapper function)
                if cls.shape_rotate_with_angle:
                    shape.shape.rotation = shape_rotation
        
        else: #is_mode_delta
            delta_angle = _get_current_angle(shape1, shapes[0])-value
            for shape in shapes:
                new_vector, shape_rotation = _get_new_shape_coords(shape1, shape, -delta_angle)

                shape.left = shape1_x+new_vector[0]
                shape.top  = shape1_y+new_vector[1]
                # set shape rotation (without using wrapper function)
                if cls.shape_rotate_with_angle:
                    shape.shape.rotation = shape_rotation


class ArrangeAdvanced(object):
    #class variables:
    # FIXME: master should be an instance-variable, other classes should get an ArrangeAdvanced-instance by dependency injection
    # master = "LAST"
    # fallback_first_last = "LAST"
    # master_indicator = True

    #instance variables:
    # margin = 0
    # resize = False
    # ref_shape = None
    # ref_frame = None
    

    # def _create_control(self, control, children, callbacks):
    #     control.children = [
    #         children['arrange_left_at_left'],
    #         children['arrange_middle_at_left'],
    #         children['arrange_right_at_left'],
    #         children['arrange_left_at_middle'],
    #         children['arrange_middle_at_middle'],
    #         children['arrange_right_at_middle'],
    #         children['arrange_left_at_right'],
    #         children['arrange_middle_at_right'],
    #         children['arrange_right_at_right'],
    #
    #         children['arrange_top_at_top'],
    #         children['arrange_middle_at_top'],
    #         children['arrange_bottom_at_top'],
    #         children['arrange_top_at_middle'],
    #         children['arrange_vmiddle_at_vmiddle'],
    #         children['arrange_bottom_at_middle'],
    #         children['arrange_top_at_bottom'],
    #         children['arrange_middle_at_bottom'],
    #         children['arrange_bottom_at_bottom'],
    #
    #         children['set_margin'],
    #         bkt.ribbon.Menu(label="Options", children=[
    #             children['set_moving'],
    #             children['set_resizing']
    #         ]),
    #         bkt.ribbon.SplitButton(children=[
    #             bkt.ribbon.ToggleButton(image_mso="PositionAbsoluteMarks"),
    #             bkt.ribbon.Menu(label='Align to', children=[
    #                 #bkt.ribbon.MenuSeparator(title='Shape from selection'),
    #                 children['set_master_first'],
    #                 children['set_master_last'],
    #                 #bkt.ribbon.MenuSeparator(title='Specified shape'),
    #                 bkt.ribbon.MenuSeparator(),
    #                 children['set_master_specified'],
    #                 #children['set_master_slide'],
    #                 children['specify_shape']
    #             ])
    #         ])
    #     ]
    #     return control

    def __init__(self):
        #instance variables:
        self.margin = 0
        self.resize = False
        self.ref_shape = None
        self.ref_frame = None
        
        #save preference for master shape (first or last selected) into class variables
        # ArrangeAdvanced.fallback_first_last = bkt.settings.get("arrange_advanced.default", "LAST")
        # ArrangeAdvanced.master = ArrangeAdvanced.fallback_first_last
        #save preference to show master shape indicator into class variables
        # ArrangeAdvanced.master_indicator = bkt.settings.get("arrange_advanced.master_indicator", True)
        # if ArrangeAdvanced.master_indicator:
        #     self._register_dialog()

        # self.position_gallery = pplib.PositionGallery(
        #     label="Choose custom area",
        #     on_position_change = bkt.Callback(self.specify_master_customarea),
        #     get_item_supertip = bkt.Callback(self.get_item_supertip)
        # )

    @property
    def master_handler(self):
        return GlobalMasterShape
    
    def get_item_supertip(self, index):
        return 'Verwende angezeigten Position/Größe als Master.'
    

    ### arrange at master's left side ###

    def arrange_left_at_left(self, shapes):
        ''' arrange left side of shapes at masters's left side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(shape) + self.get_shape_width(shape) - ( self.get_shape_left(master) + self.margin )
                if self.use_resizing() and new_size>=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left( shape, self.get_shape_left(master) + self.margin )


    def arrange_middle_at_left(self, shapes):
        ''' arrange midpoint of shapes horizontally at masters's left side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                offset = self.get_shape_left(master) - self.margin
                max_distance = max( abs(self.get_shape_left(shape) + self.get_shape_width(shape) - offset), abs(self.get_shape_left(shape) - offset )    )
                new_size = 2*max_distance
                #new_size = 2*(self.get_shape_left(shape) + self.get_shape_width(shape) - self.get_shape_left(master) + self.margin)
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) - self.get_shape_width(shape)/2 - self.margin)


    def arrange_right_at_left(self, shapes):
        ''' arrange right side of shapes at masters's left side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(master) - self.get_shape_left(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) - self.get_shape_width(shape) - self.margin)


    ### arrange at master's middle ###

    def arrange_left_at_middle(self, shapes):
        ''' arrange left side of shapes horizontally at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(shape) + self.get_shape_width(shape) - self.get_shape_left(master) - self.get_shape_width(master)/2 - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master)/2 + self.margin)


    def arrange_middle_at_middle(self, shapes):
        ''' arrange midpoint of shapes horizontally at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                offset = self.get_shape_left(master) + self.get_shape_width(master)/2
                max_distance = max( abs(self.get_shape_left(shape) + self.get_shape_width(shape) - offset), abs(self.get_shape_left(shape) - offset )    )
                new_size = 2*max_distance
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master)/2 - self.get_shape_width(shape)/2)


    def arrange_right_at_middle(self, shapes):
        ''' arrange right side of shapes horizontally at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(master) + self.get_shape_width(master)/2 - self.get_shape_left(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master)/2 - self.get_shape_width(shape) - self.margin)


    ### arrange at master's right side ###

    def arrange_left_at_right(self, shapes):
        ''' arrange left side of shapes at masters's right side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(shape) + self.get_shape_width(shape) - self.get_shape_left(master) - self.get_shape_width(master) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master) + self.margin)


    def arrange_middle_at_right(self, shapes):
        ''' arrange midpoint of shapes horizontally at masters's right side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                offset = self.get_shape_left(master) + self.get_shape_width(master) + self.margin
                max_distance = max( abs(self.get_shape_left(shape) + self.get_shape_width(shape) - offset), abs(self.get_shape_left(shape) - offset )    )
                new_size = 2*max_distance
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master) - self.get_shape_width(shape)/2 + self.margin)


    def arrange_right_at_right(self, shapes):
        ''' arrange right side of shapes at masters's right side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_left(master) + self.get_shape_width(master) - self.get_shape_left(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_width(shape, new_size)
                self.set_shape_left(shape, self.get_shape_left(master) + self.get_shape_width(master) - self.get_shape_width(shape) - self.margin)


    ############

    ### arrange at master's top side ###

    def arrange_top_at_top(self, shapes):
        ''' arrange top side of shapes at masters's top side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(shape) + self.get_shape_height(shape) - ( self.get_shape_top(master) + self.margin )
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.margin)


    def arrange_middle_at_top(self, shapes):
        ''' arrange midpoint of shapes at masters's top side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = 2*(self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) + self.margin)
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) - self.get_shape_height(shape)/2 - self.margin)


    def arrange_bottom_at_top(self, shapes):
        ''' arrange bottom side of shapes at masters's top side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(master) - self.get_shape_top(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) - self.get_shape_height(shape) - self.margin)


    ### arrange at master's vertical middle ###

    def arrange_top_at_middle(self, shapes):
        ''' arrange top side of shapes vertically at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) - self.get_shape_height(master)/2 - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master)/2 + self.margin)


    def arrange_vmiddle_at_vmiddle(self, shapes):
        ''' arrange midpoint of shapes vertically at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = 2*(self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) - self.get_shape_height(master)/2)
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master)/2 - self.get_shape_height(shape)/2)


    def arrange_bottom_at_middle(self, shapes):
        ''' arrange bottom side of shapes vertically at masters's midpoint '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(master) + self.get_shape_height(master)/2 - self.get_shape_top(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master)/2 - self.get_shape_height(shape) - self.margin)


    ### arrange at master's right side ###

    def arrange_top_at_bottom(self, shapes):
        ''' arrange top side of shapes at masters's bottom side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) - self.get_shape_height(master) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master) + self.margin)


    def arrange_middle_at_bottom(self, shapes):
        ''' arrange midpoint of shapes vertically at masters's bottom side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = 2*(self.get_shape_top(shape) + self.get_shape_height(shape) - self.get_shape_top(master) - self.get_shape_height(master)-self.margin)
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master) - self.get_shape_height(shape)/2 + self.margin)


    def arrange_bottom_at_bottom(self, shapes):
        ''' arrange bottom side of shapes at masters's bottom side '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                new_size = self.get_shape_top(master) + self.get_shape_height(master) - self.get_shape_top(shape) - self.margin
                if self.use_resizing() and new_size >=0:
                    self.set_shape_height(shape, new_size)
                self.set_shape_top(shape, self.get_shape_top(master) + self.get_shape_height(master) - self.get_shape_height(shape) - self.margin)


    ### quick arrange ###

    def arrange_quick_position(self, shapes):
        ''' position shapes at master's position '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                self.set_shape_left( shape, self.get_shape_left(master) )
                self.set_shape_top( shape, self.get_shape_top(master) )

    def arrange_quick_size(self, shapes):
        ''' change size of shapes according to master's size '''
        master = self.get_master_from_shapes(shapes)
        for shape in shapes:
            if shape != master:
                self.set_shape_width( shape, self.get_shape_width(master) )
                self.set_shape_height( shape, self.get_shape_height(master) )


    # def enabled(self, shapes):
    #     #return (self.master in ["FIXED-SHAPE", "FIXED-SLIDE", "FIXED-CONTENTAREA", "FIXED-CUSTOMAREA"] and len(shapes) > 0) or len(shapes) > 1
    #     return len(shapes) > 0


    ### detect master shape ###

    def get_master_from_shapes(self, shapes):
        def _test_ref(ref):
            try:
                ref.left #test if ref still exists
                return True
            except:
                return False
        # def _test_ref_or_use_fallback(ref):
        #     try:
        #         ref.left #test if ref still exists
        #         return ref
        #     except:
        #         bkt.message.warning("Fehler: Referenz wurde nicht gefunden. Master wird zurückgesetzt.")
        #         ArrangeAdvanced.master = self.fallback_first_last
        #         return self.get_master_from_shapes(shapes)

        ''' obtain master shape from given shapes according to master-setting '''
        if bkt.get_key_state(bkt.KeyCodes.CTRL):
            return pplib.BoundingFrame(shapes[0].parent, contentarea=True)
        
        elif self.master_handler.master == "FIXED-SHAPE" and self.ref_shape and self.ref_frame:
            #if ref_shape was deleted, use ref_frame instead
            if _test_ref(self.ref_shape):
                return self.ref_shape
            else:
                return self.ref_frame
        elif self.master_handler.master == "FIXED-SLIDE" and self.ref_frame:
            return self.ref_frame
        elif self.master_handler.master == "FIXED-CONTENTAREA" and self.ref_frame:
            return self.ref_frame
        elif self.master_handler.master == "FIXED-CUSTOMAREA" and self.ref_frame:
            return self.ref_frame
            
        elif len(shapes) == 1:
            ## fallback if only one shape in selection
            return pplib.BoundingFrame(shapes[0].parent, contentarea=True)

        elif self.master_handler.master == "PPTDEFAULT":
            return pplib.BoundingFrame.from_shapes(shapes)
            
        elif self.master_handler.master == "FIRST":
            return shapes[0]
        else:
            # default: master == "LAST"
            return shapes[-1]
    
    
    ### configure margin, resizing-option ###

    def set_margin(self, value, context):
        ''' callback to set margin-value '''
        self.margin = value

    def get_margin(self):
        ''' returns current margin-value '''
        return self.margin

    def set_moving(self, pressed):
        ''' callback to switch between moving/resizing-option '''
        self.resize=(pressed==False)

    def get_moving(self):
        ''' returns if moving-option is used '''
        return (self.resize==False)

    def set_resizing(self, pressed):
        ''' callback to switch between resizing/moving-option '''
        self.resize=(pressed==True)

    def get_resizing(self):
        ''' returns if resizing-option is set '''
        return (self.resize==True)
    
    def use_resizing(self):
        ''' returns if resizing-option should be used. Depends on setting for resize-option, and SHIFT-key-state '''
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            return True
        else:
            return self.get_resizing()
    
    
    ### configure master in selection ###

    def set_master_first(self, pressed):
        ''' callback: set master-setting to use first shape in selection '''
        self.master_handler.fallback_first_last="FIRST"
        self.master_handler.master="FIRST"
        bkt.settings["arrange_advanced.default"] = "FIRST"

    def get_master_first(self):
        ''' returns whether master-setting is set to first shape in selection '''
        return (self.master_handler.master=="FIRST")


    def set_master_last(self, pressed):
        ''' callback: set master-setting to use last shape in selection '''
        self.master_handler.fallback_first_last="LAST"
        self.master_handler.master="LAST"
        bkt.settings["arrange_advanced.default"] = "LAST"

    def get_master_last(self):
        ''' returns whether master-setting is set to last shape in selection '''
        return (self.master_handler.master=="LAST")


    def set_master_pptdefault(self, pressed):
        ''' callback: set master-setting to use outermost shape in selection (as powerpoint default aligning) '''
        self.master_handler.fallback_first_last="PPTDEFAULT"
        self.master_handler.master="PPTDEFAULT"
        bkt.settings["arrange_advanced.default"] = "PPTDEFAULT"

    def get_master_pptdefault(self):
        ''' returns whether master-setting is set to outermost shape in selection (as powerpoint default aligning) '''
        return (self.master_handler.master=="PPTDEFAULT")


    # def set_master_indicator(self, pressed):
    #     ''' callback: set whether master shape indicator is shown '''
    #     if not pressed:
    #         ArrangeAdvanced.master_indicator = False
    #         bkt.settings["arrange_advanced.master_indicator"] = False
    #         self._unregister_dialog()
    #     else:
    #         ArrangeAdvanced.master_indicator = True
    #         bkt.settings["arrange_advanced.master_indicator"] = True
    #         self._register_dialog()

    # def get_master_indicator(self):
    #     ''' returns whether master shape indicator is shown '''
    #     return self.master_indicator

    # @classmethod
    # def _register_dialog(cls):
    #     # register dialog
    #     bkt.powerpoint.context_dialogs.register_dialog(
    #         MasterShapeDialog(cls)
    #     )
    
    # @classmethod
    # def _unregister_dialog(cls):
    #     # unregister dialog
    #     bkt.powerpoint.context_dialogs.unregister("MASTER")

    # @classmethod
    # def get_master_for_indicator(cls, shapes):
    #     if cls.master=="FIRST":
    #         return shapes[0]
    #     elif cls.master=="LAST":
    #         return shapes[-1]
    #     else:
    #         return None



    ### configure fixed master-shape ###
    
    def specify_shape(self, shape, pressed):
        ''' callback: set master to selected shape. On deactivation fallback to master in selection '''
        if pressed:
            self.ref_shape = shape
            self.ref_frame = pplib.BoundingFrame.from_shape(shape) #set ref_frame as fallback if shape is deleted
            self.master_handler.master = "FIXED-SHAPE"
        else:
            self.master_handler.master = self.master_handler.fallback_first_last

    def specify_master_slide(self, slide, pressed):
        ''' callback: set master to slide. On deactivation fallback to master in selection '''
        if pressed:
            self.ref_frame = pplib.BoundingFrame(slide)
            self.master_handler.master = "FIXED-SLIDE"
        else:
            self.master_handler.master = self.master_handler.fallback_first_last

    def specify_master_contentarea(self, slide, pressed):
        ''' callback: set master to content-area of slide. On deactivation fallback to master in selection '''
        if pressed:
            self.ref_frame = pplib.BoundingFrame(slide, contentarea=True)
            self.master_handler.master = "FIXED-CONTENTAREA"
        else:
            self.master_handler.master = self.master_handler.fallback_first_last

    def specify_master_customarea(self, target_frame):
        ''' set master to given target-frame '''
        self.master_handler.master = "FIXED-CUSTOMAREA"
        self.ref_frame = target_frame

    def specify_master_customarea_toggle(self, pressed, context):
        ''' set master to previous target-frame if defined, otherwise defined content area, otherwise set userdefined by shapes '''
        if pressed:
            self.specify_master_customarea(self.ref_frame)
            # if self.ref_frame:
            #     frame = self.ref_frame
            # elif pplib.ContentArea.isset_contentarea(context.presentation):
            #     frame = pplib.BoundingFrame.from_rect(*pplib.ContentArea.read_contentarea(context.presentation))
            # else:
            #     shapes = context.shapes
            #     if len(shapes) > 0:
            #         frame = pplib.BoundingFrame.from_shapes(shapes)
            #         pplib.ContentArea.define_contentarea(context.presentation, frame)
            #     else:
            #         frame = pplib.BoundingFrame(context.slide, contentarea=True)
            # self.specify_master_customarea(frame)
        else:
            self.master_handler.master = self.master_handler.fallback_first_last

    def specify_wiz(self, pressed, context):
        ''' callback for master-wizzard: sets master to selected shape, content-area or master in selection - 
            depending on given shape-selection and current master-setting
        '''
        if not pressed:
            self.master_handler.master = self.master_handler.fallback_first_last
            
        else:
            # pressed == True
            shapes = self.get_shapes_from_context(context)
            
            if len(shapes) == 0:
                # use contentarea as default
                self.specify_master_contentarea(context.app.activewindow.selection.SlideRange.Item(1), pressed=True)
            
            else:
                # use shape
                self.specify_shape(shapes[0], pressed=True)
                # FIXME: message if multiple shapes are selected?
    
    
    def is_slide_or_shape_specified(self):
        ''' returns whether a fixed master is used '''
        return self.master_handler.master in ["FIXED-SHAPE", "FIXED-SLIDE", "FIXED-CONTENTAREA", "FIXED-CUSTOMAREA"]
        
    def is_shape_specified(self):
        ''' returns whether master is a fixed shape '''
        return self.master_handler.master == "FIXED-SHAPE"
    
    def is_slide_specified(self):
        ''' returns whether master is set to slide '''
        return self.master_handler.master == "FIXED-SLIDE"

    def is_contentarea_specified(self):
        ''' returns whether master is set to slide content area '''
        return self.master_handler.master == "FIXED-CONTENTAREA"

    def is_customarea_specified(self):
        ''' returns whether master is set to custom area '''
        return self.master_handler.master == "FIXED-CUSTOMAREA"

    def is_customarea_specifiable(self):
        ''' returns whether master is set to custom area '''
        return self.ref_frame is not None
    
    def is_shape_specified_or_shape_specifiable(self, context):
        ''' callback: returns whether master can be changed to fixed shape.
            Can either use a shape from selection or the last fixed shape.
        '''
        if self.is_shape_specified():
            return True
        else:
            shapes = self.get_shapes_from_context(context)
            return len(shapes) > 0
    
    
    
    ### helper function for shapes ###
    
    def get_shapes_from_context(self, context):
        ''' retrieves shapes from context, depending on selection-type (text, shapes, group) '''
        try:
            if context.app.ActiveWindow.ViewType != 9:
               return []

            selection = context.app.ActiveWindow.Selection
        except:
            return []
        
        return pplib.get_shapes_from_selection(selection)

        # # ShapeRange accessible if shape or text selected
        # shapes = []
        # if selection.Type == 2 or selection.Type == 3:
        #     try:
        #         if selection.HasChildShapeRange:
        #             # shape selection inside grouped shapes
        #             shapes = list(iter(selection.ChildShapeRange))
        #         else:
        #             shapes = list(iter(selection.ShapeRange))
        #     except:
        #         shapes = []
        
        # return shapes
        
    

    ### helper functions to compute/change actual left/top/width/height of rotated shapes ###

    ### actual shape's left ###

    def get_shape_left(self, shape):
        ''' return actual shape's left boundary considering the shape's rotation '''
        if isinstance(shape, pplib.BoundingFrame):
            return shape.left
        shapes_left = min( p[0] for p in self.get_shape_bounding_nodes(shape) )
        return shapes_left

    def set_shape_left(self, shape, value):
        ''' set shape's left considering the shape's rotation '''
        delta = shape.left - self.get_shape_left(shape)
        shape.left = value + delta

    ### actual shape's width ###

    def get_shape_width(self, shape):
        ''' return actual shape's width considering the shape's rotation '''
        if isinstance(shape, pplib.BoundingFrame):
            return shape.width
        return max( p[0] for p in self.get_shape_bounding_nodes(shape) ) - self.get_shape_left(shape)

    def set_shape_width(self, shape, value):
        ''' set shape's width considering the shape's rotation '''
        if shape.rotation == 0 or shape.rotation == 180:
            shape.width = value
        elif shape.rotation == 90 or shape.rotation == 270:
            shape.height = value
        else:
            delta = value - self.get_shape_width(shape)
            # delta_vector (delta-width, 0) um shape-rotation drehen
            # delta_vector = self.rotate_point(delta, 0, 0, 0, shape.rotation)
            delta_vector = algos.rotate_point_by_shape_rotation(delta, 0, shape)
            # vorzeichen beibehalten (entweder vergrößern oder verkleinern - nicht beides)
            vorzeichen = 1 if delta > 0 else -1
            delta_vector = [vorzeichen * abs(delta_vector[0]), vorzeichen * abs(delta_vector[1]) ]
            # shape anpassen
            shape.width += delta_vector[0]
            shape.height += delta_vector[1]

    ### actual shape's top ###

    def get_shape_top(self, shape):
        ''' return actual shape's top boundary considering the shape's rotation '''
        if isinstance(shape, pplib.BoundingFrame):
            return shape.top
        return min( p[1] for p in self.get_shape_bounding_nodes(shape) )

    def set_shape_top(self, shape, value):
        ''' set shape's top considering the shape's rotation '''
        delta = shape.top - self.get_shape_top(shape)
        shape.top = value + delta

    ### actual shape's height ###

    def get_shape_height(self, shape):
        ''' return actual shape's height considering the shape's rotation '''
        if isinstance(shape, pplib.BoundingFrame):
            return shape.height
        return max( p[1] for p in self.get_shape_bounding_nodes(shape) ) - self.get_shape_top(shape)

    def set_shape_height(self, shape, value):
        ''' set shape's height considering the shape's rotation '''
        if shape.rotation == 0 or shape.rotation == 180:
            shape.height = value
        elif shape.rotation == 90 or shape.rotation == 270:
            shape.width = value
        else:
            delta = value - self.get_shape_height(shape)
            # delta_vector (0, delta-height) um shape-rotation drehen
            # delta_vector = self.rotate_point(0, delta, 0, 0, shape.rotation)
            delta_vector = algos.rotate_point_by_shape_rotation(0, delta, shape)
            # vorzeichen beibehalten (entweder vergrößern oder verkleinern - nicht beides)
            vorzeichen = 1 if delta > 0 else -1
            delta_vector = [vorzeichen * abs(delta_vector[0]), vorzeichen * abs(delta_vector[1]) ]
            # shape anpassen
            shape.width += delta_vector[0]
            shape.height += delta_vector[1]


    ### helper fuctions to compute rotated points ###

    def get_shape_bounding_nodes(self, shape):
        ''' compute bounding nodes (surrounding-square) for rotated shapes '''
        return algos.get_bounding_nodes(shape)
        # points = [ [ shape.left, shape.top ], [shape.left, shape.top+shape.height], [shape.left+shape.width, shape.top+shape.height], [shape.left+shape.width, shape.top] ]

        # x0 = shape.left + shape.width/2
        # y0 = shape.top + shape.height/2

        # rotated_points = [
        #     list(algos.rotate_point(p[0], p[1], x0, y0, 360-shape.rotation))
        #     for p in points
        # ]
        # return rotated_points



    ### ui generation functions ###
    def get_master_button(self, postfix="", **kwargs):
        return bkt.ribbon.SplitButtonFixed(id="arrange_master_splitbutton"+postfix, children=[
            
            bkt.ribbon.ToggleButton(
                id="arrange_use_master"+postfix,
                label="Reference",
                image_mso="PositionAbsoluteMarks", on_toggle_action=bkt.Callback(self.specify_wiz), get_pressed=bkt.Callback(self.is_slide_or_shape_specified),
                screentip="Toggle alignment to reference (shape/slide)", 
                supertip="When activated, all shapes are then aligned to the selected shape (reference shape). If no shape is selected on activation, alignment is done to the content area."
            ),
            
            bkt.ribbon.Menu(label='Reference menu', screentip='Selection of alignment', supertip="Define which reference object the shapes should be aligned to", children=[
                
                bkt.ribbon.Button(
                    id="arrange_master_auto"+postfix,
                    label="Automatic reference selection",
                    image_mso="PositionAbsoluteMarks", 
                    supertip="Makes the first selected shape the reference shape. If no shape is selected, the content area is used as reference.",
                    on_action=bkt.Callback(lambda context:self.specify_wiz(pressed=True, context=context))
                ),
                
                bkt.ribbon.MenuSeparator(title="Alignment within selection…"),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_first_shape"+postfix,
                    label="…to first shape",  on_toggle_action=bkt.Callback(self.set_master_first), get_pressed=bkt.Callback(self.get_master_first),
                    screentip="Alignment to the first shape within the selection",
                    supertip="Shapes are aligned to the first selected shape in the selection."
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_last_shape"+postfix,
                    label="…to last shape", on_toggle_action=bkt.Callback(self.set_master_last),  get_pressed=bkt.Callback(self.get_master_last),
                    screentip="Alignment to the last shape within the selection",
                    supertip="Shapes are aligned to the last selected shape in the selection."
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_pptdefault_shape"+postfix,
                    label="…to outermost shape (PPT default)", on_toggle_action=bkt.Callback(self.set_master_pptdefault),  get_pressed=bkt.Callback(self.get_master_pptdefault),
                    screentip="Alignment to the outermost shape within the selection",
                    supertip="Shapes are aligned to the outermost selected shape in the selection."
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.ToggleButton(
                    id="arrange_show_master_shape"+postfix,
                    label="Indicator on reference shape", on_toggle_action=bkt.Callback(self.master_handler.set_master_indicator),  get_pressed=bkt.Callback(self.master_handler.get_master_indicator),
                    screentip="Show indicator on the reference shape within the selection",
                    supertip="When at least two shapes are selected, a small indicator with the text 'Reference' is shown at the lower-left corner of the reference shape (first or last)."
                ),
                
                bkt.ribbon.MenuSeparator(title="Alignment to reference"),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_shape"+postfix,
                    label="Shape", on_toggle_action=bkt.Callback(self.specify_shape),         get_pressed=bkt.Callback(self.is_shape_specified), get_enabled=bkt.Callback(self.is_shape_specified_or_shape_specifiable),
                    screentip="Alignment to the selected shape (reference shape)",
                    supertip="The selected shape is set as the reference shape. Shapes are aligned to the current position of the reference shape."
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_slide"+postfix,
                    label="Slide", on_toggle_action=bkt.Callback(self.specify_master_slide),  get_pressed=bkt.Callback(self.is_slide_specified),
                    screentip="Alignment to the slide",
                    supertip="Shapes are aligned to the slide."
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_contentarea"+postfix,
                    label="Content area [CTRL]",
                    on_toggle_action=bkt.Callback(self.specify_master_contentarea),
                    get_pressed=bkt.Callback(self.is_contentarea_specified),
                    screentip="Alignment to content area",
                    supertip="Shapes are aligned to the content area.\n\nThe content area corresponds to the area of the text placeholder on the reference slide.\n\nWith the CTRL key, this master is temporarily activated.",
                ),
                bkt.ribbon.ToggleButton(
                    id="arrange_use_customarea"+postfix,
                    label="Custom area",
                    on_toggle_action=bkt.Callback(self.specify_master_customarea_toggle),
                    get_pressed=bkt.Callback(self.is_customarea_specified),
                    get_enabled=bkt.Callback(self.is_customarea_specifiable),
                    screentip="Alignment to custom area",
                    supertip="Shapes are aligned to a fixed area that is previously defined by the user.",
                ),
                bkt.ribbon.MenuSeparator(),
                pplib.PositionGallery(
                    id="arrange_set_customarea"+postfix,
                    label="Choose custom area",
                    on_position_change = bkt.Callback(self.specify_master_customarea),
                    on_userdefined_area_change = bkt.Callback(self.specify_master_customarea),
                    get_item_supertip = bkt.Callback(self.get_item_supertip)
                ),
            ])
        ],
        **kwargs)

    def get_button(self, arrange_id, postfix="", **kwargs):
        return bkt.ribbon.Button(
            id=arrange_id+postfix,
            on_action=bkt.Callback(getattr(self, arrange_id), shapes=True),
            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
            image=arrange_id,
            show_label=False,
        **kwargs)


class ArrangeAdvancedEasy(ArrangeAdvanced):

    def __init__(self, resize_mode):
        self.resize_mode = resize_mode
        self.margin = 0
        # super(ArrangeAdvancedEasy, self).__init__()

    def get_master_from_shapes(self, shapes):
        if len(shapes) == 1:
            ## fallback if only one shape in selection
            return pplib.BoundingFrame(shapes[0].parent, contentarea=True)
        elif self.master_handler.master == "FIRST":
            return shapes[0]
        elif self.master_handler.master == "PPTDEFAULT":
            return pplib.BoundingFrame.from_shapes(shapes)
        else:
            # default: master == "LAST"
            return shapes[-1]

    # POSITION (resize=False)
    #arrange_left_at_left
    #arrange_middle_at_middle
    #arrange_right_at_right
    #arrange_top_at_top
    #arrange_vmiddle_at_vmiddle
    #arrange_bottom_at_bottom

    # DOCK (resize=False)
    #arrange_left_at_right
    #arrange_right_at_left
    #arrange_top_at_bottom
    #arrange_bottom_at_top

    # STRETCH (resize=True)
    #arrange_left_at_left
    #arrange_right_at_right
    #arrange_top_at_top
    #arrange_bottom_at_bottom

    # FILL (resize=True)
    #arrange_left_at_right
    #arrange_right_at_left
    #arrange_top_at_bottom
    #arrange_bottom_at_top

    def use_resizing(self):
        return self.resize_mode


class TAGallery(bkt.ribbon.Gallery):
    def __init__(self, **kwargs):
        self._locpin = None
        self.items = [
            ("fix_locpin_auto", "Auto", "Shapes werden bei Shape-Anordnung in:\n\n- Tabellenzellen horizontal und vertikal zentriert angeordnet,\n- Shapes horizontal zentriert sofern Shape höher ist als breit und vertikal zentriert sofern Shape breiter ist als hoch,\n- Textabsätzen vertikal zentriert von der ersten Zeile und horizontal nicht verändert."),
            ("fix_locpin_l", "Links", "Shapes werden links angeordnet und vertikal nicht verändert."),
            ("fix_locpin_c", "Zentriert", "Shapes werden zentriert angeordnet und vertikal nicht verändert."),
            ("fix_locpin_r", "Rechts", "Shapes werden rechts angeordnet und vertikal nicht verändert."),

            ("fix_locpin_t", "Oben", "Shapes werden oben angeordnet und horizontal nicht verändert."),
            ("fix_locpin_tl", "Oben-links", "Shapes werden oben links angeordnet."),
            ("fix_locpin_tm", "Oben-mitte", "Shapes werden oben zentriert angeordnet."),
            ("fix_locpin_tr", "Oben-rechts", "Shapes werden oben rechts angeordnet."),

            ("fix_locpin_m_line", "Mitte 1. Zeile", "Shapes werden vertikal zentriert von dem Text in der ersten Zeile angeordnet und horizontal nicht verändert."),
            ("fix_locpin_ml_line", "Mitte 1. Zeile-links", "Shapes werden vertikal zentriert von dem Text in der ersten Zeile und horizontal links angeordnet."),
            ("fix_locpin_mm_line", "Mitte 1. Zeile-mitte", "Shapes werden vertikal zentriert von dem Text in der ersten Zeile und horizontal zentriert angeordnet."),
            ("fix_locpin_mr_line", "Mitte 1. Zeile-rechts", "Shapes werden vertikal zentriert von dem Text in der ersten Zeile und horizontal rechts angeordnet."),

            ("fix_locpin_m", "Mitte", "Shapes werden vertikal zentriert angeordnet und horizontal nicht verändert."),
            ("fix_locpin_ml", "Mitte-links", "Shapes werden mittig links angeordnet."),
            ("fix_locpin_mm", "Mitte-mitte", "Shapes werden mittig zentriert angeordnet."),
            ("fix_locpin_mr", "Mitte-rechts", "Shapes werden mittig rechts angeordnet."),

            ("fix_locpin_b", "Unten", "Shapes werden unten angeordnet und horizontal nicht verändert."),
            ("fix_locpin_bl", "Unten-links", "Shapes werden unten links angeordnet."),
            ("fix_locpin_bm", "Unten-mitte", "Shapes werden unten zentriert angeordnet."),
            ("fix_locpin_br", "Unten-rechts", "Shapes werden unten rechts angeordnet."), 
        ]
        
        my_kwargs = dict(
            columns="4",
            item_height="32",
            item_width="32",
            on_action_indexed  = bkt.Callback(self.locpin_on_action_indexed),
            get_selected_item_index = bkt.Callback(self.locpin_get_index),
            get_image = bkt.Callback(self.locpin_get_image, context=True),
            get_item_count = bkt.Callback(lambda: len(self.items)),
            # get_item_label = bkt.Callback(lambda index: self.items[index][1]),
            get_item_image = bkt.Callback(self.locpin_get_image, context=True),
            get_item_screentip = bkt.Callback(lambda index: self.items[index][1]),
            get_item_supertip = bkt.Callback(lambda index: self.items[index][2]),
            # children = [
            #     bkt.ribbon.Item(image=gal_item[0], screentip=gal_item[1], supertip=gal_item[2])
            #     for gal_item in self.items
            # ]
        )
        my_kwargs.update(kwargs)
        
        super(TAGallery, self).__init__(**my_kwargs)

    @property
    def locpin(self):
        if not self._locpin:
            from .models.tablearrange import TALocPin
            self._locpin = TALocPin()
        return self._locpin


    def locpin_get_index(self):
        return self.locpin.index

    def locpin_on_action_indexed(self, selected_item, index):
        from .models.tablearrange import TableArrange
        self.locpin.index = index
        TableArrange.vertical_arrangement, TableArrange.horizontal_arrangement = self.locpin.fixation
    
    def locpin_get_image(self, context, index=None):
        if index is None:
            return context.python_addin.load_image(self.items[self.locpin.index][0])
        else:
            return context.python_addin.load_image(self.items[index][0])


tablearrange_button = bkt.ribbon.SplitButtonFixed(
    show_label=False,
    get_enabled = bkt.apps.ppt_shapes_min2_selected,
    children=[
        bkt.ribbon.Button(
            id = 'table-shapes',
            label="Arrange on table/paragraph/shapes",
            image='arrange_shape_table',
            screentip="Arrange shape objects in tables/paragraphs/shapes",
            supertip="When tables and shapes are selected:\nArrange each shape that lies over a (also selected) table within the nearest table cell. The cell is determined by the shape center.\n\nWhen shapes are selected:\nArrange each shape that lies completely over another (also selected) shape (=reference shape) in the nearest text paragraph within the reference shape. If fewer than 2 text paragraphs are present, arrangement is done within the entire shape.",
            on_action=bkt.CallbackLazy("toolbox.models.tablearrange", "TableArrange", "arrange_overlay_shapes", shapes=True, shapes_min=2),
            # get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.Menu(label="Arrange on table/paragraph/shapes menu", supertip="Settings for arranging on tables, paragraphs or shapes", item_size="large", children=[
            bkt.ribbon.MenuSeparator(title="Arrange shapes"),
            bkt.ribbon.Button(
                id = 'table-shapes2',
                label="Arrange automatically",
                image='arrange_shape_table_auto',
                screentip="Arrange shape objects in table cells/paragraphs/shapes",
                description="Automatic selection of the arrange functions (table cells/paragraphs/shapes)",
                on_action=bkt.CallbackLazy("toolbox.models.tablearrange", "TableArrange", "arrange_overlay_shapes", shapes=True, shapes_min=2),
                # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            ),
            TAGallery(
                id="arrange_shape_table_mode",
                label="Adjust arrangement",
                screentip="Set the shape position when arranging.",
                description="Selection of horizontal and vertical arrangement."
            ),
            bkt.ribbon.MenuSeparator(title="Manual selection"),
            bkt.ribbon.Button(
                id="arrange_on_table",
                image="arrange_on_table",
                label="Arrange on table cells",
                screentip="Arrange shape objects in table cells",
                description="The center of the shapes must lie within the table cell for automatic arrangement.",
                on_action=bkt.CallbackLazy("toolbox.models.tablearrange", "TableArrange", "arrange_table_shapes", shapes=True, shapes_min=2),
                # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            ),
            bkt.ribbon.Button(
                id="arrange_on_paragraph",
                image="arrange_on_paragraph",
                label="Arrange on text paragraphs",
                screentip="Arrange shape objects in text paragraphs",
                description="The center of the shapes must lie within the text shape and the respective paragraph for automatic arrangement.",
                on_action=bkt.CallbackLazy("toolbox.models.tablearrange", "TableArrange", "arrange_paragraph_shapes", shapes=True, shapes_min=2),
                # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            ),
            bkt.ribbon.Button(
                id="arrange_on_shapes",
                image="arrange_on_shapes",
                label="Arrange on background shapes",
                screentip="Arrange shape objects in subordinate shapes",
                description="The center of the shapes must lie within the underlying shapes for automatic arrangement.",
                on_action=bkt.CallbackLazy("toolbox.models.tablearrange", "TableArrange", "arrange_shapes_shapes", shapes=True, shapes_min=2),
                # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            ),

            # bkt.ribbon.MenuSeparator(title="Horizontal arrangement"),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_hauto",
            #     label="Automatic",
            #     screentip="Shape position automatic",
            #     supertip="When arranging shapes, shapes are:\n\n- arranged horizontally centered in table cells,\n- centered horizontally in shapes if the shape is taller than wide,\n- not arranged in text paragraphs.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_HAUTO) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_HAUTO)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_hnone",
            #     label="No change",
            #     screentip="Do not change shape position",
            #     supertip="Shapes are not arranged horizontally when arranging shapes in table cells/text paragraphs/shapes.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_HNONE) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_HNONE)
            # ),
            # bkt.ribbon.MenuSeparator(),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_left",
            #     label="Left",
            #     screentip="Arrange shapes left",
            #     supertip="Shapes are arranged left when arranging shapes in table cells/text paragraphs/shapes.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_LEFT) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_LEFT)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_center",
            #     label="Centered",
            #     screentip="Arrange shapes centered",
            #     supertip="Shapes are arranged horizontally centered when arranging shapes in table cells/text paragraphs/shapes.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_HCENTER) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_HCENTER)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_right",
            #     label="Right",
            #     screentip="Arrange shapes right",
            #     supertip="Shapes are arranged right when arranging shapes in table cells/text paragraphs/shapes.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'horizontal_arrangement', TableArrange.ARRANGE_RIGHT) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.horizontal_arrangement == TableArrange.ARRANGE_RIGHT)
            # ),
            # bkt.ribbon.MenuSeparator(title="Vertical arrangement"),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_vauto",
            #     label="Automatic",
            #     screentip="Shape position automatic",
            #     supertip="When arranging shapes, shapes are:\n\n- arranged vertically centered in table cells,\n- centered vertically in shapes if the shape is wider than tall,\n- arranged vertically centered on the first row in text paragraphs.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_VAUTO) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_VAUTO)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_vnone",
            #     label="No change",
            #     screentip="Do not change shape position",
            #     supertip="Shapes are not arranged vertically when arranging shapes in table cells/text paragraphs/shapes.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_VNONE) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_VNONE)
            # ),
            # bkt.ribbon.MenuSeparator(),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_top",
            #     label="Top",
            #     screentip="Arrange shapes top",
            #     supertip="Shapes are arranged top when arranging shapes in table cells/text paragraphs/shapes.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_TOP) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_TOP)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_vcenter",
            #     label="Center",
            #     screentip="Arrange shapes centered",
            #     supertip="Shapes are arranged vertically centered when arranging shapes in table cells/text paragraphs/shapes.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_VCENTER) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_VCENTER)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_lcenter",
            #     label="Center of 1st row",
            #     screentip="Arrange shapes centered on the first row",
            #     supertip="Shapes are arranged vertically centered on the text in the first row when arranging shapes in table cells/text paragraphs/shapes.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_LCENTER) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_LCENTER)
            # ),
            # bkt.ribbon.CheckBox(
            #     id="arrange_shape_table_bottom",
            #     label="Bottom",
            #     screentip="Arrange shapes bottom",
            #     supertip="Shapes are arranged bottom when arranging shapes in table cells/text paragraphs/shapes.",
            #     on_toggle_action=bkt.Callback(lambda pressed: setattr(TableArrange, 'vertical_arrangement', TableArrange.ARRANGE_BOTTOM) ),
            #     get_pressed=bkt.Callback(lambda : TableArrange.vertical_arrangement == TableArrange.ARRANGE_BOTTOM)
            # )
        ])
    ]
)


class UiFactory(object):
    def __init__(self):
        self._groups = dict()

    def __getattr__(self, attr):
        return self._groups[attr]()
    
    def __setattr__(self, attr, value):
        if attr.startswith("_"):
            super().__setattr__(attr, value)
        else:
            self._groups[attr] = value

UiGroups = UiFactory()


UiGroups.arrange_group = lambda: bkt.ribbon.Group(
    id="bkt_align_group",
    label='Arrange',
    image_mso='ObjectsArrangeMenu',
    children = [
        bkt.ribbon.Box(box_style="vertical",
            children = [
                equal_height_button,
                equal_width_button,
                swap_button,
            ]
        ),

        bkt.mso.control.ObjectsAlignLeftSmart,
        bkt.mso.control.ObjectsAlignRightSmart,
        bkt.mso.control.ObjectsAlignCenterHorizontalSmart,
        bkt.mso.control.ObjectsAlignTopSmart,
        bkt.mso.control.ObjectsAlignBottomSmart,
        bkt.mso.control.ObjectsAlignMiddleVerticalSmart,
        bkt.mso.control.AlignDistributeHorizontally,
        bkt.mso.control.AlignDistributeVertically,
        
        #bkt.mso.control.ObjectRotateGallery,
        bkt.ribbon.DynamicMenu(
            label='☰', #⋮
            # show_label=False,
            # image_mso='TableDesign',
            screentip="More functions",
            supertip="Functions like positioning, linked shapes, ...",
            get_content=bkt.CallbackLazy("toolbox.models.arrange_menu", "arrange_menu")
        ),
        
        bkt.mso.control.ObjectsGroup,
        bkt.mso.control.ObjectsUngroup,
        bkt.mso.control.ObjectsRegroup,
        # bkt.mso.control.ObjectBringToFrontMenu,
        bkt.ribbon.SplitButtonFixed(
            id="bkt_bringtofront_menu",
            show_label=False,
            children=[
                bkt.mso.button.ObjectBringToFront,
                bkt.ribbon.Menu(label="Foreground menu", supertip="Functions to bring shapes to the front", children=[
                    bkt.mso.control.ObjectBringToFront,
                    bkt.mso.control.ObjectBringForward,
                    bkt.ribbon.Button(
                        label="Back to front",
                        supertip="Brings all back shapes exactly in front of the frontmost shape",
                        image="zorder_back_to_front",
                        get_enabled=bkt.apps.ppt_shapes_min2_selected,
                        on_action=bkt.CallbackLazy("toolbox.shapes", "PositionSize", "back_to_front", shapes=True),
                    ),
                ])
            ]
        ),
        # bkt.mso.control.ObjectSendToBackMenu,
        bkt.ribbon.SplitButtonFixed(
            id="bkt_sendtoback_menu",
            show_label=False,
            children=[
                bkt.mso.button.ObjectSendToBack,
                bkt.ribbon.Menu(label="Background menu", supertip="Functions to send shapes to the back", children=[
                    bkt.mso.control.ObjectSendToBack,
                    bkt.mso.control.ObjectSendBackward,
                    bkt.ribbon.Button(
                        label="Front to back",
                        supertip="Brings all front shapes exactly behind the backmost shape",
                        image="zorder_front_to_back",
                        get_enabled=bkt.apps.ppt_shapes_min2_selected,
                        on_action=bkt.CallbackLazy("toolbox.shapes", "PositionSize", "front_to_back", shapes=True),
                    ),
                ])
            ]
        ),

        tablearrange_button

    ]
)

UiGroups.distance_rotation_group = lambda: bkt.ribbon.Group(
    id="bkt_shapesep_group",
    label="Distance/rotation",
    image_mso='VerticalSpacingIncrease',
    children=[
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_sep_v',
            image_mso='VerticalSpacingIncrease',
            # image_button=True,
            size_string="-###",
            label="Object distance vertical",
            show_label=False,
            screentip="Vertical object distance",
            supertip="Change the vertical object distance to the specified value (in cm).\n\nIcon click for 0 spacing.\nShift click for 0.2 cm spacing.\nCtrl click to match spacing.",
            on_change = bkt.Callback(ShapeDistance.set_shape_sep_vertical, shapes=True),
            get_text  = bkt.Callback(ShapeDistance.get_shape_sep_vertical, shapes=True),
            # on_action = bkt.Callback(ShapeDistance.set_shape_sep_vertical_zero, shapes=True),
            get_enabled = bkt.Callback(ShapeDistance.get_enabled_min2_group, shapes=True),
            # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            round_cm = True,
            convert="pt_to_cm",
            image_element=bkt.ribbon.SplitButtonFixed(show_label=False, children=[
                bkt.ribbon.Button(
                    on_action = bkt.Callback(ShapeDistance.set_shape_sep_vertical_zero, shapes=True),
                ),
                bkt.ribbon.Menu(label="Object distance vertical menu", supertip="Options for calculating the vertical distance between shapes", item_size="large", children=[
                    bkt.ribbon.MenuSeparator(title="Vertical spacing"),
                    bkt.ribbon.ToggleButton(
                        label="Shape spacing (default)",
                        image="shapedis_vdistance",
                        description="Show distance from bottom edge to top edge.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="distance"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "distance")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Visual distance (with rotation)",
                        image="shapedis_vvisual",
                        description="Show distance from visual bottom edge to visual top edge. Helpful for rotated shapes.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="visual"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "visual")),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.ToggleButton(
                        label="Distance to top edges",
                        image="shapedis_top",
                        description="Show distance from top edge to top edge.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="top"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "top")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Distance to center points",
                        image="shapedis_vcenter",
                        description="Show distance from the respective center points.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="center"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "center")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Distance to bottom edges",
                        image="shapedis_bottom",
                        description="Show distance from bottom edge to bottom edge.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_edges=="bottom"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_edges", "bottom")),
                    ),
                    bkt.ribbon.MenuSeparator(title="Direction of movement"),
                    bkt.ribbon.ToggleButton(
                        label="Down",
                        image="shapedis_movedown",
                        description="Changes the distance starting from the top shape and pushes all others down. Temporary switch to 'Up' with [ALT].",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_fix=="top"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_fix", "top")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Up",
                        image="shapedis_moveup",
                        description="Changes the distance starting from the bottom shape and pushes all others up. Temporary switch to 'Down' with [ALT].",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.vertical_fix=="bottom"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("vertical_fix", "bottom")),
                    ),
                ])
            ]),
        ),

        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_sep_h',
            image_mso='HorizontalSpacingIncrease',
            # image_button=True,
            size_string="-###",
            label="Object distance horizontal",
            show_label=False,
            screentip="Horizontal object distance",
            supertip="Change the horizontal object distance to the specified value (in cm).\n\nIcon click for 0 spacing.\nShift click for 0.2 cm spacing.\nCtrl click to match spacing.",
            on_change = bkt.Callback(ShapeDistance.set_shape_sep_horizontal, shapes=True),
            get_text  = bkt.Callback(ShapeDistance.get_shape_sep_horizontal, shapes=True),
            # on_action = bkt.Callback(ShapeDistance.set_shape_sep_horizontal_zero, shapes=True),
            # get_enabled = bkt.apps.ppt_shapes_min2_selected,
            get_enabled = bkt.Callback(ShapeDistance.get_enabled_min2_group, shapes=True),
            round_cm = True,
            convert="pt_to_cm",
            image_element=bkt.ribbon.SplitButtonFixed(show_label=False, children=[
                bkt.ribbon.Button(
                    on_action = bkt.Callback(ShapeDistance.set_shape_sep_horizontal_zero, shapes=True),
                ),
                bkt.ribbon.Menu(label="Object distance horizontal menu", supertip="Options for calculating the horizontal distance between shapes", item_size="large", children=[
                    bkt.ribbon.MenuSeparator(title="Horizontal spacing"),
                    bkt.ribbon.ToggleButton(
                        label="Shape spacing (default)",
                        image="shapedis_hdistance",
                        description="Show distance from right edge to left edge.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="distance"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "distance")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Visual distance (with rotation)",
                        image="shapedis_hvisual",
                        description="Distance from visual right edge to visual left edge. Helpful for rotated shapes.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="visual"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "visual")),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.ToggleButton(
                        label="Distance to left edges",
                        image="shapedis_left",
                        description="Show distance from left edge to left edge.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="left"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "left")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Distance to center points",
                        image="shapedis_hcenter",
                        description="Show distance from the respective center points.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="center"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "center")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Distance to right edges",
                        image="shapedis_right",
                        description="Show distance from right edge to right edge.",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_edges=="right"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_edges", "right")),
                    ),
                    bkt.ribbon.MenuSeparator(title="Direction of movement"),
                    bkt.ribbon.ToggleButton(
                        label="To the right",
                        image="shapedis_moveright",
                        description="Changes the distance starting from the left shape and pushes all others to the right. Temporary switch to 'To the left' with [ALT].",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_fix=="left"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_fix", "left")),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="To the left",
                        image="shapedis_moveleft",
                        description="Changes the distance starting from the right shape and pushes all others to the left. Temporary switch to 'To the right' with [ALT].",
                        get_pressed=bkt.Callback(lambda: ShapeDistance.horizontal_fix=="right"),
                        on_toggle_action=bkt.Callback(lambda pressed: ShapeDistance.change_settings("horizontal_fix", "right")),
                    ),
                ])
            ]),
        ),

        bkt.ribbon.RoundingSpinnerBox(
            id = 'rotation',
            label="Rotation",
            show_label=False,
            image_mso='Repeat',
            # image_button=True,
            size_string="-###",
            screentip="Shape rotation",
            supertip="Change the rotation of the shape to the specified value (in degrees).\n\nIcon click for rotation=0.\nShift click for rotation=180.\nCtrl click to match rotation.",
            on_change = bkt.Callback(ShapeRotation.set_rotation, shapes=True),
            get_text  = bkt.Callback(ShapeRotation.get_rotation, shapes=True),
            # on_action = bkt.Callback(ShapeRotation.set_rotation_zero, shapes=True),
            get_enabled = bkt.Callback(ShapeRotation.get_enabled, selection=True),
            # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            round_int = True,
            huge_step = 45,
            image_element=bkt.ribbon.SplitButtonFixed(show_label=False, children=[
                bkt.ribbon.Button(
                    on_action = bkt.Callback(ShapeRotation.set_rotation_zero, shapes=True),
                ),
                bkt.ribbon.Menu(label="Rotation and flip menu", supertip="Rotate shapes and show flip", children=[
                    bkt.ribbon.MenuSeparator(title="Rotation"),
                    pplib.LocpinGallery(
                        label="Anchor point when rotating",
                        screentip="Set anchor point when rotating",
                        supertip="Sets the point that should be fixed when rotating the shapes.",
                        locpin=ShapeRotation.locpin,
                    ),
                    bkt.ribbon.Button(
                        label="Set to 0 degrees",
                        screentip="Set shape rotation to 0 degrees",
                        supertip="Set the shape rotation of all selected shapes to 0 degrees",
                        on_action=bkt.Callback(lambda shapes: ShapeRotation.set_rotation(shapes, 0), shapes=True),
                    ),
                    bkt.ribbon.Button(
                        label="Set to 180 degrees",
                        screentip="Set shape rotation to 180 degrees",
                        supertip="Set the shape rotation of all selected shapes to 180 degrees",
                        on_action=bkt.Callback(lambda shapes: ShapeRotation.set_rotation(shapes, 180), shapes=True),
                    ),
                    # bkt.mso.button.ObjectRotateRight90,
                    # bkt.mso.button.ObjectRotateLeft90,
                    bkt.ribbon.MenuSeparator(title="Flip"),
                    bkt.ribbon.ToggleButton(
                        label="Toggle flipped horizontally",
                        screentip="Shape is flipped horizontally",
                        supertip="Activated when the first selected shape is flipped horizontally",
                        image_mso="ObjectFlipHorizontal",
                        get_pressed=bkt.Callback(ShapeRotation.get_pressed_fliph, shapes=True),
                        on_toggle_action=bkt.Callback(ShapeRotation.set_fliph, shapes=True),
                    ),
                    bkt.ribbon.ToggleButton(
                        label="Toggle flipped vertically",
                        screentip="Shape is flipped vertically",
                        supertip="Activated when the first selected shape is flipped vertically",
                        image_mso="ObjectFlipVertical",
                        get_pressed=bkt.Callback(ShapeRotation.get_pressed_flipv, shapes=True),
                        on_toggle_action=bkt.Callback(ShapeRotation.set_flipv, shapes=True),
                    ),
                    # bkt.ribbon.MenuSeparator(),
                    # bkt.mso.button.ObjectFlipHorizontal,
                    # bkt.mso.button.ObjectFlipVertical,
                    bkt.ribbon.MenuSeparator(title="Options"),
                    bkt.mso.button.ObjectRotationOptionsDialog,
                ])
            ]),
        ),
    ]
)

UiGroups.euclid_angle_group = lambda: bkt.ribbon.Group(
    id="bkt_shapesep_advanced",
    label="Adv. pos.",
    image_mso='VerticalSpacingIncrease',
    children=[
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_sep_euclid',
            image='shape_euclid_distance',
            #image_button=True,
            label="Object distance Euclidean",
            show_label=False,
            screentip="Euclidean object distance",
            supertip="Changes the Euclidean object distance to the specified value (in cm). Measured from the defined reference point of the two first selected shapes.\n\nWith the ALT key, each shape is aligned at the same delta distance relative to the first shape.",
            on_change = bkt.Callback(ShapeEuclid.set_shape_sep_euclid, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(ShapeEuclid.get_shape_sep_euclid, shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
            round_cm = True,
            convert="pt_to_cm",
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'shape_angle',
            image='shape_angle',
            #image_button=True,
            label="Angle",
            show_label=False,
            screentip="Angle of the shapes to each other",
            supertip="Changes the angle of the shapes to the X axis (in degrees). Measured from the defined reference point of the two first selected shapes.\n\nWith the ALT key, each shape is moved by the same delta angle relative to the first shape.",
            on_change = bkt.Callback(ShapeEuclid.set_shape_angle, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(ShapeEuclid.get_shape_angle, shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
            round_int = True,
            huge_step = 45,
        ),
        bkt.ribbon.Menu(
            label="Ref. point",
            image="fix_locpin_mm",
            screentip="Set reference points",
            supertip="Set reference or anchor points for distance and angle for the first shape and the other shapes",
            children=[
                bkt.ribbon.MenuSeparator(title="Set center shape"),
                bkt.ribbon.MenuSeparator(title="Set reference points"),
                pplib.LocpinGallery(
                    label="Inside center shape",
                    screentip="Reference or anchor point for distance and angle",
                    supertip="Sets the point within the center shape from which the distance or angle of the shapes is measured.",
                    locpin=LocPinCollection.dis1,
                ),
                pplib.LocpinGallery(
                    label="Inside other shapes",
                    screentip="Reference or anchor point for distance and angle",
                    supertip="Sets the point within all other shapes from which the distance or angle of the shapes is measured.",
                    locpin=LocPinCollection.dis2,
                ),
                bkt.ribbon.MenuSeparator(title="Set advanced settings"),
                bkt.ribbon.Menu(
                    label="Select center shape",
                    screentip="Select shape in the center",
                    supertip="Determines which shape within the selection should be in the center.",
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="First selected shape (default)",
                            supertip="Specifies that the first selected shape within the selection should be in the center.",
                            get_pressed=bkt.Callback(lambda: ShapeEuclid.shape1_index==0),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeEuclid, "shape1_index", 0)),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Last selected shape",
                            supertip="Specifies that the last selected shape within the selection should be in the center.",
                            get_pressed=bkt.Callback(lambda: ShapeEuclid.shape1_index==-1),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeEuclid, "shape1_index", -1)),
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    label="Behavior for >2 shapes",
                    screentip="Behavior when more than 2 shapes are selected",
                    supertip="When more than 2 shapes are selected, different options can be chosen for the calculation between the center shape and all other shapes.",
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="Distance/angle individually from center to shapes (default)",
                            supertip="Specifies that distance or angle is calculated for each shape individually from the center to the respective shape.\n\nThis function is useful to change the angle of a diagonal of shapes or to achieve a star-shaped arrangement of shapes.",
                            get_pressed=bkt.Callback(lambda: ShapeEuclid.euclid_multi_shape_mode == "centric"),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeEuclid, "euclid_multi_shape_mode", "centric")),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Always change distances/angles by an equal difference [ALT]",
                            supertip="Specifies that distance or angle is always changed by the same difference from the center to the respective shape.\n\nTemporary toggle possible with the ALT key.",
                            get_pressed=bkt.Callback(lambda: ShapeEuclid.euclid_multi_shape_mode == "delta"),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeEuclid, "euclid_multi_shape_mode", "delta")),
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Distribute distances/angles evenly between shapes",
                            supertip="Specifies that distance or angle is distributed evenly from the center and between each shape.\n\nThis function is useful to distribute several shapes along a diagonal or around a center.",
                            get_pressed=bkt.Callback(lambda: ShapeEuclid.euclid_multi_shape_mode == "distribute"),
                            on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeEuclid, "euclid_multi_shape_mode", "distribute")),
                        ),
                    ]
                ),
                bkt.ribbon.ToggleButton(
                    label="Match shape rotation",
                    screentip="Toggle match shape rotation to angle",
                    supertip="Adjusts the shape rotation to the distance vector or the angle.",
                    get_pressed=bkt.Callback(lambda: getattr(ShapeEuclid, "shape_rotate_with_angle")),
                    on_toggle_action=bkt.Callback(lambda pressed: setattr(ShapeEuclid, "shape_rotate_with_angle", pressed)),
                ),
            ]
        )
    ]
)

def get_advanced_group():
    arrange_advaced = ArrangeAdvanced()

    return bkt.ribbon.Group(
    id="bkt_arrage_adv_group",
    label='Advanced arranging',
    image='arrange_left_at_left',
    children=[
        arrange_advaced.get_button('arrange_left_at_left',         label="Left to left",   screentip='Align left edge to left edge',   supertip='Alignment of the left edge to the left edge of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_middle_at_left',       label="Center to left",   screentip='Align center to left edge',         supertip='Alignment of the shape center to the left edge of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_right_at_left',        label="Right to left",  screentip='Align right edge to left edge',  supertip='Alignment of the right edge to the left edge of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_left_at_middle',       label="Left to center",   screentip='Align left edge to shape center',    supertip='Alignment of the left edge to the shape center of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_middle_at_middle',     label="Center to center",   screentip='Align shape center to shape center',     supertip='Alignment of the shape center to the shape center of the reference shape.\n(no spacing)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_right_at_middle',      label="Right to center",  screentip='Align right edge to shape center',   supertip='Alignment of the right edge to the shape center of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_left_at_right',        label="Left to right",  screentip='Align left edge to right edge',  supertip='Alignment of the left edge to the right edge of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_middle_at_right',      label="Center to right",  screentip='Align shape center to right edge',   supertip='Alignment of the shape center to the right edge of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_right_at_right',       label="Right to right", screentip='Align right edge to right edge', supertip='Alignment of the right edge to the right edge of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_top_at_top',           label="Top to top",     screentip='Align top edge to top edge',   supertip='Alignment of the top edge to the top edge of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_middle_at_top',        label="Center to top",    screentip='Align shape center to top edge',    supertip='Alignment of the shape center to the top edge of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_bottom_at_top',        label="Bottom to top",    screentip='Align bottom edge to top edge',  supertip='Alignment of the bottom edge to the top edge of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_top_at_middle',        label="Top to center",    screentip='Align top edge to shape center',    supertip='Alignment of the top edge to the shape center of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_vmiddle_at_vmiddle',   label="Center to center",   screentip='Align shape center to shape center',     supertip='Alignment of the shape center to the shape center of the reference shape.\n(no spacing)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_bottom_at_middle',     label="Bottom to center",   screentip='Align bottom edge to shape center',   supertip='Alignment of the bottom edge to the shape center of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_top_at_bottom',        label="Top to bottom",    screentip='Align top edge to bottom edge',  supertip='Alignment of the top edge to the bottom edge of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_middle_at_bottom',     label="Center to bottom",   screentip='Align shape center to bottom edge',   supertip='Alignment of the shape center to the bottom edge of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button('arrange_bottom_at_bottom',     label="Bottom to bottom",   screentip='Align bottom edge to bottom edge', supertip='Alignment of the bottom edge to the bottom edge of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        
        bkt.ribbon.Separator(),

        bkt.ribbon.RoundingSpinnerBox(
            id='arrange_distance',
            label="Alignment spacing",
            show_label=False,
            round_cm=True, convert="pt_to_cm",
            image_mso="HorizontalSpacingIncrease", on_change=bkt.Callback(arrange_advaced.set_margin), get_text=bkt.Callback(arrange_advaced.get_margin),
            screentip="Spacing when aligning",
            supertip="The set spacing is taken into account when aligning shapes left/right.\n\nThe spacing is added: when aligning the left/top edge (of the shape to be moved) and when aligning the shape center to the left/top edge of the reference shape.\n\nThe spacing is subtracted: when aligning the right/bottom edge (of the shape to be moved) and when aligning the shape center to the right/bottom edge of the reference shape."
        ),
        
        #bkt.ribbon.Menu(label="Options", children=[
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Label(label="Mode:"),
            bkt.ribbon.ToggleButton(id="arrange_move",   label="Move",         show_label=False, supertip="The desired shape arrangement is achieved by positioning shapes", image_mso="ObjectNudgeRight", on_toggle_action=bkt.Callback(arrange_advaced.set_moving),   get_pressed=bkt.Callback(arrange_advaced.get_moving)),
            bkt.ribbon.ToggleButton(id="arrange_resize", label="Stretch/compress [SHIFT]", show_label=False, supertip="The desired shape arrangement is achieved by shrinking/enlarging shapes", image_mso="ShapeWidth",       on_toggle_action=bkt.Callback(arrange_advaced.set_resizing), get_pressed=bkt.Callback(arrange_advaced.get_resizing))
        ]),
        
        arrange_advaced.get_master_button(),

        bkt.ribbon.Separator(),

        bkt.ribbon.Label(label="Quick select:"),
        bkt.ribbon.Button(
            id="arrange_quick_position",
            label="Position",
            image_mso="ControlAlignToGrid",
            on_action=bkt.Callback(arrange_advaced.arrange_quick_position, shapes=True),
            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
            screentip="Same position as reference",
            supertip="Shapes get the same position as the reference shape.",
        ),
        bkt.ribbon.Button(
            id="arrange_quick_size",
            label="Size",
            image_mso="SizeToControlHeightAndWidth",
            on_action=bkt.Callback(arrange_advaced.arrange_quick_size, shapes=True),
            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
            screentip="Same size as reference",
            supertip="Shapes get the same size as the reference shape.",
        ),
    ]
)

UiGroups.arrange_advanced_group = get_advanced_group

def get_advanced_small():
    arrange_advaced = ArrangeAdvanced()
    
    return bkt.ribbon.Group(
    id="bkt_arrage_adv_small_group",
    label='Adv. arrange',
    image='arrange_left_at_left',
    children=[
        arrange_advaced.get_button("arrange_left_at_left", "-small",        label="Left to left",   screentip='Align left edge to left edge',   supertip='Alignment of the left edge to the left edge of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button("arrange_right_at_right", "-small",      label="Right to right", screentip='Align right edge to right edge', supertip='Alignment of the right edge to the right edge of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button("arrange_middle_at_middle", "-small",    label="Center to center",   screentip='Align shape center to shape center',     supertip='Alignment of the shape center to the shape center of the reference shape.\n(no spacing)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button("arrange_top_at_top", "-small",          label="Top to top",     screentip='Align top edge to top edge',   supertip='Alignment of the top edge to the top edge of the reference shape.\n(spacing is added)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button("arrange_bottom_at_bottom", "-small",    label="Bottom to bottom",   screentip='Align bottom edge to bottom edge', supertip='Alignment of the bottom edge to the bottom edge of the reference shape.\n(spacing is subtracted)\n\nWith the SHIFT key it switches to stretch/compress.'),
        arrange_advaced.get_button("arrange_vmiddle_at_vmiddle", "-small",  label="Center to center",   screentip='Align shape center to shape center',     supertip='Alignment of the shape center to the shape center of the reference shape.\n(no spacing)\n\nWith the SHIFT key it switches to stretch/compress.'),

        bkt.ribbon.Button(
            id="arrange_quick_position-small",
            label="Position",
            show_label=False,
            image_mso="ControlAlignToGrid",
            on_action=bkt.Callback(arrange_advaced.arrange_quick_position, shapes=True),
            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
            screentip="Same position as reference",
            supertip="Shapes get the same position as the reference shape.",
        ),
        bkt.ribbon.Button(
            id="arrange_quick_size-small",
            label="Size",
            show_label=False,
            image_mso="SizeToControlHeightAndWidth",
            on_action=bkt.Callback(arrange_advaced.arrange_quick_size, shapes=True),
            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
            screentip="Same size as reference",
            supertip="Shapes get the same size as the reference shape.",
        ),

        arrange_advaced.get_master_button("-small", show_label=False)
    ]
)

UiGroups.arrange_advanced_small_group = get_advanced_small

def get_advanced_easy():
    arrange_adv_position = ArrangeAdvancedEasy(False)
    arrange_adv_size     = ArrangeAdvancedEasy(True)

    return bkt.ribbon.Group(
    id="bkt_arrage_adv_easy_group",
    label='Simple arranging',
    image='arrange_left_at_left',
    children=[
        #POSITION
        bkt.ribbon.Button(id='arrange_position_left',     on_action=bkt.Callback(arrange_adv_position.arrange_left_at_left, shapes=True),       get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_position_left',        label="Left to left",   show_label=False, screentip='Align left edge to left edge',   supertip='Alignment of the left edge to the left edge of the reference shape.'),
        bkt.ribbon.Button(id='arrange_position_right',    on_action=bkt.Callback(arrange_adv_position.arrange_right_at_right, shapes=True),     get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_position_right',       label="Right to right", show_label=False, screentip='Align right edge to right edge', supertip='Alignment of the right edge to the right edge of the reference shape.'),
        bkt.ribbon.Button(id='arrange_position_middle',   on_action=bkt.Callback(arrange_adv_position.arrange_middle_at_middle, shapes=True),   get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_position_middle',      label="Center to center",   show_label=False, screentip='Align shape center to shape center',     supertip='Alignment of the shape center to the shape center of the reference shape.'),

        bkt.ribbon.Button(id='arrange_position_top',      on_action=bkt.Callback(arrange_adv_position.arrange_top_at_top, shapes=True),         get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_position_top',        label="Top to top",     show_label=False, screentip='Align top edge to top edge',   supertip='Alignment of the top edge to the top edge of the reference shape.'),
        bkt.ribbon.Button(id='arrange_position_bottom',   on_action=bkt.Callback(arrange_adv_position.arrange_bottom_at_bottom, shapes=True),   get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_position_bottom',     label="Bottom to bottom",   show_label=False, screentip='Align bottom edge to bottom edge', supertip='Alignment of the bottom edge to the bottom edge of the reference shape.'),
        bkt.ribbon.Button(id='arrange_position_vmiddle',  on_action=bkt.Callback(arrange_adv_position.arrange_vmiddle_at_vmiddle, shapes=True), get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_position_vmiddle',    label="Center to center",   show_label=False, screentip='Align shape center to shape center',     supertip='Alignment of the shape center to the shape center of the reference shape.'),
        bkt.ribbon.Separator(),

        #DOCK
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Button(id='arrange_dock_left',    on_action=bkt.Callback(arrange_adv_position.arrange_left_at_right, shapes=True),      get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_dock_left',      label="Left to right",  show_label=False, screentip='Dock right',  supertip='Alignment of the left edge to the right edge of the reference shape.'),
            bkt.ribbon.Button(id='arrange_dock_right',   on_action=bkt.Callback(arrange_adv_position.arrange_right_at_left, shapes=True),      get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_dock_right',     label="Right to left",  show_label=False, screentip='Dock left',  supertip='Alignment of the right edge to the left edge of the reference shape.'),
            bkt.ribbon.Button(id='arrange_dock_bottom',  on_action=bkt.Callback(arrange_adv_position.arrange_bottom_at_top, shapes=True),      get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_dock_bottom',       label="Bottom to top",    show_label=False, screentip='Dock top',  supertip='Alignment of the bottom edge to the top edge of the reference shape.'),
            bkt.ribbon.Button(id='arrange_dock_top',     on_action=bkt.Callback(arrange_adv_position.arrange_top_at_bottom, shapes=True),      get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_dock_top',    label="Top to bottom",    show_label=False, screentip='Dock bottom',  supertip='Alignment of the top edge to the bottom edge of the reference shape.'),
        ]),

        #STRETCH
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Button(id='arrange_stretch_left',   on_action=bkt.Callback(arrange_adv_size.arrange_left_at_left, shapes=True),       get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_stretch_left',        label="Left to left",   show_label=False, screentip='Stretch to the left',   supertip='Alignment of the left edge to the left edge of the reference shape.'),
            bkt.ribbon.Button(id='arrange_stretch_right',  on_action=bkt.Callback(arrange_adv_size.arrange_right_at_right, shapes=True),     get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_stretch_right',       label="Right to right", show_label=False, screentip='Stretch to the right', supertip='Alignment of the right edge to the right edge of the reference shape.'),

            bkt.ribbon.Button(id='arrange_stretch_top',    on_action=bkt.Callback(arrange_adv_size.arrange_top_at_top, shapes=True),         get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_stretch_top',         label="Top to top",     show_label=False, screentip='Stretch up',   supertip='Alignment of the top edge to the top edge of the reference shape.'),
            bkt.ribbon.Button(id='arrange_stretch_bottom', on_action=bkt.Callback(arrange_adv_size.arrange_bottom_at_bottom, shapes=True),   get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_stretch_bottom',      label="Bottom to bottom",   show_label=False, screentip='Stretch down', supertip='Alignment of the bottom edge to the bottom edge of the reference shape.'),
        ]),

        #FILL
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Button(id='arrange_fill_left',       on_action=bkt.Callback(arrange_adv_size.arrange_left_at_right, shapes=True),      get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_fill_left',          label="Left to right",  show_label=False, screentip='Fill left gap',  supertip='Alignment of the left edge to the right edge of the reference shape.'),
            bkt.ribbon.Button(id='arrange_fill_right',      on_action=bkt.Callback(arrange_adv_size.arrange_right_at_left, shapes=True),      get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_fill_right',         label="Right to left",  show_label=False, screentip='Fill right gap',  supertip='Alignment of the right edge to the left edge of the reference shape.'),
            bkt.ribbon.Button(id='arrange_fill_bottom',     on_action=bkt.Callback(arrange_adv_size.arrange_bottom_at_top, shapes=True),      get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_fill_bottom',        label="Bottom to top",    show_label=False, screentip='Fill top gap',  supertip='Alignment of the bottom edge to the top edge of the reference shape.'),
            bkt.ribbon.Button(id='arrange_fill_top',        on_action=bkt.Callback(arrange_adv_size.arrange_top_at_bottom, shapes=True),      get_enabled=bkt.apps.ppt_shapes_or_text_selected, image='arrange_fill_top',           label="Top to bottom",    show_label=False, screentip='Fill bottom gap',  supertip='Alignment of the top edge to the bottom edge of the reference shape.'),
        ]),

    ]
)

UiGroups.arrange_adv_easy_group = get_advanced_easy


