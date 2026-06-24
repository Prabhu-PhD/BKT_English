# -*- coding: utf-8 -*-
'''
Created on 19.01.2023

'''

import bkt

import bkt.library.algorithms as algos
import bkt.library.powerpoint as pplib


from .linkshapes import LinkedShapes
from bkt.helpers import memoize

class RepositionGallery(pplib.PositionGallery):
    
    def __init__(self, **kwargs):
        super(RepositionGallery, self).__init__(
            on_position_change = bkt.Callback(self.on_position_change),
            **kwargs
        )
    
    def get_item_supertip(self, index):
        return 'Positioniere die ausgewählten Shapes an der angezeigten Position/Größe.\nNur Position ändern [STRG],\nNur Größe ändern [SHIFT]'
    
    def on_position_change(self, target_frame, selection, shapes):
        if len(shapes) > 1:
            shape = selection.ShapeRange.Group()
            self.change_shape_position(shape, target_frame)
            shape.Ungroup().Select()
        else:
            self.change_shape_position(shapes[0], target_frame)
    
    def change_shape_position(self, shape, target_frame):
        # position shape
        # CTRL = position only
        # SHIFT = size only
        
        if not bkt.get_key_state(bkt.KeyCodes.SHIFT):
            shape.left   = target_frame.left
            shape.top    = target_frame.top
        
        if not bkt.get_key_state(bkt.KeyCodes.CTRL):
            shape.width  = target_frame.width
            shape.height = target_frame.height

reposition_gallery = RepositionGallery(id="positions", label="Position shapes")


class ChartShapes(object):
    chart_dimensions = [None, None] #height, width
    plotarea_dimensions = [None, None, None, None] #top, left, height, width

    @classmethod
    def is_chart_shape(cls, shape):
        try:
            # HasChart throws NotImplementedError for SmartArts
            return shape.HasChart == -1
        except NotImplementedError:
            return False
        # return shape.Type == pplib.MsoShapeType['msoChart'] or shape.Type == pplib.MsoShapeType['msoDiagram']

    @classmethod
    def is_paste_enabled(cls, shapes):
        return cls.chart_dimensions[0] is not None and all(cls.is_chart_shape(shape) for shape in shapes)

    @classmethod
    def copy_dimensions(cls, shape):
        plotarea = shape.Chart.PlotArea
        cls.chart_dimensions = [shape.Height, shape.Width]
        cls.plotarea_dimensions = [plotarea.Top, plotarea.Left, plotarea.Height, plotarea.Width]

    @classmethod
    def paste_dimensions(cls, shapes):
        for shape in shapes:
            plotarea = shape.Chart.PlotArea
            shape.Height, shape.Width = cls.chart_dimensions
            plotarea.Top, plotarea.Left, plotarea.Height, plotarea.Width = cls.plotarea_dimensions


class GroupsMore(object):
    @staticmethod
    def add_into_group(shapes):
        if shapes[0].Type == pplib.MsoShapeType['msoGroup']:
            master = pplib.GroupManager(shapes.pop(0))
            master.add_child_items(shapes).select()
        elif shapes[-1].Type == pplib.MsoShapeType['msoGroup']:
            master = pplib.GroupManager(shapes.pop(-1))
            master.add_child_items(shapes).select()
        else:
            pplib.shapes_to_range(shapes).group().select()

    @staticmethod
    def remove_from_group(shapes):
        master = pplib.GroupManager(shapes[0].ParentGroup)
        master.remove_child_items(shapes)
        pplib.shapes_to_range(shapes).select()
    
    @classmethod
    def visible_add_into_group(cls, shapes):
        return len(shapes) > 1 and cls.contains_group(shapes)

    @staticmethod
    def visible_remove_from_group(shapes):
        return all(pplib.shape_is_group_child(shape) for shape in shapes)
    
    @staticmethod
    def contains_group(shapes):
        return any(shp.Type == pplib.MsoShapeType['msoGroup'] for shp in shapes)

    @staticmethod
    def is_or_within_group(shape):
        return shape.Type == pplib.MsoShapeType['msoGroup'] or pplib.shape_is_group_child(shape)

    @staticmethod
    def recursive_ungroup(shapes):
        for shape in shapes:
            if shape.Type == pplib.MsoShapeType['msoGroup']:
                grp = pplib.GroupManager(shape)
                grp.recursive_ungroup().select(False)

    @staticmethod
    def select_all_groupitems(shape):
        if shape.Type == pplib.MsoShapeType['msoGroup']:
            all_shapes = list(iter(shape.GroupItems))
        elif pplib.shape_is_group_child(shape):
            all_shapes = list(iter(shape.ParentGroup.GroupItems))
        else:
            return
        pplib.shapes_to_range(all_shapes).select()


    @classmethod
    def auto_group(cls, shapes):
        shapes = pplib.wrap_shapes(sorted(shapes, key=lambda s: s.ZOrderPosition))
        processed_shapes = set()
        groups = []
        for child in shapes:
            if child in processed_shapes:
                continue
            within_shapes = [
                s for s in shapes
                if s != child and s not in processed_shapes and
                cls._is_shape_within(child, s)
            ]
            if len(within_shapes) > 0:
                within_shapes.append(child)
                groups.append(pplib.shapes_to_range(within_shapes).group())
                processed_shapes.update(within_shapes)
        if len(groups) > 0:
            pplib.shapes_to_range(groups).select()

    @classmethod
    def _is_shape_within(cls, outer_s, inner_s):
        #test if center point of inner_s is within bounds of outer_s
        return inner_s.width<=outer_s.width and inner_s.height<=outer_s.height and outer_s.x <= inner_s.center_x <= outer_s.x1 and outer_s.y <= inner_s.center_y <= outer_s.y1


class ArrangeCenter(object):
    @staticmethod
    def shape_in_center(center_shape, around_shapes):
        midpoint = algos.mid_point_shapes(around_shapes)
        center_shape.left = midpoint[0] - center_shape.width/2.0
        center_shape.top = midpoint[1] - center_shape.height/2.0
    
    @classmethod
    def arrange_shapes(cls, shapes):
        cls.shape_in_center(shapes.pop(-1), shapes)


class PictureFormat(object):
    shape_dimensions = [None, None] #ShapeHeight, ShapeWidth #, ShapeTop, ShapeLeft
    pic_dimensions   = [None, None, None, None] #PictureHeight, PictureWidth, PictureOffsetX, PictureOffsetY

    @classmethod
    def is_pic_shape(cls, shape):
        try:
            return shape.Type == pplib.MsoShapeType["msoPicture"]
        except:
            return False

    @classmethod
    def is_paste_enabled(cls, shapes):
        return cls.shape_dimensions[0] is not None and all(cls.is_pic_shape(shape) for shape in shapes)

    @classmethod
    def copy_dimensions(cls, shape):
        croparea = shape.PictureFormat.crop
        cls.shape_dimensions = [croparea.ShapeHeight, croparea.ShapeWidth] #, croparea.ShapeTop, croparea.ShapeLeft
        cls.pic_dimensions   = [croparea.PictureHeight, croparea.PictureWidth, croparea.PictureOffsetX, croparea.PictureOffsetY]

    @classmethod
    def paste_dimensions(cls, shapes):
        for shape in shapes:
            croparea = shape.PictureFormat.crop
            croparea.ShapeHeight, croparea.ShapeWidth = cls.shape_dimensions
            croparea.PictureHeight, croparea.PictureWidth, croparea.PictureOffsetX, croparea.PictureOffsetY = cls.pic_dimensions


class TableFormat(object):
    col_widths = []
    row_heights = []

    @classmethod
    def is_table_shape(cls, shape):
        try:
            # return shape.Type == pplib.MsoShapeType["msoTable"]
            return shape.HasTable == -1 #also covers tables in placeholders
        except:
            return False

    @classmethod
    def is_paste_enabled(cls, shapes):
        return len(cls.col_widths) > 0 and all(cls.is_table_shape(shape) for shape in shapes)

    @classmethod
    def copy_dimensions(cls, shape):
        cls.col_widths = [col.width for col in shape.table.columns]
        cls.row_heights = [row.height for row in shape.table.rows]

    @classmethod
    def paste_dimensions(cls, shapes):
        for shape in shapes:
            for i, col_width in enumerate(cls.col_widths):
                if shape.table.columns.count-1 < i:
                    break
                shape.table.columns(i+1).width = col_width
            
            for i, row_height in enumerate(cls.row_heights):
                if shape.table.rows.count-1 < i:
                    break
                shape.table.rows(i+1).height = row_height


class EdgeAutoFixer(object):
    threshold  = bkt.settings.get("toolbox.autofixer_threshold", 0.3)
    groupitems = bkt.settings.get("toolbox.autofixer_groupitems", True)
    order_key  = bkt.settings.get("toolbox.autofixer_order_key", "diagonal-down")

    @classmethod
    def settings_setter(cls, name, value):
        bkt.settings["toolbox.autofixer_"+name] = value
        setattr(cls, name, value)

    @classmethod
    def _iterate_all_shapes(cls, shapes, groupitems=True):
        for shape in shapes:
            #shapes that are rotated other than 0, 90, 180 or 270 degree are excluded
            if shape.rotation % 90 != 0:
                continue
            #connected connectors should not be moved
            if shape.Connector and (shape.ConnectorFormat.BeginConnected or shape.ConnectorFormat.EndConnected):
                continue
            
            if groupitems and shape.Type == 6: #pplib.MsoShapeType['msoGroup']
                for gShape in shape.GroupItems:
                    yield gShape
            else:
                yield shape
    
    @classmethod
    def get_image(cls, context):
        if cls.order_key == "diagonal-down":
            return context.python_addin.load_image("autofixer_dd")
        elif cls.order_key == "top-down":
            return context.python_addin.load_image("autofixer_td")
        else:
            return context.python_addin.load_image("autofixer_lr")

    @classmethod
    def autofix_edges_diagonal_down(cls, shapes):
        cls.settings_setter("order_key", "diagonal-down")
        cls.autofix_edges(shapes)
    
    @classmethod
    def autofix_edges_left_right(cls, shapes):
        cls.settings_setter("order_key", "left-right")
        cls.autofix_edges(shapes)
    
    @classmethod
    def autofix_edges_top_down(cls, shapes):
        cls.settings_setter("order_key", "top-down")
        cls.autofix_edges(shapes)

    @classmethod
    def autofix_edges(cls, shapes):
        cls._autofix_edges(shapes, pplib.cm_to_pt(cls.threshold), cls.groupitems, cls.order_key)
    
    @classmethod
    def _autofix_edges(cls, shapes, threshold=None, groupitems=True, order_key="diagonal-down"):
        #TODO: how to handle locked aspect-ratio and autosize? rotated shapes? ojects with 0 height/width? exclude placeholders?

        threshold = threshold or pplib.cm_to_pt(0.3)

        shapes = pplib.wrap_shapes(cls._iterate_all_shapes(shapes, groupitems))

        # shapes.sort(key=lambda shape: (shape.left, shape.top))
        order_keys = {
            "diagonal-down": [lambda shape: shape.visual_x+shape.visual_y, False],
            "diagonal-up":   [lambda shape: shape.visual_x+shape.visual_y, True],
            "left-right": [lambda shape: (shape.visual_x,shape.visual_y), False],
            "top-down":   [lambda shape: (shape.visual_y,shape.visual_x), False],
            "right-left": [lambda shape: (shape.visual_x,shape.visual_y), True],
            "bottom-up":  [lambda shape: (shape.visual_y,shape.visual_x), True],
        }
        shapes.sort(key=order_keys[order_key][0], reverse=order_keys[order_key][1])

        # logging.debug("Autofix: top-left")
        child_shapes = shapes[:]
        for master_shape in shapes:
            child_shapes.remove(master_shape)
            
            for shape in child_shapes:
                # logging.debug("Autofix1: %s x %s", master_shape.name, shape.name)

                #save values before moving shape
                # visual_x1, visual_y1 = shape.visual_x1, shape.visual_y1

                if 1e-4 < abs(shape.visual_x-master_shape.visual_x) < threshold:
                    #resize to left edge
                    delta = master_shape.visual_x - shape.visual_x
                    shape.visual_x += delta
                    shape.visual_width -= delta

                if 1e-4 < abs(shape.visual_y-master_shape.visual_y) < threshold:
                    #resize to top edge
                    delta = master_shape.visual_y - shape.visual_y
                    shape.visual_y += delta
                    shape.visual_height -= delta

                if 1e-4 < abs(shape.visual_x1-master_shape.visual_x1) < threshold:
                    #resize to right edge
                    shape.visual_width = master_shape.visual_x1-shape.visual_x

                if 1e-4 < abs(shape.visual_y1-master_shape.visual_y1) < threshold:
                    #resize to bottom edge
                    shape.visual_height = master_shape.visual_y1-shape.visual_y


arrange_menu = memoize(lambda: bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
                # bkt.ribbon.MenuSeparator(title="Rotation"),
                # bkt.mso.button.ObjectRotateRight90,
                # bkt.mso.button.ObjectRotateLeft90,
                # bkt.mso.button.ObjectFlipHorizontal,
                # bkt.mso.button.ObjectFlipVertical,
                # bkt.ribbon.MenuSeparator(),
                # bkt.mso.button.ObjectRotationOptionsDialog,
                bkt.ribbon.MenuSeparator(title="Position"),
                bkt.mso.control.ObjectRotateGallery,
                reposition_gallery,
                bkt.ribbon.Menu(
                    label='Transfer dimensions/sizes',
                    supertip="Copy dimensions of charts, image crops and tables from one object to another",
                    image_mso='PasteWithColumnWidths',
                    children=[
                        bkt.ribbon.Button(
                            id = 'chart_dimensions_copy',
                            label="Copy chart dimensions",
                            image_mso="ChartPlotArea",
                            screentip="Copy size and position of the chart area",
                            supertip="Copies the height and width of the chart as well as the size and position of the plot area to match another chart.",
                            on_action=bkt.Callback(ChartShapes.copy_dimensions, shape=True),
                            get_enabled = bkt.Callback(ChartShapes.is_chart_shape, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'chart_dimensions_paste',
                            label="Insert chart dimensions",
                            image_mso="PasteWithColumnWidths",
                            screentip="Insert size and position of the chart area",
                            supertip="Transfers the copied size and position of the chart or plot area to the selected chart.",
                            on_action=bkt.Callback(ChartShapes.paste_dimensions, shapes=True),
                            get_enabled = bkt.Callback(ChartShapes.is_paste_enabled, shapes=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'pic_crop_copy',
                            label="Copy image crop",
                            image_mso="PictureCrop",
                            screentip="Copy size and position of the image crop",
                            supertip="Copies the height and width of the crop of a cropped image to match the crop with another image.",
                            on_action=bkt.Callback(PictureFormat.copy_dimensions, shape=True),
                            get_enabled = bkt.Callback(PictureFormat.is_pic_shape, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'pic_crop_paste',
                            label="Insert image crop",
                            image_mso="PasteWithColumnWidths",
                            screentip="Insert size and position of the image crop",
                            supertip="Transfers the copied size and position of the image crop to the selected image.",
                            on_action=bkt.Callback(PictureFormat.paste_dimensions, shapes=True),
                            get_enabled = bkt.Callback(PictureFormat.is_paste_enabled, shapes=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        bkt.ribbon.Button(
                            id = 'table_dimensions_copy',
                            label="Copy table sizes",
                            image_mso="TableColumnsDistribute",
                            screentip="Copy width/height of table columns/rows",
                            supertip="Copies the height and width of the table rows or table columns to match them with another table.",
                            on_action=bkt.Callback(TableFormat.copy_dimensions, shape=True),
                            get_enabled = bkt.Callback(TableFormat.is_table_shape, shape=True),
                        ),
                        bkt.ribbon.Button(
                            id = 'table_dimensions_paste',
                            label="Insert table sizes",
                            image_mso="PasteWithColumnWidths",
                            screentip="Insert width/height of table columns/rows",
                            supertip="Transfers the copied table dimensions to the selected table.",
                            on_action=bkt.Callback(TableFormat.paste_dimensions, shapes=True),
                            get_enabled = bkt.Callback(TableFormat.is_paste_enabled, shapes=True),
                        ),
                    ]
                ),
                bkt.ribbon.SplitButton(
                    id = 'edge_autofixer_splitbutton',
                    children=[
                        bkt.ribbon.Button(
                            id = 'edge_autofixer',
                            label="Edge auto-fixer",
                            # image_mso='GridSettings',
                            get_image=bkt.Callback(EdgeAutoFixer.get_image, context=True),
                            supertip="Compensates for minimal edge shifts of the selected shapes.",
                            on_action=bkt.Callback(EdgeAutoFixer.autofix_edges, shapes=True),
                            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                        ),
                        bkt.ribbon.Menu(
                            label="Edge auto-fixer menu",
                            supertip="Setting options for the edge auto-fixer",
                            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                            children=[
                                bkt.ribbon.Button(
                                    id = 'edge_autofixer-dd',
                                    label="Edge auto-fixer diagonally from top-left",
                                    image='autofixer_dd',
                                    supertip="Compensates for minimal edge shifts of the selected shapes by enlarging onto shapes to the upper-left of the shapes to be adjusted.",
                                    on_action=bkt.Callback(EdgeAutoFixer.autofix_edges_diagonal_down, shapes=True),
                                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                                ),
                                bkt.ribbon.Button(
                                    id = 'edge_autofixer-td',
                                    label="Edge auto-fixer from top to bottom",
                                    image='autofixer_td',
                                    supertip="Compensates for minimal edge shifts of the selected shapes by enlarging onto shapes to the left of the shapes to be adjusted.",
                                    on_action=bkt.Callback(EdgeAutoFixer.autofix_edges_top_down, shapes=True),
                                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                                ),
                                bkt.ribbon.Button(
                                    id = 'edge_autofixer-lr',
                                    label="Edge auto-fixer from left to right",
                                    image='autofixer_lr',
                                    supertip="Compensates for minimal edge shifts of the selected shapes by enlarging onto shapes above the shapes to be adjusted.",
                                    on_action=bkt.Callback(EdgeAutoFixer.autofix_edges_left_right, shapes=True),
                                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                                ),
                                bkt.ribbon.MenuSeparator(),
                                bkt.ribbon.ToggleButton(
                                    label="Adjust group elements individually",
                                    supertip="Specifies whether elements of a group are considered individually or the entire group as a whole.",
                                    get_pressed=bkt.Callback(lambda: EdgeAutoFixer.groupitems is True),
                                    on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("groupitems", pressed)),
                                ),
                                bkt.ribbon.Menu(
                                    label="Change tolerance",
                                    screentip="Edge auto-fixer tolerance",
                                    supertip="Adjust the threshold for the edge auto-fixer.",
                                    children=[
                                        bkt.ribbon.ToggleButton(
                                            label="Small 0.1 cm",
                                            screentip="Tolerance small 0.1 cm",
                                            supertip="Sets the edge auto-fixer tolerance to small = 0.1 cm",
                                            get_pressed=bkt.Callback(lambda: EdgeAutoFixer.threshold == 0.1),
                                            on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("threshold", 0.1)),
                                        ),
                                        bkt.ribbon.ToggleButton(
                                            label="Medium 0.3 cm",
                                            screentip="Tolerance medium 0.3 cm",
                                            supertip="Sets the edge auto-fixer tolerance to medium = 0.3 cm",
                                            get_pressed=bkt.Callback(lambda: EdgeAutoFixer.threshold == 0.3),
                                            on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("threshold", 0.3)),
                                        ),
                                        bkt.ribbon.ToggleButton(
                                            label="Large 1 cm",
                                            screentip="Tolerance large 1 cm",
                                            supertip="Sets the edge auto-fixer tolerance to large = 1 cm",
                                            get_pressed=bkt.Callback(lambda: EdgeAutoFixer.threshold == 1.0),
                                            on_toggle_action=bkt.Callback(lambda pressed: EdgeAutoFixer.settings_setter("threshold", 1.0)),
                                        ),
                                    ]
                                ),
                            ]
                        ),
                    ]
                ),
                bkt.ribbon.Button(
                    id = 'arrange_shape_center',
                    label="Position shape in the center",
                    image_mso="DiagramRadialInsertClassic",
                    screentip="Position last shape at the center",
                    supertip="Places the last selected shape in the weighted center of all other selected shapes.",
                    on_action=bkt.Callback(ArrangeCenter.arrange_shapes, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_min2_selected,
                ),
                bkt.ribbon.MenuSeparator(title="Grouping"),
                bkt.ribbon.Button(
                    id = 'auto_Group',
                    label="Auto grouping",
                    image_mso="ObjectsGroup",
                    screentip="Group related shapes",
                    supertip="Searches for related shapes/shape groups and groups them automatically.",
                    on_action=bkt.Callback(GroupsMore.auto_group, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_min2_selected,
                ),
                bkt.ribbon.Button(
                    id = 'add_into_group',
                    label="Insert into group",
                    image_mso="ObjectsRegroup",
                    screentip="Insert shapes into group",
                    supertip="If the first or last selected shape is a group, all other shapes are inserted into this group. Otherwise all shapes are grouped.",
                    on_action=bkt.Callback(GroupsMore.add_into_group, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_min2_selected,
                ),
                bkt.ribbon.Button(
                    id = 'recursive_ungroup',
                    label="Recursive ungroup",
                    image_mso="ObjectsUngroup",
                    screentip="Ungroup recursively",
                    supertip="Applies ungroup until all nested groups are dissolved.",
                    on_action=bkt.Callback(GroupsMore.recursive_ungroup, shapes=True),
                    get_enabled = bkt.Callback(GroupsMore.contains_group, shapes=True),
                ),
                bkt.ribbon.Button(
                    id = 'select_all_groupitems',
                    label="Select elements of the group",
                    image_mso="ObjectsMultiSelect",
                    screentip="Select all elements of the group",
                    supertip="Selects all elements within the group.",
                    on_action=bkt.Callback(GroupsMore.select_all_groupitems, shape=True),
                    get_enabled = bkt.Callback(GroupsMore.is_or_within_group, shape=True),
                ),
                bkt.ribbon.Button(
                    id = 'remove_from_group',
                    label="Detach from group",
                    image_mso="ObjectsUngroup",
                    screentip="Detach shapes from group",
                    supertip="The selected shapes are detached from the current group without changing the group.",
                    on_action=bkt.Callback(GroupsMore.remove_from_group, shapes=True),
                    get_enabled = bkt.Callback(GroupsMore.visible_remove_from_group, shapes=True),
                ),
                bkt.ribbon.MenuSeparator(title="Linked shapes"),
                bkt.ribbon.Button(
                    id = 'shape_copy_to_all',
                    label="Copy and link shape on following slides…",
                    image_mso="ShapesDuplicate",
                    screentip="Duplicate shape on following slides",
                    supertip="Duplicates the current shape onto all slides after the current slide and links them for future operations.",
                    on_action=bkt.Callback(LinkedShapes.copy_to_all),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'shape_find_similar_and_link',
                    label="Search and link shape on following slides…",
                    image_mso="FindTag",
                    screentip="Search for the same shape on following slides and link it",
                    supertip="Searches for the current shape on all slides after the current slide by position and size and links them together.",
                    on_action=bkt.Callback(LinkedShapes.find_similar_and_link),
                    get_enabled = bkt.apps.ppt_shapes_exactly1_selected,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'link_shapes',
                    label="Link selected shapes together",
                    image_mso="HyperlinkCreate",
                    screentip="Link all selected shapes together",
                    supertip="Link the selected shapes for future operations. The link is preserved when the shapes are copied.",
                    on_action=bkt.Callback(LinkedShapes.link_shapes, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'each_link_shapes',
                    label="Convert selected shapes individually into a link",
                    # image_mso="HyperlinkCreate",
                    screentip="Link all selected shapes individually",
                    supertip="The selected shapes each get an internal link ID. The link is preserved when the shapes are copied.",
                    on_action=bkt.Callback(LinkedShapes.each_link_shapes, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'extend_link_shapes',
                    label="Extend existing shape link",
                    # image_mso="HyperlinkCreate",
                    screentip="Extend existing shape link",
                    supertip="To extend the existing shape link, the internal ID is stored. Via 'Add selected shapes to the link', further shapes can then be added to the link.",
                    on_action=bkt.Callback(LinkedShapes.extend_link_shapes, shape=True),
                    get_enabled = bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
                ),
                bkt.ribbon.Button(
                    id = 'add_to_link_shapes',
                    label="Add selected shapes to the link",
                    # image_mso="HyperlinkCreate",
                    screentip="Add selected shapes to the link",
                    supertip="Add selected shapes to the stored ID. A new link must be created first, or an existing one extended.",
                    on_action=bkt.Callback(LinkedShapes.add_to_link_shapes, shapes=True),
                    get_enabled = bkt.Callback(LinkedShapes.enabled_add_linked_shapes),
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'unlink_shape',
                    label="Remove single shape link",
                    image_mso="HyperlinkRemove",
                    screentip="Remove the link of the selected shape",
                    supertip="Removes the link ID from the current shape. All other shapes with the same ID remain linked.",
                    on_action=bkt.Callback(LinkedShapes.unlink_shape, shape=True),
                    get_enabled = bkt.Callback(LinkedShapes.is_linked_shape),
                ),
                bkt.ribbon.Button(
                    id = 'unlink_all_shapes',
                    label="Dissolve entire shape link",
                    # image_mso="HyperlinkRemove",
                    screentip="Remove all shape links",
                    supertip="Removes the link ID from the current shape as well as all linked shapes with the same ID.",
                    on_action=bkt.Callback(LinkedShapes.unlink_all_shapes, shape=True, context=True),
                    get_enabled = bkt.Callback(LinkedShapes.is_linked_shape),
                ),

                # bkt.ribbon.MenuSeparator(),
                # bkt.ribbon.Menu(
                #     label='Linked shapes',
                #     image_mso='ControlAlignToGrid',
                #     screentip="Operations on linked shapes",
                #     supertip="Delete or align all linked shapes. Options are also available in the context menu of linked shapes.",
                #     get_enabled = bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
                #     children=[
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_count',
                #             label="Number of linked shapes",
                #             image_mso="FindDialog",
                #             screentip="Count linked shapes",
                #             supertip="Counts the number of linked shapes on all slides.",
                #             on_action=bkt.Callback(LinkedShapes.count_link_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_next',
                #             label="Find next linked shape",
                #             image_mso="FindNext",
                #             screentip="Go to the next linked shape",
                #             supertip="Searches for the next linked shape. If there is no further shape on the following slides, the first linked shape of the presentation is searched for.",
                #             on_action=bkt.Callback(LinkedShapes.goto_linked_shape, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.MenuSeparator(),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_delete',
                #             label="Delete others",
                #             image_mso="HyperlinkRemove",
                #             screentip="Delete linked shapes",
                #             supertip="Delete all linked shapes on all slides.",
                #             on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_replace',
                #             label="Replace others with this one",
                #             image_mso="HyperlinkCreate",
                #             screentip="Replace linked shapes",
                #             supertip="Replace all linked shapes on all slides with the selected shape.",
                #             on_action=bkt.Callback(LinkedShapes.replace_with_this, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.MenuSeparator(),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_align',
                #             label="Match position",
                #             image_mso="ControlAlignToGrid",
                #             screentip="Match position of linked shapes",
                #             supertip="Set the position and rotation of all linked shapes to the position of the selected shape.",
                #             on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_size',
                #             label="Match size",
                #             image_mso="SizeToControlHeightAndWidth",
                #             screentip="Match size of linked shapes",
                #             supertip="Set the size of all linked shapes to the size of the selected shape.",
                #             on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_format',
                #             label="Match formatting",
                #             image_mso="FormatPainter",
                #             screentip="Match formatting of linked shapes",
                #             supertip="Set the formatting of all linked shapes to the size of the selected shape.",
                #             on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_text',
                #             label="Match text",
                #             image_mso="TextBoxInsert",
                #             screentip="Match text of linked shapes",
                #             supertip="Set the text of all linked shapes to the size of the selected shape.",
                #             on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #         bkt.ribbon.MenuSeparator(),
                #         bkt.ribbon.Button(
                #             id = 'linked_shapes_all',
                #             label="Match everything",
                #             image_mso="GroupUpdate",
                #             screentip="Match all properties of linked shapes",
                #             supertip="Set all properties of all linked shapes like the selected shape.",
                #             on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shape=True, context=True),
                #             # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                #         ),
                #     ]
                # )
                ]
        ))