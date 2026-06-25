# -*- coding: utf-8 -*-
'''
Created on 2017-08-16
@author: Florian Stallmann
'''



#for caching
# import time

import bkt
# import bkt.library.powerpoint as pplib

from bkt.library.powerpoint import PositionGallery, pt_to_cm, LocPin, LocpinGallery
from bkt.library.table import TableRecognition

# pt_to_cm_factor = 2.54 / 72;
# def pt_to_cm(pt):
#     return float(pt) * pt_to_cm_factor;
# def cm_to_pt(cm):
#     return float(cm) / pt_to_cm_factor;


class ShapeTables(object):
    resize_cells = False
    fit_cells = False
    alignment_locpin = LocPin()
    # alignment_horizontal = "left"
    # alignment_vertical = "top"
    equal_spacing = False

    # contentarea = {"left": None, "top": None, "width": None, "height": None}
    tr_cache = None
    last_time_tr_cache_changed = 0

    def __init__(self):
        self.position_gallery = PositionGallery(
            label="Fit table into area",
            description="Adjust shape table to the size of the chosen area",
            on_position_change = bkt.Callback(self.table_contentarea_fit),
            get_item_supertip = bkt.Callback(self.get_item_supertip)
        )
        self.locpin_gallery = LocpinGallery(
            id="table_alignment",
            label="Shape alignment",
            image_mso="ObjectAlignMenu",
            supertip="Sets the alignment of the shapes within the table cells.",
            item_height="32",
            item_width="32",
            locpin=self.alignment_locpin,
            item_supertip="Shapes are arranged {} in table cells when arranging shapes.",
        )
    
    @property
    def alignment_horizontal(self):
        return ["left", "center", "right"][self.alignment_locpin.fixation[1]-1]
    @property
    def alignment_vertical(self):
        return ["top", "middle", "bottom"][self.alignment_locpin.fixation[0]-1]

    def get_item_supertip(self, index):
        return 'The selected shapes are fitted into the area as a table.'

    def clear_cache(self):
        self.tr_cache = None

    def _prepare_table(self, shapes, force=False):
        # Run table recognition only one time in 500ms
        #NOTE: timing-based cache not compatible with comrelease
        if force or self.tr_cache is None: #or time.time() - self.last_time_tr_cache_changed > 0.5:
            self.tr_cache = TableRecognition(shapes)
            self.tr_cache.run()
            # self.last_time_tr_cache_changed = time.time()
        return self.tr_cache


    ### ALIGN TABLE / do not consider resize-option ###

    def align_table(self, shapes):
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            self.align_table_zero(shapes)
        else:
            tr = self._prepare_table(shapes, force=True)
            spac_rows = max(0,tr.min_spacing_rows(max_rows=2))
            spac_cols = max(0,tr.min_spacing_cols(max_cols=2))
            if self.equal_spacing:
                spacing = (spac_rows+spac_cols)/2.0
            else:
                spacing = (spac_rows, spac_cols)
            tr.align(spacing=spacing, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def align_table_default(self, shapes):
        tr = self._prepare_table(shapes, force=True)
        tr.align(fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def align_table_median(self, shapes):
        tr = self._prepare_table(shapes, force=True)
        tr.align(tr.median_spacing(), fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def align_table_zero(self, shapes):
        tr = self._prepare_table(shapes, force=True)
        tr.align(0, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)


    ### SPACING / consider resize option ###

    # def get_spacing(self, shapes):
    #     tr = self._prepare_table(shapes)
    #     res = tr.median_spacing()
    #     return res
    #     # return round(pt_to_cm(res), 2)

    # def set_spacing(self, shapes, value):
    #     # if type(value) == str:
    #     #     value = float(value.replace(',', '.'))
    #     # value = max(0,cm_to_pt(value))
    #     value = max(0, value)

    #     tr = self._prepare_table(shapes)
    #     if self.resize_cells:
    #         bounds = tr.get_bounds()
    #         tr.fit_content(*bounds, spacing=value, fit_cells=self.fit_cells)
    #     else:
    #         tr.align(spacing=value, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)


    def get_resize_cells(self):
        if bkt.get_key_state(bkt.KeyCodes.ALT):
            return not self.resize_cells
        else:
            return self.resize_cells

    def enabled_spacing_rows(self, shapes):
        if len(shapes) < 2:
            return False
        tr = self._prepare_table(shapes)
        return tr.dimension[0] > 1

    def get_spacing_rows(self, shapes):
        tr = self._prepare_table(shapes)
        res = tr.min_spacing_rows(max_rows=2)
        return res
        # return round(pt_to_cm(res), 2)

    def set_spacing_rows(self, shapes, value):
        # if type(value) == str:
        #     value = float(value.replace(',', '.'))
        # value = max(0,cm_to_pt(value))
        # value = max(0, value)

        if self.equal_spacing:
            spacing = value
        else:
            spacing = (value, None)

        tr = self._prepare_table(shapes, force=True)
        if self.get_resize_cells():
            bounds = tr.get_bounds()
            tr.fit_content(*bounds, spacing=spacing, fit_cells=self.fit_cells)
        else:
            tr.align(spacing=spacing, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def enabled_spacing_cols(self, shapes):
        if len(shapes) < 2:
            return False
        tr = self._prepare_table(shapes)
        return tr.dimension[1] > 1

    def get_spacing_cols(self, shapes):
        tr = self._prepare_table(shapes)
        res = tr.min_spacing_cols(max_cols=2)
        return res
        # return round(pt_to_cm(res), 2)

    def set_spacing_cols(self, shapes, value):
        # if type(value) == str:
        #     value = float(value.replace(',', '.'))
        # value = max(0,cm_to_pt(value))
        # value = max(0, value)

        if self.equal_spacing:
            spacing = value
        else:
            spacing = (None, value)

        tr = self._prepare_table(shapes, force=True)
        if self.get_resize_cells():
            bounds = tr.get_bounds()
            tr.fit_content(*bounds, spacing=spacing, fit_cells=self.fit_cells)
        else:
            tr.align(spacing=spacing, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)


    ### SPECIAL FUNCTIONS / separate method for resizing ###

    def table_transpose_destructive(self,shapes):
        tr = self._prepare_table(shapes, force=True)
        spacing = tr.median_spacing()
        tr.transpose()
        tr.align(spacing=spacing, fit_cells=self.fit_cells, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def table_transpose_in_bounds(self,shapes):
        tr = self._prepare_table(shapes, force=True)
        spacing = tr.median_spacing()
        bounds = tr.get_bounds()
        tr.transpose()
        tr.transpose_cell_size()
        tr.fit_content(*bounds, spacing=spacing, fit_cells=self.fit_cells)

    def table_info(self,shapes):
        tr = self._prepare_table(shapes, force=True)
        msg = ""
        msg += "Table size: rows=%d, columns=%d" % tr.dimension
        msg += "Median-Abstand: %s cm" % round(pt_to_cm(tr.median_spacing()),2)
        bkt.message(msg)

    def table_info_desc(self,context):
        try:
            if context.selection.Type != 2 or context.selection.ShapeRange.Count < 2:
                raise ValueError("invalid selection for table")

            # shapes = pplib.get_shapes_from_selection(selection)
            tr = self._prepare_table(context.shapes)
            return "Tabelle: %d Z. \xd7 %d S." % tr.dimension #Zeilen x Spalten
        except:
            return "Tabelle: -"

    ### FIT CELLS ###

    def table_fit_cells_destructive(self,shapes):
        tr = self._prepare_table(shapes, force=True)
        spacing = tr.median_spacing()
        tr.align(spacing=spacing, fit_cells=True, align_x=self.alignment_horizontal, align_y=self.alignment_vertical)

    def table_fit_cells_in_bounds(self,shapes):
        tr = self._prepare_table(shapes, force=True)
        spacing = tr.median_spacing()
        bounds = tr.get_bounds()
        tr.fit_content(*bounds, spacing=spacing, fit_cells=True)

    def table_contentarea_fit(self,target_frame,shapes):
        if len(shapes) < 2:
            return
        tr = self._prepare_table(shapes, force=True)
        spacing = tr.median_spacing()
        tr.fit_content(target_frame.left, target_frame.top, target_frame.width, target_frame.height, spacing)

    ### DISTRIBUTE CELLS ###

    def table_distribute_cols(self, shapes):
        tr = self._prepare_table(shapes, force=True)
        spacing = tr.min_spacing_cols()
        bounds = tr.get_bounds()
        tr.fit_content(*bounds, spacing=(None, spacing), fit_cells=True) #equalize spacing in first run
        tr.fit_content(*bounds, spacing=(None, spacing), fit_cells=True, distribute_cols=True) #distribute in second run

    def table_distribute_rows(self, shapes):
        tr = self._prepare_table(shapes, force=True)
        spacing = tr.min_spacing_rows()
        bounds = tr.get_bounds()
        tr.fit_content(*bounds, spacing=(spacing, None), fit_cells=True) #equalize spacing in first run
        tr.fit_content(*bounds, spacing=(spacing, None), fit_cells=True, distribute_rows=True) #distribute in second run


shape_tables = ShapeTables()

bkt.AppEvents.selection_changed       += bkt.Callback(shape_tables.clear_cache)


tabellen_gruppe = bkt.ribbon.Group(
    id="bkt_shapetables_group",
    label='Table from shapes',
    supertip="Enables the table-shaped arrangement of shapes. The `ppt_shapetables` feature must be installed.",
    image='align_table',
    children = [
        bkt.ribbon.SplitButton(
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
            size="large",
            children=[
                bkt.ribbon.Button(
                    id = 'align_table',
                    label="Align as table",
                    show_label=True,
                    # size="large",
                    image='align_table',
                    screentip="Align table (auto)",
                    supertip="Aligns the selected shapes as a table with calculated row and column spacing. With the SHIFT key, spacing is set to 0.",
                    on_action=bkt.Callback(shape_tables.align_table, shapes=True, shapes_min=2),
                    # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(label="Menu for aligning as a table", supertip="Align shape table with various spacing options", item_size="large", children=[
                    bkt.ribbon.Button(
                        id = 'align_table2',
                        label="Align as table (standard spacing)",
                        description="Align shapes as a table with 0.35 cm spacing",
                        # show_label=True,
                        image='align_table',
                        supertip="Aligns the selected shapes as a table with a standard spacing of 0.35 cm (10 pt)",
                        on_action=bkt.Callback(shape_tables.align_table_default, shapes=True, shapes_min=2),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'align_table_median',
                        label="Align as table (median spacing)",
                        description="Align shapes as a table with median spacing",
                        # show_label=True,
                        image='align_table_median',
                        supertip="Aligns the selected shapes as a table with the median spacing of the shapes.",
                        on_action=bkt.Callback(shape_tables.align_table_median, shapes=True, shapes_min=2),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'align_table_zero',
                        label="Align as table without spacing [Shift]",
                        description="Align shapes as a table without spacing.",
                        # show_label=True,
                        image='align_table_zero',
                        supertip="Aligns the selected shapes as a table without spacing.",
                        on_action=bkt.Callback(shape_tables.align_table_zero, shapes=True, shapes_min=2),
                        # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    # bkt.ribbon.MenuSeparator(),
                    # bkt.ribbon.Button(
                    #     id = 'table_fit_cells',
                    #     label="Fit shapes into cells",
                    #     description="Sets the size of the shapes to the size of the table cells",
                    #     # show_label=True,
                    #     image='table_fit_cells',
                    #     supertip="Sets the shape size to the size of the table cell. Depending on the chosen mode, the shapes are moved or enlarged/shrunk.",
                    #     on_action=bkt.Callback(shape_tables.table_fit_cells, shapes=True, shapes_min=2),
                    #     # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    # ),
                ])
            ]
        ),
        shape_tables.locpin_gallery,
        bkt.ribbon.Menu(
            image='shape_table_transpose',
            label="Adjust table",
            supertip="Adjust or transpose shape table in various ways",
            item_size="large",
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
            children=[
                bkt.ribbon.Menu(
                    image='table_fit_cells',
                    label="Fit shapes into cells",
                    description="Sets the size of the shapes to the size of the table cells.",
                    item_size="large",
                    children=[
                        bkt.ribbon.Button(
                            id = 'table_fit_cells_in_bounds',
                            label="Fit with equal table size",
                            description="Sets the size of the shapes to the size of the table cells and keeps the current table size",
                            # show_label=True,
                            image='table_fit_cells_1',
                            supertip="Sets the shape size to the size of the table cell. The shapes are enlarged or shrunk so as not to change the current table size.",
                            on_action=bkt.Callback(shape_tables.table_fit_cells_in_bounds, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'table_fit_cells_destructive',
                            label="Fit with adjusted table size",
                            description="Sets the size of the shapes to the size of the table cells and changes the current table size",
                            # show_label=True,
                            image='table_fit_cells_2',
                            supertip="Sets the shape size to the size of the table cell. The shapes are moved, thereby changing the current table size.",
                            on_action=bkt.Callback(shape_tables.table_fit_cells_destructive, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    image_mso='TableColumnsDistribute',
                    label="Match cell sizes",
                    description="Normalizes the width or height of the cells.",
                    item_size="large",
                    children=[
                        bkt.ribbon.Button(
                            id = 'table_distribute_cols',
                            label="Distribute column width",
                            description="Distribute the column widths evenly",
                            # show_label=True,
                            image_mso='TableColumnsDistribute',
                            supertip="Normalizes the width of all columns without changing the table size.",
                            on_action=bkt.Callback(shape_tables.table_distribute_cols, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'table_distribute_rows',
                            label="Distribute row height",
                            description="Distribute the row heights evenly",
                            # show_label=True,
                            image_mso='TableRowsDistribute',
                            supertip="Normalizes the height of all rows without changing the table size.",
                            on_action=bkt.Callback(shape_tables.table_distribute_rows, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    image='shape_table_transpose',
                    label="Transpose table",
                    description="Transposes (i.e. mirrors) the table.",
                    item_size="large",
                    children=[
                        bkt.ribbon.Button(
                            id = 'table_transpose_in_bounds',
                            label="Transpose table with equal table size",
                            description="Transpose shape table without changing the table size",
                            # show_label=True,
                            image='shape_table_transpose_1',
                            supertip="Transposes the table, i.e. mirrors the cells along the main diagonal. The cell size is also changed so as not to change the table size.",
                            on_action=bkt.Callback(shape_tables.table_transpose_in_bounds, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                        bkt.ribbon.Button(
                            id = 'table_transpose_destructive',
                            label="Transpose table with adjusted table size",
                            description="Transpose shape table without changing the cell sizes",
                            # show_label=True,
                            image='shape_table_transpose_2',
                            supertip="Transposes the table, i.e. mirrors the cells along the main diagonal. The cell size is not changed, only the table size.",
                            on_action=bkt.Callback(shape_tables.table_transpose_destructive, shapes=True, shapes_min=2),
                            # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                        ),
                    ]
                ),
                shape_tables.position_gallery,
            ]
        ),
        bkt.ribbon.Button(
            id = 'table_info',
            # label="Show table info",
            get_label=bkt.Callback(shape_tables.table_info_desc, context=True),
            # get_description=bkt.Callback(shape_tables.table_info_desc, shapes=True, shapes_min=2),
            # description="Show information about the detected table",
            # show_label=True,
            image='shape_table_info',
            screentip="Table information",
            supertip="Shows information about the table. Useful for finding out beforehand whether a table is correctly detected.",
            on_action=bkt.Callback(shape_tables.table_info, shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        # bkt.ribbon.Menu(
        #     label=u"More table features",
        #     image="shape_table_info",
        #     item_size="large",
        #     children = [
        #         # bkt.ribbon.Button(
        #         #     id = 'table_info',
        #         #     label="Show table info",
        #         #     get_description=bkt.Callback(shape_tables.table_info_desc, shapes=True, shapes_min=2),
        #         #     # description="Show information about the detected table",
        #         #     # show_label=True,
        #         #     image='shape_table_info',
        #         #     supertip="Shows information about the table. Useful for finding out beforehand whether a table is correctly detected.",
        #         #     on_action=bkt.Callback(shape_tables.table_info, shapes=True, shapes_min=2),
        #         #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #         # ),
        #         bkt.ribbon.Button(
        #             id = 'table_transpose',
        #             label="Transpose table",
        #             description="Transpose shape table (i.e. mirror)",
        #             # show_label=True,
        #             image='shape_table_transpose',
        #             supertip="Transposes the table, i.e. mirrors the cells along the main diagonal.",
        #             on_action=bkt.Callback(shape_tables.table_transpose, shapes=True, shapes_min=2),
        #             get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #         ),
        #         bkt.ribbon.MenuSeparator(),
        #         shape_tables.position_gallery,
        #         # bkt.ribbon.Button(
        #         #     id = 'table_contentarea_set',
        #         #     label="Define content area",
        #         #     show_label=True,
        #         #     image='table_contentarea_set',
        #         #     supertip="Defines the position and size of the selected shape as the content area to fit tables into this area. The selected shape is then deleted.",
        #         #     on_action=bkt.Callback(shape_tables.table_contentarea_set, presentation=True, shape=True),
        #         #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #         # ),
        #         # bkt.ribbon.Button(
        #         #     id = 'table_contentarea_unset',
        #         #     label="Reset content area",
        #         #     show_label=True,
        #         #     image='table_contentarea_unset',
        #         #     supertip="Deletes the defined content area, i.e. the slide's default content area is used again.",
        #         #     on_action=bkt.Callback(shape_tables.table_contentarea_unset, presentation=True),
        #         #     get_enabled = bkt.Callback(shape_tables.table_contentarea_defined, presentation=True)
        #         # ),
        #         # bkt.ribbon.Button(
        #         #     id = 'table_contentarea_fit',
        #         #     label="Fit table into content area",
        #         #     show_label=True,
        #         #     image='table_contentarea_fit',
        #         #     supertip="Fits the selected shapes as a table into the dimensions of the content area. The content area can be defined beforehand, otherwise the slide's default content area is used.",
        #         #     on_action=bkt.Callback(shape_tables.table_contentarea_fit, presentation=True, shapes=True, shapes_min=2),
        #         #     #get_enabled = bkt.Callback(shape_tables.table_contentarea_enabled, presentation=True, shapes=True)
        #         #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #         # ),
        #     ]
        # ),
        bkt.ribbon.Separator(),
        # bkt.ribbon.RoundingSpinnerBox(
        #     id = 'align_table_spacing',
        #     label=u"Distance",
        #     show_label=False,
        #     image="align_table_spacing",
        #     supertip="Changes the spacing of the shapes. Depending on the chosen mode, the shapes are moved or enlarged/shrunk.",
        #     on_change = bkt.Callback(shape_tables.set_spacing, shapes=True, shapes_min=2),
        #     get_text  = bkt.Callback(shape_tables.get_spacing, shapes=True, shapes_min=2),
        #     get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        #     round_cm = True,
        #     convert = 'pt_to_cm'
        # ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'align_table_spacing_rows',
            label="Row spacing",
            show_label=False,
            image_mso="VerticalSpacingIncrease",
            supertip="Changes the row spacing of the shapes. [ALT] switches between moving and stretch/compress.",
            on_change = bkt.Callback(shape_tables.set_spacing_rows, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(shape_tables.get_spacing_rows, shapes=True, shapes_min=2),
            get_enabled = bkt.Callback(shape_tables.enabled_spacing_rows, shapes=True),
            round_cm = True,
            convert = 'pt_to_cm'
        ),
        bkt.ribbon.RoundingSpinnerBox(
            id = 'align_table_spacing_cols',
            label="Column spacing",
            show_label=False,
            image_mso="HorizontalSpacingIncrease",
            supertip="Changes the column spacing of the shapes. [ALT] switches between moving and stretch/compress.",
            on_change = bkt.Callback(shape_tables.set_spacing_cols, shapes=True, shapes_min=2),
            get_text  = bkt.Callback(shape_tables.get_spacing_cols, shapes=True, shapes_min=2),
            get_enabled = bkt.Callback(shape_tables.enabled_spacing_cols, shapes=True),
            round_cm = True,
            convert = 'pt_to_cm'
        ),
        bkt.ribbon.Menu(
            label="Options",
            supertip="Setting options when aligning shape tables",
            item_size="large",
            children=[
                bkt.ribbon.MenuSeparator(title="Set spacing by:"),
                bkt.ribbon.ToggleButton(
                    id="toggle_table_move",
                    label="Move",
                    description="The desired shape spacing is achieved by positioning shapes",
                    image_mso="ObjectNudgeRight",
                    on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', False)),
                    get_pressed=bkt.Callback(lambda : not shape_tables.resize_cells)
                ),
                bkt.ribbon.ToggleButton(
                    id="toggle_table_resize",
                    label="Stretch/compress",
                    description="The desired shape spacing is achieved by shrinking/enlarging shapes",
                    image_mso="ShapeWidth",
                    on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', True)),
                    get_pressed=bkt.Callback(lambda : shape_tables.resize_cells)
                ),
                bkt.ribbon.MenuSeparator(title="Equalize spacing:"),
                bkt.ribbon.ToggleButton(
                    id="toggle_table_equal_spacing",
                    label="Row spacing = column spacing",
                    description="When the row spacing changes, the column spacing is also changed and vice versa",
                    image="align_table_spacing",
                    on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'equal_spacing', pressed)),
                    get_pressed=bkt.Callback(lambda : shape_tables.equal_spacing)
                ),
            ]
        ),
        # bkt.ribbon.ToggleButton(id="toggle_table_resize", label="Adj. size", show_label=True, supertip="The desired shape arrangement is achieved by shrinking/enlarging (instead of positioning) shapes", image_mso="ShapeWidth",   on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', pressed)),  get_pressed=bkt.Callback(lambda : shape_tables.resize_cells))
        # bkt.ribbon.Box(box_style="horizontal", children=[
        #     bkt.ribbon.Label(label="Mode:"),
        #     bkt.ribbon.ToggleButton(id="toggle_table_move",   label="Move",         show_label=False, supertip="The desired shape arrangement is achieved by positioning shapes", image_mso="ObjectNudgeRight",         on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', False)), get_pressed=bkt.Callback(lambda : not shape_tables.resize_cells)),
        #     bkt.ribbon.ToggleButton(id="toggle_table_resize", label="Stretch/compress", show_label=False, supertip="The desired shape arrangement is achieved by shrinking/enlarging shapes", image_mso="ShapeWidth",   on_toggle_action=bkt.Callback(lambda pressed: setattr(shape_tables, 'resize_cells', True)),  get_pressed=bkt.Callback(lambda : shape_tables.resize_cells))
        # ]),
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
        tabellen_gruppe,
    ]
), extend=True)