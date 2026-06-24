# -*- coding: utf-8 -*-
'''

@author: rdebeerst
'''

import bkt

MODEL_MODULE = __package__ + ".circular_model"
MODEL_CONTAINER = "CircularArrangement"


group_circlify = bkt.ribbon.Group(
    id="bkt_circlify_group",
    label="Circular arrangement",
    image="circlify",
    supertip="Enables the circular arrangement of shapes. The `ppt_circlify` feature must be installed.",
    children=[
        bkt.ribbon.SplitButton(
            id="circlify_splitbutton",
            size='large',
            children=[
                bkt.ribbon.Button(
                    id="circlify_button",
                    label="Arrange in a circle",
                    image="circlify", #image_mso="DiagramRadialInsertClassic",
                    # size='large',
                    supertip="Selected shapes are arranged in a circle according to the set width/height.\nThe order of the shapes depends on the selection order: the first selected shape is positioned at 12 o'clock, the other shapes follow clockwise.",
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular", shapes=True, shapes_min=3),
                    get_enabled="PythonGetEnabled"
                ),
                bkt.ribbon.Menu(
                    label="Circular arrangement options",
                    supertip="Settings for the circular alignment of shapes",
                    item_size="large",
                    children=[
                        bkt.ribbon.MenuSeparator(title="Options:"),
                        bkt.ribbon.ToggleButton(
                            label="Toggle shape rotation",
                            image_mso="ObjectRotateFree",
                            description="Rotate objects in the circular arrangement according to their position in the circle",
                            on_toggle_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_rotated"),
                            get_pressed=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_rotated_pressed")
                        ),
                        bkt.ribbon.ToggleButton(
                            label="Toggle circle (width = height)",
                            description="When the height changes, the width is also changed and vice versa",
                            image_mso="ShapeDonut",
                            on_toggle_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_fixed"),
                            get_pressed=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_fixed_pressed")
                        ),
                        bkt.ribbon.ToggleButton(
                            label="First shape in the center",
                            description="The first selected shape is placed at the circle center",
                            image_mso="DiagramTargetInsertClassic",
                            on_toggle_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_centerpoint"),
                            get_pressed=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "arrange_circular_centerpoint_pressed")
                        ),
                        bkt.ribbon.MenuSeparator(title="Functions:"),
                        bkt.ribbon.Button(
                            label="Interpolate current parameters",
                            description="An attempt is made to approximately determine the current radius, start angle and options of the selected shapes",
                            image_mso="DiagramRadialInsertClassic",
                            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "determine_ellipse_params", shapes=True, shapes_min=3),
                            get_enabled="PythonGetEnabled",
                        ),
                    ]
                ),
            ]
        ),
        bkt.ribbon.RoundingSpinnerBox(
            label="Width",
            round_cm=True,
            convert = 'pt_to_cm',
            image_mso="ShapeWidth",
            show_label=False,
            supertip="Width of the ellipse (diagonal) for the circular arrangement",
            on_change=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_circ_width", shapes=True, shapes_min=3),
            get_enabled="PythonGetEnabled",
            get_text=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "get_circ_width", shapes=True, shapes_min=3),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            label="Height",
            round_cm=True,
            convert = 'pt_to_cm',
            image_mso="ShapeHeight",
            show_label=False,
            supertip="Height of the ellipse (diagonal) for the circular arrangement",
            on_change=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_circ_height", shapes=True, shapes_min=3),
            get_enabled="PythonGetEnabled",
            get_text=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "get_circ_height", shapes=True, shapes_min=3),
        ),
        bkt.ribbon.RoundingSpinnerBox(
            label="Rotation",
            round_int = True,
            huge_step = 45,
            image_mso="DiagramCycleInsertClassic",
            show_label=False,
            supertip="The angle of the first shape specifies the rotation of the circular arrangement.",
            on_change=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_segment_start", shapes=True, shapes_min=3),
            get_enabled="PythonGetEnabled",
            get_text=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "get_segment_start", shapes=True, shapes_min=3),
        ),
    ]
)



bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    insert_before_mso="TabHome",
    label='Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        group_circlify,
        # group_segmented_circle
    ]
), extend=True)

