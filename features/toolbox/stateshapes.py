# -*- coding: utf-8 -*-
'''
Created on 21.12.2017

@author: fstallmann
'''


import bkt
import bkt.library.powerpoint as pplib


class StateShapeUi(object):
    BKT_DIALOG_TAG = 'BKT_DIALOG_STATESHAPE'

    @classmethod
    def is_convertable_to_state_shape(cls, shapes):
        try:
            if len(shapes) > 1:
                return not any(cls.is_state_shape(s) for s in shapes)
            else:
                shape = shapes[0]
                return shape.Type == pplib.MsoShapeType['msoGroup'] and not cls.is_state_shape(shape)
        except:
            return False

    @classmethod
    def is_state_shape(cls, shape):
        return pplib.TagHelper.has_tag(shape, bkt.contextdialogs.BKT_CONTEXTDIALOG_TAGKEY, cls.BKT_DIALOG_TAG)
        # return shape.Type == pplib.MsoShapeType['msoGroup']
    
    @classmethod
    def are_state_shapes(cls, shapes):
        return all(cls.is_state_shape(s) for s in shapes)



def stateshape_fill1_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
                    label = 'Change color 1 (background)',
                    image_mso = 'ShapeFillColorPicker',
                    screentip="Change the background color of a toggle shape",
                    supertip="Adjusts the background color of all shapes in the toggle shape. The background color is the color of the first found shape.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_rgb1", shapes=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_theme1", shapes=True),
                    # get_selected_color    = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "get_selected_color1", shapes=True),
                    get_enabled           = bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
                    children=[
                        bkt.ribbon.Button(
                            label="No background",
                            supertip="Set toggle shape background color to transparent",
                            on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_none1", shapes=True),
                        ),
                    ],
                    **kwargs
                )

def stateshape_fill2_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
                    label = 'Change color 2 (foreground)',
                    image_mso = 'ShapeFillColorPicker',
                    screentip="Change the foreground color of a toggle shape",
                    supertip="Adjusts the foreground color of all shapes in the toggle shape. The foreground color is any color other than the background color.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_rgb2", shapes=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_fill_theme2", shapes=True),
                    # get_selected_color    = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "get_selected_color2", shapes=True),
                    get_enabled           = bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
                    **kwargs
                )

def stateshape_line_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
                    label = 'Change line',
                    image_mso = 'ShapeOutlineColorPicker',
                    screentip="Change the line of a toggle shape",
                    supertip="Adjusts the line color of all shapes in the toggle shape that match the first found line color.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_line_rgb", shapes=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_line_theme", shapes=True),
                    # get_selected_color    = bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "get_selected_line", shapes=True),
                    get_enabled           = bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
                    children=[
                        bkt.ribbon.Button(
                            label="No line",
                            supertip="Set toggle shape line to transparent",
                            on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_color_line_none", shapes=True),
                        ),
                    ],
                    **kwargs
                )



stateshape_gruppe = bkt.ribbon.Group(
    id="bkt_stateshape_group",
    label='Toggle shapes',
    image_mso='GroupSmartArtQuickStyles',
    children = [
        bkt.ribbon.SplitButton(
            id="stateshape_convert_splitbutton",
            # size="large",
            children=[
                bkt.ribbon.Button(
                    id="stateshape_convert",
                    label="Convert",
                    image_mso='GroupSmartArtQuickStyles',
                    screentip="Convert grouped shapes into a toggle shape",
                    supertip="For grouped shapes (toggle shapes), you can switch between the shapes within the group, i.e. only one shape of the group is ever visible. This is useful e.g. for traffic lights, scales, etc.",
                    on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "convert_to_state_shape", shapes=True),
                    get_enabled=bkt.Callback(StateShapeUi.is_convertable_to_state_shape, shapes=True),
                ),
                bkt.ribbon.Menu(
                    label="Toggle shapes menu",
                    supertip="Convert to toggle shapes or make all shapes visible again",
                    children=[
                        bkt.ribbon.Button(
                            id="stateshape_convert2",
                            label="Convert to toggle shape",
                            image_mso='GroupSmartArtQuickStyles',
                            screentip="Convert grouped shapes into a toggle shape",
                            supertip="For grouped shapes (toggle shapes), you can switch between the shapes within the group, i.e. only one shape of the group is ever visible. This is useful e.g. for traffic lights, scales, etc.",
                            on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "convert_to_state_shape", shapes=True),
                            get_enabled=bkt.Callback(StateShapeUi.is_convertable_to_state_shape, shapes=True),
                        ),
                        bkt.ribbon.MenuSeparator(),
                        # bkt.ribbon.ToggleButton(
                        bkt.ribbon.Button(
                            id="stateshape_show_all",
                            label="Show all shapes again",
                            screentip="Make all shapes visible",
                            supertip="With this button, the shapes within the toggle-shape group can be shown.",
                            # image_mso='GroupSmartArtQuickStyles',
                            # get_pressed=bkt.Callback(StateShape.get_show_all),
                            # on_toggle_action=bkt.Callback(StateShape.toggle_show_all),
                            on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "show_all", shape=True),
                            get_enabled=bkt.Callback(StateShapeUi.is_state_shape, shape=True),
                        ),
                    ]
                )
            ]
        ),
        # bkt.ribbon.Separator(),
        # bkt.ribbon.LabelControl(label="Toggle:"),
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.Button(
                id="stateshape_reset",
                image_mso="Undo",
                label="Reset",
                show_label=False,
                screentip="Reset to first shape",
                supertip="Resets all toggle shapes to the first state, i.e. the first shape of the group.",
                on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "reset_state", shapes=True),
                get_enabled=bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
            ),
            bkt.ribbon.Button(
                id="stateshape_prev",
                image_mso="PreviousResource",
                label='Previous',
                show_label=False,
                screentip="Previous shape",
                supertip="Switches to the previous state (i.e. shape in the group) of the toggle shape.",
                on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "previous_state", shapes=True),
                get_enabled=bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
            ),
            # bkt.ribbon.EditBox(
            #     id="stateshape_set",
            #     label="Position",
            #     show_label=False,
            #     size_string="#",
            #     on_change=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "set_state"),
            #     get_enabled=bkt.Callback(StateShapeUi.are_state_shapes),
            #     get_text=bkt.Callback(lambda: None),
            # ),
            bkt.ribbon.Button(
                id="stateshape_next",
                image_mso="NextResource",
                label="Next",
                # show_label=False,
                screentip="Next shape",
                supertip="Switches to the next state (i.e. shape in the group) of the toggle shape.",
                on_action=bkt.CallbackLazy("toolbox.models.stateshapes", "StateShape", "next_state", shapes=True),
                get_enabled=bkt.Callback(StateShapeUi.are_state_shapes, shapes=True),
            )
        ]),
        bkt.ribbon.Menu(
            id="stateshape_color_menu",
            label="Change color",
            supertip="Adjust the colors of toggle shapes",
            image_mso="RecolorColorPicker",
            children=[
                stateshape_fill1_gallery(id="stateshape_color_fill1"),
                stateshape_fill2_gallery(id="stateshape_color_fill2"),
                stateshape_line_gallery(id="stateshape_color_line"),
            ]
        ),
        # bkt.ribbon.Button(
        #     id="stateshape_help",
        #     image_mso="Help",
        #     label=u"Instructions",
        #     on_action=bkt.Callback(StateShape.show_help),
        #     # get_enabled=bkt.Callback(StateShape.are_state_shapes),
        # ),
        # likert_button,
    ]
)
