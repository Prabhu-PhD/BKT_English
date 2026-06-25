# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

import bkt
import bkt.library.powerpoint as powerpoint


class HarveyBallsUi(object):
    BKT_HARVEY_DIALOG_TAG = "BKT_SHAPE_HARVEY"
    BKT_HARVEY_DENOM_TAG = "BKT_HARVEY_DENOM_TAG"
    BKT_HARVEY_VERSION = "BKT_HARVEY_V2"

    BKT_HARVEY_LEGACY_VERSION = ("BKT_HARVEY_V1")

    # =====================================
    # = Feature-Logik und Hilfsfunktionen =
    # =====================================
    
    def is_harvey_group(self, shape):
        # "new" method via tags
        if powerpoint.TagHelper.has_tag(shape, self.BKT_HARVEY_DIALOG_TAG):
            return True
        # "old" method via shape types
        pie, _ = self.get_pie_circ(shape)
        return pie != None

    def get_pie_circ(self, shape):
        if not shape.Type == powerpoint.MsoShapeType['msoGroup']:
            return None, None
        if not shape.GroupItems.Count == 2:
            return None, None

        pie_types = (powerpoint.MsoAutoShapeType['msoShapePie'],powerpoint.MsoAutoShapeType['msoShapeArc'],powerpoint.MsoAutoShapeType['msoShapeBlockArc'])

        if shape.GroupItems(1).AutoShapeType in pie_types:
            return shape.GroupItems(1), shape.GroupItems(2)
        elif shape.GroupItems(2).AutoShapeType in pie_types:
            return shape.GroupItems(2), shape.GroupItems(1)
        else:
            return None, None

    def change_harvey_enabled(self, shapes):
        return self.is_harvey_group(shapes[0])

harveyui = HarveyBallsUi()


def harvey_color_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
        label = 'Change color',
        #image_mso = 'RecolorColorPicker',
        image='harvey ball color',
        screentip="Change the color of a Harvey ball",
        supertip="Adjust the color of a Harvey ball according to the selection.\n\nA Harvey ball shape is a group of a circle and a pie shape.",
        on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "color_gallery_action", shapes=True),
        on_theme_color_change = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "color_gallery_theme_color_change", shapes=True),
        get_selected_color    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_selected_color", shapes=True),
        get_enabled           = bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
        item_width=16, item_height=16,
        **kwargs
    )

def harvey_background_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
        label = 'Change background',
        #image_mso = 'RecolorColorPicker',
        image='harvey ball background',
        screentip="Change the background of a Harvey ball",
        supertip="Adjust the background color of a Harvey ball according to the selection.\n\nA Harvey ball shape is a group of a circle and a pie shape.",
        on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "background_gallery_action", shapes=True),
        on_theme_color_change = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "background_gallery_theme_color_change", shapes=True),
        get_selected_color    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_selected_background", shapes=True),
        get_enabled           = bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
        children=[
            bkt.ribbon.Button(
                label='Background off',
                supertip="Set Harvey ball background to transparent",
                #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
                image='harvey ball background',
                on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_background_off", shapes=True),
            ),
        ],
        item_width=16, item_height=16,
        **kwargs
    )

def harvey_line_gallery(**kwargs):
    return bkt.ribbon.ColorGallery(
        label = 'Change line color',
        #image_mso = 'RecolorColorPicker',
        image='harvey ball line',
        screentip="Change the line color of a Harvey ball",
        supertip="Adjust the line color of a Harvey ball according to the selection.\n\nA Harvey ball shape is a group of a circle and a pie shape.",
        on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "line_gallery_action", shapes=True),
        on_theme_color_change = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "line_gallery_theme_color_change", shapes=True),
        get_selected_color    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_selected_line", shapes=True),
        get_enabled           = bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
        children=[
            bkt.ribbon.Button(
                label='Line off',
                supertip="Hide Harvey ball line",
                #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
                image='harvey ball line off',
                on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_line_off", shapes=True),
            ),
            bkt.ribbon.Button(
                label='Toggle line outside only',
                supertip="The Harvey ball line is shown either around the whole pie or only around the outer circle",
                #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
                image='harvey ball line outside',
                on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_line_outside_only", shapes=True),
            ),
        ],
        item_width=16, item_height=16,
        **kwargs
    )

def harvey_size_gallery(**kwargs):
    return bkt.ribbon.Gallery(
        label = 'Change fill level',
        image = 'harvey ball size',
        #get_image=bkt.Callback(lambda: harvey_balls.get_harvey_image(0.6, 64)),
        screentip="Change the fill level of a Harvey ball",
        supertip="Adjust the fill level of a Harvey ball according to the selection.\n\nA Harvey ball shape is a group of a circle and a pie shape.",
        columns="9", #9=harvey_columns
        on_action_indexed = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "change_harvey", shapes=True),
        get_item_count    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_item_count"),
        get_item_label    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_item_label"),
        get_item_screentip = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_item_screentip"),
        get_item_supertip = bkt.Callback(lambda index: "Adjust the fill level of a Harvey ball according to the selection."),
        get_enabled       = bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
        get_item_image    = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_item_image"),
        item_width=16, item_height=16,
        **kwargs
    )


harvey_create_button = bkt.ribbon.Button(
    id='create_harvey_ball',
    label='Harvey Ball',
    screentip='Create Harvey ball',
    image='harvey ball',
    on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "create_harvey_ball", context=True, slide=True),
    supertip="Insert a Harvey ball that can be configured for color/fill level.\n\nColor and fill level can be configured via the context menu and context tab; in the tab, percentage values are also possible.\n\nA Harvey ball shape is a group of a circle and a pie shape."
)


harvey_ball_group = bkt.ribbon.Group(
    id="bkt_harvey_group",
    label = "Harvey Balls",
    children = [
        bkt.ribbon.Button(
            id='harvey_ball_create',
            size='large',
            label='New Harvey ball',
            screentip='Create Harvey ball',
            image='harvey ball',
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "create_harvey_ball", context=True, slide=True),
            supertip="Insert a Harvey ball that can be configured for color/fill level.\n\nColor and fill level can be configured via the context menu and context tab; in the tab, percentage values are also possible.\n\nA Harvey ball shape is a group of a circle and a pie shape."
        ),
        bkt.ribbon.Button(
            id='harvey_ball_duplicate',
            size='large',
            label='Duplicate Harvey ball',
            screentip='Duplicate Harvey ball',
            image='harvey ball duplicate',
            on_action=bkt.Callback(lambda selection: selection.ShapeRange.Duplicate().Select()),
            supertip="Duplicates the currently selected Harvey ball."
        ),
        bkt.ribbon.Separator(),

        #bkt.ribbon.SplitButton(show_label=False, children=[
            # bkt.ribbon.Button(
            #     id='create_harvey_ball',
            #     label='Create Harvey ball',
            #     screentip='Create Harvey ball',
            #     image='harvey ball',
            #     on_action=bkt.Callback(harvey_balls.create_harvey_ball)
            # ),
            # bkt.ribbon.Menu(label='menu',
            #     children = [
        harvey_size_gallery(id='harvey_ball_size_gallery', size="large"),
        harvey_color_gallery(id='harvey_ball_color_gallery', size="large"),
        #         ]
        #     )
        # ]),

        harvey_background_gallery(id='harvey_ball_background', size="large"),
        harvey_line_gallery(id='harvey_ball_line', size="large"),

        bkt.ribbon.Separator(),

        bkt.ribbon.Button(
            id='harvey_ball_style_classic',
            size='large',
            label='Style classic',
            supertip="Display the Harvey ball in the classic style without an additional border.",
            image='harvey ball classic',
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_change_style_classic", shapes=True),
        ),

        bkt.ribbon.Button(
            id='harvey_ball_style_modern',
            size='large',
            label='Style modern',
            supertip="Display the Harvey ball in the modern style with a white border.",
            image='harvey ball modern',
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_change_style_modern", shapes=True),
        ),

        bkt.ribbon.Button(
            id='harvey_ball_style_chart',
            size='large',
            label='Style chart',
            supertip="Harvey ball in chart style with highlighted fill level.",
            image='harvey ball diagram',
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_change_style_chart", shapes=True),
        ),

        bkt.ribbon.ToggleButton(
            id='harvey_ball_flip',
            size='large',
            label='Counterclockwise',
            supertip="Mirror the Harvey ball to show the fill level clockwise or counterclockwise.",
            image='harvey ball flip',
            get_pressed=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_fliph_pressed", shapes=True),
            on_toggle_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_fliph", shapes=True),
        ),

        bkt.ribbon.Separator(),
        #bkt.ribbon.LabelControl(label="Fill level:"),
        
        bkt.ribbon.SpinnerBox(label='Fill level in %', size_string='33,33',
            id = 'harvey_spinner',
            screentip="Change the fill level of a Harvey ball",
            supertip="Adjust the fill level of a Harvey ball according to the selection.\n\nA Harvey ball shape is a group of a circle and a pie shape.",
            on_change = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_percent_setter", shapes=True),
            get_text  = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_percent", shapes=True),
            increment = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_percent_inc", shapes=True),
            decrement = bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "harvey_percent_dec", shapes=True)
        ),
        bkt.ribbon.LabelControl(label="   with SHIFT: step size +/-25"),
        bkt.ribbon.LabelControl(label="   with ALT: delta per Harvey Ball"),

        bkt.ribbon.Separator(),

        bkt.ribbon.Button(
            id='harvey_ball_0',
            size='large',
            label='0%',
            screentip="Harvey ball to 0%",
            supertip="Sets all selected Harvey balls to 0%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="0",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_25',
            size='large',
            label='25%',
            screentip="Harvey ball to 25%",
            supertip="Sets all selected Harvey balls to 25%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="25",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_33',
            size='large',
            label='33%',
            screentip="Harvey ball to 33%",
            supertip="Sets all selected Harvey balls to 33%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="33.3",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_50',
            size='large',
            label='50%',
            screentip="Harvey ball to 50%",
            supertip="Sets all selected Harvey balls to 50%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="50",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_66',
            size='large',
            label='66%',
            screentip="Harvey ball to 66%",
            supertip="Sets all selected Harvey balls to 66%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="66.6",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_75',
            size='large',
            label='75%',
            screentip="Harvey ball to 75%",
            supertip="Sets all selected Harvey balls to 75%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="75",
        ),
        bkt.ribbon.Button(
            id='harvey_ball_100',
            size='large',
            label='100%',
            screentip="Harvey ball to 100%",
            supertip="Sets all selected Harvey balls to 100%",
            get_image=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "get_harvey_image_by_control", current_control=True),
            on_action=bkt.CallbackLazy("toolbox.models.harvey", "harvey_balls", "set_harvey_by_control", shapes=True, current_control=True),
            tag="100",
        ),

        # bkt.ribbon.Separator(),

        # bkt.ribbon.Button(
        #     id='harvey_legacy_upgrade',
        #     size='large',
        #     label='Update version',
        #     screentip="Update Harvey ball to the latest version",
        #     supertip="The Harvey ball is updated to the latest version, which makes it look a bit better and allows the line color to be adjusted. However, it is then no longer compatible with older versions.",
        #     image_mso="UpdateIcon",
        #     get_visible=bkt.Callback(harvey_balls.is_legacy_any),
        #     on_action=bkt.Callback(harvey_balls.upgrade_all),
        # ),
    ]
)

harvey_ball_tab = bkt.ribbon.Tab(
    id = "bkt_context_tab_harvey",
    label = "[BKT] Harvey Balls",
    get_visible=bkt.Callback(harveyui.change_harvey_enabled, shapes=True),
    children = [
        # Harvey Balls
        harvey_ball_group
    ]
)