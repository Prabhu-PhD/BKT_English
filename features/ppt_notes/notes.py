# -*- coding: utf-8 -*-
'''
Created on 29.03.2017

@author: tweuffel
'''

import bkt

MODEL_MODULE = __package__ + ".notes_model"
MODEL_CONTAINER = "EditModeShapes"


notes_gruppe = bkt.ribbon.Group(
    id="bkt_notes_group",
    label='Slide notes',
    supertip="Enables inserting editing notes on slides. The `ppt_notes` feature must be installed.",
    image='noteAdd',
    children = [
        bkt.ribbon.Button(
            id = 'notes_add',
            label='Create', screentip='Add note',
            supertip="Inserts an editing note in the top right of the slide incl. author and date.",
            image='noteAdd',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "addNote", slide=True, context=True)
        ),
        bkt.ribbon.Button(
            id = 'notes_toggle',
            label='On/off', screentip='Show/hide notes on slide',
            supertip="Temporarily hide and show all notes on the current slide.",
            image='noteToggle',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "toogleNotesOnSlides", slides=True)
        ),
        bkt.ribbon.Button(
            id = 'notes_remove',
            label='Delete', screentip='Delete notes on slide',
            supertip="Remove all notes on the current slide.",
            image='noteRemove',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "removeNotesOnSlides", slides=True)
        ),
        bkt.ribbon.Button(
            id = 'notes_toggle_all',
            label='All on/off', screentip='Show/hide all notes',
            supertip="Temporarily hide and show all notes on all slides.",
            image='noteToggleAll',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "toggleNotesOnAllSlides", slide=True)
        ),
        bkt.ribbon.Button(
            id = 'notes_remove_all',
            label='Delete all', screentip='Delete all notes',
            supertip="Remove all notes on all slides.",
            image='noteRemoveAll',
            on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "removeNotesOnAllSlides", slide=True)
        ),
        bkt.ribbon.ColorGallery(
            id = 'notes_color',
            label='Color',
            screentip='Change note color',
            supertip="Change the background color for new editing notes.",
            on_rgb_color_change = bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_color_rgb"),
            on_theme_color_change = bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_color_theme"),
            get_selected_color = bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "get_color"),
            children=[
                bkt.ribbon.Button(
                    id="notes_color_default",
                    label="Default color",
                    supertip="Reset the background color for editing notes to default.",
                    on_action=bkt.CallbackLazy(MODEL_MODULE, MODEL_CONTAINER, "set_color_default"),
                    image_mso="ColorTeal",
                )
            ]
            # get_enabled = bkt.apps.ppt_shapes_or_text_selected,
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
        notes_gruppe,
    ]
), extend=True)


