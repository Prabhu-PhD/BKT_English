# -*- coding: utf-8 -*-
'''
Created on 18.05.2016

@author: rdebeerst
'''


import bkt


# Menu
#   Neue Agenda erstellen --> macht immer neue Id
#   Agenda aktualisieren --> nur verfügbar, wenn Id gefunden
#   Remove Agenda-Information from slide
#   Remove Agenda-Information from presentation
#   Show/Hide Agenda-Overview
#
# Einstellungen
# {
#   id
#   hide-agenda-overview
#   hide other sub-agenda
#   tracker-color
#   ...
# }
# 
# Vorhandene Agenda-Slides finden --> [ [Nr, Text, Slide], ...]
# Neue Agenda-Slides bestimmen --> [ [Nr, Text, Indent, Slide], ...]
# 
# Mapping neu auf vorhanden
#   for a in new_agenda:
#       take first that matches text
#       mapped_b = [b for b in current_agenda if b[1] == a[1]][1]
#       remove mapped_b from current_agenda
#

TOOLBOX_AGENDA = "TOOLBOX-AGENDA"
TOOLBOX_AGENDA_SLIDENO  = "TOOLBOX-AGENDA-SLIDENO"
TOOLBOX_AGENDA_SELECTOR = "TOOLBOX-AGENDA-SELECTOR"
TOOLBOX_AGENDA_TEXTBOX  = "TOOLBOX-AGENDA-TEXTBOX"
# TOOLBOX_AGENDA_SETTINGS = "TOOLBOX-AGENDA-SETTINGS"

class ToolboxAgendaUi(object):
    @classmethod
    def can_create_agenda_from_slide(cls, slide):
        ''' check if agenda textbox is on slide in order to create agenda from textbox '''
        try:
            if slide.Tags.Item(TOOLBOX_AGENDA) != "":
                textbox = cls.get_agenda_textbox_on_slide(slide)
                return textbox != None
            else:
                return False
        except: #AttributeError
            return False

    @classmethod
    def is_agenda_slide(cls, slide):
        ''' check if current slide is agenda-slide '''
        try:
            return slide.Tags.Item(TOOLBOX_AGENDA_SLIDENO) != ""
        except: #AttributeError
            return False

    @classmethod
    def presentation_has_agenda(cls, presentation):
        '''
        check if any slide has agenda-tag on slide
        '''
        for slide in presentation.slides:
            if cls.is_agenda_slide(slide):
                return True
        return False
    
    @classmethod
    def get_agenda_textbox_on_slide(cls, sld):
        '''
        return agenda-textbox on given slide
        agenda-textbox is recognised by the tag TOOLBOX_AGENDA_TEXTBOX
        '''
        return cls.get_shape_with_tag_item(sld, TOOLBOX_AGENDA_TEXTBOX)

    @staticmethod
    def get_shape_with_tag_item(sld, tagKey):
        ''' Shape auf Slide finden, das einen bestimmten TagKey enthaelt '''
        for shp in sld.shapes:
            if shp.Tags.Item(tagKey) != "":
                return shp 
        return None


agendamenu = bkt.ribbon.Menu(
    label="Agenda",
    children=[
        bkt.ribbon.Button(
            id='add-agenda-textbox',
            label="Insert agenda text box",
            supertip="Insert a standard agenda text box to generate an updatable agenda from it.",
            imageMso="TextBoxInsert",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda" ,"create_agenda_textbox_on_slide", slide=True, context=True)
        ),
        bkt.ribbon.Button(
            id='agenda-new-create',
            label="Recreate agenda",
            supertip="Create a new agenda based on the current slide. The current slide becomes the agenda main slide.",
            imageMso="TableOfContentsAddTextGallery",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "create_agenda_from_slide", slide=True, context=True),
            get_enabled=bkt.Callback(ToolboxAgendaUi.can_create_agenda_from_slide, slide=True)
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id='agenda-new-update',
            label="Update agenda",
            supertip="Update agenda and replace with the agenda on the agenda main slide; slides are recreated in the process.",
            imageMso="SaveSelectionToTableOfContentsGallery",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "update_agenda_slides_by_slide", slide=True),
            get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True)
        ),
        bkt.ribbon.DynamicMenu(
            id='agenda-options-menu',
            label="Agenda settings",
            get_content=bkt.CallbackLazy("toolbox.models.agenda", "agenda_options_menu"),
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id='agenda-remove',
            label="Remove agenda slide",
            supertip="Removes agenda slides of the selected agenda; all meta information is deleted.",
            imageMso="TableOfContentsRemove",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "remove_agenda", slide=True, presentation=True),
            get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True)
        ),
        bkt.ribbon.Button(
            id='agenda-remove-all',
            label="Remove all agendas from the presentation",
            supertip="Removes all agenda slides in the entire presentation; all meta information is deleted.",
            imageMso="TableOfContentsRemove",
            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "remove_agendas_from_presentation", presentation=True),
            get_enabled=bkt.Callback(ToolboxAgendaUi.presentation_has_agenda, presentation=True)
        ),
        # bkt.ribbon.Menu(
        #     label="Old",
        #     children=[
        #         bkt.ribbon.Button(
        #             id='agenda-update',
        #             label="Update agenda",
        #             screentip="Update existing agenda (based on the first agenda slide) or create agenda slides from the current slide.",
        #             imageMso="GroupAddInsMenuCommands",
        #             on_action=bkt.Callback(ToolboxAgenda.create_or_update_agenda)
        #         ),
        #     ]
        # )
        
    ]
)

agenda_tab = bkt.ribbon.Tab(
    id = "bkt_context_tab_agenda",
    label = "[BKT] Agenda",
    # get_visible=bkt.Callback(ToolboxAgenda.is_agenda_slide, slide=True),
    get_visible=bkt.Callback(ToolboxAgendaUi.can_create_agenda_from_slide, slide=True),
    children = [
        bkt.ribbon.Group(
            id="bkt_agenda_manual",
            label = "Instructions",
            children = [
                bkt.ribbon.Label(label='Step 1: Fill the text box with the agenda and "Recreate agenda"'),
                bkt.ribbon.Label(label='Step 2: After every further change, "Update agenda"'),
                bkt.ribbon.Label(label='Note: the agenda main slide should not be deleted!'),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_agenda_group",
            label = "Agenda",
            children = [
                bkt.ribbon.Button(
                    id='agenda_new_create',
                    label="Recreate agenda",
                    size="large",
                    supertip="Create a new agenda based on the current slide. The current slide becomes the agenda main slide.",
                    imageMso="TableOfContentsAddTextGallery",
                    on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "create_agenda_from_slide", slide=True, context=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.can_create_agenda_from_slide, slide=True)
                ),
                bkt.ribbon.Button(
                    id='agenda_new_update',
                    label="Update agenda",
                    size="large",
                    supertip="Update agenda and replace with the agenda on the agenda main slide; slides are recreated in the process.",
                    imageMso="SaveSelectionToTableOfContentsGallery",
                    on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "update_agenda_slides_by_slide", slide=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True)
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.DynamicMenu(
                    label="Options",
                    screentip="Agenda options",
                    supertip="Change various agenda options",
                    imageMso="TableProperties",
                    size="large",
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True),
                    get_content=bkt.CallbackLazy("toolbox.models.agenda", "agenda_options_menu"),
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id='agenda_remove',
                    label="Remove agenda slides",
                    size="large",
                    screentip="Remove all associated agenda slides",
                    supertip="Removes all agenda slides belonging to the current agenda except the main slide. All meta information is deleted.",
                    imageMso="TableOfContentsRemove",
                    on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "remove_agenda", slide=True, presentation=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True)
                ),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_agenda_selector_group",
            label = "Agenda selector",
            children = [
                bkt.ribbon.ColorGallery(
                    label = 'Change background',
                    size="large",
                    image_mso = 'ShapeFillColorPicker',
                    screentip="Background color for selector",
                    supertip="Adjust the background color for the selector that highlights the active agenda item.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_fillcolor_rgb", slide=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_fillcolor_theme", slide=True),
                    get_selected_color    = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_selector_fillcolor", slide=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True),
                    children=[
                        bkt.ribbon.Button(
                            label="No fill",
                            supertip="Change selector background to transparent",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "hide_selector_fill", slide=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_fillcolor", slide=True),
                        ),
                        bkt.ribbon.Button(
                            label="Reset",
                            screentip="Reset selector background",
                            supertip="Reset selector background to default",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "reset_selector_fillcolor", slide=True)
                        ),
                    ]
                ),
                bkt.ribbon.ColorGallery(
                    label = 'Change border',
                    size="large",
                    image_mso = 'ShapeOutlineColorPicker',
                    screentip="Line color for selector",
                    supertip="Adjust the line color for the selector that highlights the active agenda item.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_linecolor_rgb", slide=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_linecolor_theme", slide=True),
                    get_selected_color    = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_selector_linecolor", slide=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True),
                    children=[
                        bkt.ribbon.Button(
                            label="No border",
                            supertip="Change selector border to transparent",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "hide_selector_line", slide=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_linecolor", slide=True),
                        ),
                        bkt.ribbon.Button(
                            label="Reset",
                            screentip="Reset selector border",
                            supertip="Reset selector border to default",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "reset_selector_linecolor", slide=True)
                        ),
                    ]
                ),
                bkt.ribbon.ColorGallery(
                    label = 'Change text',
                    size="large",
                    image_mso = 'TextFillColorPicker',
                    screentip="Text color for selector",
                    supertip="Adjust the text color for the selector that highlights the active agenda item.",
                    on_rgb_color_change   = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_textcolor_rgb", slide=True),
                    on_theme_color_change = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_textcolor_theme", slide=True),
                    get_selected_color    = bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_selector_textcolor", slide=True),
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide, slide=True),
                    children=[
                        bkt.ribbon.Button(
                            label="Bold",
                            screentip="Selector text bold",
                            supertip="Toggle selector text bold",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "toggle_selector_text_style", slide=True, current_control=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_textcolor", slide=True, current_control=True),
                            tag="bold",
                        ),
                        bkt.ribbon.Button(
                            label="Italic",
                            screentip="Selector text italic",
                            supertip="Toggle selector text italic",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "toggle_selector_text_style", slide=True, current_control=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_textcolor", slide=True, current_control=True),
                            tag="italic",
                        ),
                        bkt.ribbon.Button(
                            label="Underlined",
                            screentip="Selector text underlined",
                            supertip="Toggle selector text underlined",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "toggle_selector_text_style", slide=True, current_control=True),
                            get_image=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_check_textcolor", slide=True, current_control=True),
                            tag="underline",
                        ),
                        bkt.ribbon.Button(
                            label="Reset",
                            screentip="Reset selector text",
                            supertip="Reset selector text to default",
                            on_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "reset_selector_textcolor", slide=True)
                        ),
                    ]
                ),
                bkt.ribbon.Menu(
                    label = "Change height",
                    size="large",
                    image_mso = 'GroupInkEdit',
                    screentip="Height for selector",
                    supertip="Adjusts the height of the selector relative to the font size.",
                    get_enabled=bkt.Callback(ToolboxAgendaUi.is_agenda_slide),
                    children=[
                        bkt.ribbon.ToggleButton(
                            label="20% (default)",
                            screentip="Selector height 20%",
                            supertip="Selector overhang equals 20% of the font size",
                            get_pressed=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_pressed_selector_margin", slide=True, current_control=True),
                            on_toggle_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_margin", slide=True, current_control=True),
                            tag="0.2",
                        ),
                        bkt.ribbon.ToggleButton(
                            label="40%",
                            screentip="Selector height 40%",
                            supertip="Selector overhang equals 40% of the font size",
                            get_pressed=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_pressed_selector_margin", slide=True, current_control=True),
                            on_toggle_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_margin", slide=True, current_control=True),
                            tag="0.4",
                        ),
                        bkt.ribbon.ToggleButton(
                            label="60%",
                            screentip="Selector height 60%",
                            supertip="Selector overhang equals 60% of the font size",
                            get_pressed=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_pressed_selector_margin", slide=True, current_control=True),
                            on_toggle_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_margin", slide=True, current_control=True),
                            tag="0.6",
                        ),
                        bkt.ribbon.ToggleButton(
                            label="80% (very large)",
                            screentip="Selector height 80%",
                            supertip="Selector overhang equals 80% of the font size",
                            get_pressed=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "get_pressed_selector_margin", slide=True, current_control=True),
                            on_toggle_action=bkt.CallbackLazy("toolbox.models.agenda", "ToolboxAgenda", "set_selector_margin", slide=True, current_control=True),
                            tag="0.8",
                        ),
                    ]
                ),
            ]
        )
    ]
)
