# -*- coding: utf-8 -*-
'''
Created on 06.09.2018

@author: fstallmann
'''

import bkt
import bkt.library.powerpoint as pplib


BKT_LINK_UUID = "BKT_LINK_UUID"


class LinkedShapesUi(object):
    ### Enabled/Visible callbacks ###

    @staticmethod
    def is_linked_shape(shape):
        return pplib.TagHelper.has_tag(shape, BKT_LINK_UUID)

    @classmethod
    def not_is_linked_shape(cls, shape):
        return not cls.is_linked_shape(shape)

    @classmethod
    def are_linked_shapes(cls, shapes):
        return all(cls.is_linked_shape(shape) for shape in shapes)

    @classmethod
    def enabled_add_linked_shapes(cls):
        return cls.current_link_guid != None



linkshapes_tab = bkt.ribbon.Tab(
    id = "bkt_context_tab_linkshapes",
    label = "[BKT] Linked shapes",
    get_visible=bkt.Callback(LinkedShapesUi.are_linked_shapes, shapes=True),
    children = [
        bkt.ribbon.Group(
            id="bkt_linkshapes_find_group",
            label = "Find linked shapes",
            get_visible=bkt.apps.ppt_shapes_exactly1_selected,
            children = [
                bkt.ribbon.Box(box_style="horizontal", children=[
                    bkt.ribbon.Button(
                        id = 'linked_shapes_first',
                        label="Find first linked shape",
                        show_label=False,
                        image_mso="MailMergeGoToFirstRecord",
                        screentip="Go to the first linked shape",
                        supertip="Searches for the first linked shape.",
                        on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "goto_first_shape", shape=True, context=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'linked_shapes_previous',
                        label="Find previous linked shape",
                        show_label=False,
                        image_mso="MailMergeGoToPreviousRecord",
                        screentip="Go to the previous linked shape",
                        supertip="Searches for the previous linked shape. If there is no further shape on the previous slides, the last linked shape of the presentation is searched for.",
                        on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "goto_previous_shape", shape=True, context=True),
                    ),
                    bkt.ribbon.Label(
                        label="Go to",
                    ),
                    bkt.ribbon.Button(
                        id = 'linked_shapes_next',
                        label="Find next linked shape",
                        show_label=False,
                        image_mso="MailMergeGoToNextRecord",
                        screentip="Go to the next linked shape",
                        supertip="Searches for the next linked shape. If there is no further shape on the following slides, the first linked shape of the presentation is searched for.",
                        on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "goto_next_shape", shape=True, context=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'linked_shapes_last',
                        label="Find last linked shape",
                        show_label=False,
                        image_mso="MailMergeGotToLastRecord",
                        screentip="Go to the last linked shape",
                        supertip="Searches for the last linked shape.",
                        on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "goto_last_shape", shape=True, context=True),
                    ),
                ]),
                bkt.ribbon.Button(
                    id = 'linked_shapes_count',
                    label="Count shapes",
                    image_mso="FormattingUnique",
                    screentip="Count all linked shapes",
                    supertip="Counts the number of linked shapes on all slides.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "count_link_shapes", shape=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_select',
                    label="Show slides",
                    image_mso="SlideTransitionApplyToAll",
                    screentip="Show all slide numbers with linked shapes",
                    supertip="Shows all slide numbers that contain associated linked shapes.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "select_link_shapes_slides", shape=True, context=True),
                ),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_linkshapes_align_group",
            label = "Match linked shapes",
            children = [
                bkt.ribbon.DynamicMenu(
                    id = 'linked_shapes_master',
                    label="Choose reference",
                    image_mso="CircularReferences",
                    size="large",
                    screentip="Select reference shape",
                    supertip="Select whether the selected, first or last shape should be used as the reference for all matching functions. The default is the currently selected shape.",
                    get_content=bkt.CallbackLazy("toolbox.models.linkshapes", "reference_menu")
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_all',
                    label="Match everything",
                    image_mso="GroupUpdate",
                    size="large",
                    screentip="Match all properties of linked shapes",
                    supertip="Set all properties of all linked shapes like the selected shape.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "equalize_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id = 'linked_shapes_align',
                    label="Match position",
                    image_mso="ControlAlignToGrid",
                    screentip="Match position of linked shapes",
                    supertip="Set the position and rotation of all linked shapes to the position of the selected shape.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "align_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_size',
                    label="Match size",
                    image_mso="SizeToControlHeightAndWidth",
                    screentip="Match size of linked shapes",
                    supertip="Set the size of all linked shapes to the size of the selected shape.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "size_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_format',
                    label="Match formatting",
                    image_mso="FormatPainter",
                    screentip="Match formatting of linked shapes",
                    supertip="Set the formatting of all linked shapes to the size of the selected shape.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "format_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_text',
                    label="Match text",
                    image_mso="TextBoxInsert",
                    screentip="Match text of linked shapes",
                    supertip="Set the text of all linked shapes to the size of the selected shape.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "text_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.DynamicMenu(
                    id="linked_shapes_actions",
                    label="Execute action",
                    supertip="Perform various actions on all linked shapes",
                    image_mso="ObjectBringToFront",
                    get_content=bkt.CallbackLazy("toolbox.models.linkshapes", "action_menu")
                ),
                bkt.ribbon.DynamicMenu(
                    id="linked_shapes_properties",
                    label="Match property",
                    supertip="Transfer a single property to all linked shapes",
                    image_mso="ObjectNudgeRight",
                    get_content=bkt.CallbackLazy("toolbox.models.linkshapes", "properties_menu")
                ),
                bkt.ribbon.Separator(),
                bkt.ribbon.Button(
                    id = 'linked_shapes_delete',
                    label="Delete other shapes",
                    image_mso="HyperlinkRemove",
                    screentip="Delete linked shapes",
                    supertip="Delete all linked shapes on all slides.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "delete_linked_shapes", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_replace',
                    label="Replace with reference",
                    image_mso="HyperlinkCreate",
                    screentip="Replace linked shapes",
                    supertip="Replace all linked shapes on all slides with the reference shape (by default the selected shape).",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "replace_with_this", shapes=True, context=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_search',
                    label="Search for more shapes",
                    image_mso="FindTag",
                    screentip="Search for the same shape on following slides and link it",
                    supertip="Search again for shapes by position and size to add more shapes to this link.",
                    get_enabled=bkt.apps.ppt_shapes_exactly1_selected,
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "find_similar_and_link", shape=True, context=True),
                ),
                ### Custom action
                # bkt.ribbon.Button(
                #     id = 'linked_shapes_xyz',
                #     label="Custom Action",
                #     image_mso="HappyFace",
                #     on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "linked_shapes_xyz", shape=True, context=True),
                #     # get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                # ),
            ]
        ),
        bkt.ribbon.Group(
            id="bkt_linkshapes_unlink_group",
            label = "Remove link",
            children = [
                bkt.ribbon.Button(
                    id = 'linked_shapes_unlink',
                    label="Remove single shape link",
                    image_mso="HyperlinkRemove",
                    screentip="Remove the link of the selected shape",
                    supertip="Removes the link ID from the current shape. All other shapes with the same ID remain linked.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "unlink_shapes", shapes=True),
                ),
                bkt.ribbon.Button(
                    id = 'linked_shapes_unlink_all',
                    label="Dissolve entire shape link",
                    image_mso="HyperlinkRemove",
                    screentip="Remove all shape links",
                    supertip="Removes the link ID from the current shape as well as all linked shapes with the same ID.",
                    on_action=bkt.CallbackLazy("toolbox.models.linkshapes", "LinkedShapes", "unlink_all_shapes", shapes=True, context=True),
                ),
            ]
        ),
    ]
)