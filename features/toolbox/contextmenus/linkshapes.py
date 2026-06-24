# -*- coding: utf-8 -*-
'''
Created on 29.04.2021

@author: fstallmann
'''



import bkt

from ..models.linkshapes import LinkedShapes


class ContextLinkedShapes(object):
    @staticmethod
    def get_buttons(shapes):
        if len(shapes) == 1 and LinkedShapes.not_is_linked_shape(shapes[0]):
            return [
                bkt.ribbon.MenuSeparator(title="Linked shapes"),
                bkt.ribbon.Button(
                    label="Search and link shape on following slides…",
                    image_mso="ShapesDuplicate",
                    on_action=bkt.Callback(LinkedShapes.find_similar_and_link, shape=True, context=True),
                    # get_visible=bkt.Callback(LinkedShapes.not_is_linked_shape, shape=True),
                ),
                bkt.ribbon.Button(
                    label="Copy and link shape on following slides…",
                    image_mso="FindTag",
                    on_action=bkt.Callback(LinkedShapes.copy_to_all, shape=True, context=True),
                    # get_visible=bkt.Callback(LinkedShapes.not_is_linked_shape, shape=True),
                ),
            ]
        
        if len(shapes) == 1 and LinkedShapes.is_linked_shape(shapes[0]):
            return [
                bkt.ribbon.MenuSeparator(title="Linked shapes"),
                # bkt.ribbon.SplitButton(
                #     get_visible=bkt.Callback(LinkedShapes.is_linked_shape, shape=True),
                #     children=[
                #         bkt.ribbon.Button(
                #             label="Match linked shapes",
                #             image_mso="HyperlinkCreate",
                #             on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shapes=True, context=True),
                #         ),
                #         bkt.ribbon.Menu(children=[
                            bkt.ribbon.Button(
                                label="Match everything",
                                # image_mso="GroupUpdate",
                                image_mso='HyperlinkCreate',
                                on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shapes=True, context=True),
                            ),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(
                                label="Match position",
                                image_mso="ControlAlignToGrid",
                                on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shapes=True, context=True),
                            ),
                            bkt.ribbon.Button(
                                label="Match size",
                                image_mso="SizeToControlHeightAndWidth",
                                on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shapes=True, context=True),
                            ),
                            bkt.ribbon.Button(
                                label="Match formatting",
                                image_mso="FormatPainter",
                                on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shapes=True, context=True),
                            ),
                            bkt.ribbon.Button(
                                label="Match text",
                                image_mso="TextBoxInsert",
                                on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shapes=True, context=True),
                            ),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(
                                label="Bring to front",
                                image_mso="ObjectBringToFront",
                                on_action=bkt.Callback(LinkedShapes.linked_shapes_tofront, shapes=True, context=True),
                            ),
                            bkt.ribbon.Button(
                                label="Send to back",
                                image_mso="ObjectSendToBack",
                                on_action=bkt.Callback(LinkedShapes.linked_shapes_toback, shapes=True, context=True),
                            ),
                            bkt.ribbon.MenuSeparator(),
                            bkt.ribbon.Button(
                                label="Delete others",
                                image_mso="HyperlinkRemove",
                                on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shapes=True, context=True),
                            ),
                            bkt.ribbon.Button(
                                label="Replace others with this one",
                                image_mso="HyperlinkCreate",
                                on_action=bkt.Callback(LinkedShapes.replace_with_this, shapes=True, context=True),
                            ),
                #         ])
                #     ]
                # ),
            ]
        return []

    @staticmethod
    def get_children_create():
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
            bkt.ribbon.Button(
                label="Search for similar shapes…",
                on_action=bkt.Callback(LinkedShapes.find_similar_and_link, shape=True, context=True),
                get_visible=bkt.Callback(LinkedShapes.not_is_linked_shape, shape=True),
            ),
            bkt.ribbon.Button(
                label="Copy this shape…",
                on_action=bkt.Callback(LinkedShapes.copy_to_all, shape=True, context=True),
                get_visible=bkt.Callback(LinkedShapes.not_is_linked_shape, shape=True),
            ),
                    ]
            )
            
    @staticmethod
    def get_children_align():
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
            bkt.ribbon.Button(
                label="Match everything",
                # image_mso="GroupUpdate",
                image_mso='HyperlinkCreate',
                on_action=bkt.Callback(LinkedShapes.equalize_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                label="Match position",
                image_mso="ControlAlignToGrid",
                on_action=bkt.Callback(LinkedShapes.align_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="Match size",
                image_mso="SizeToControlHeightAndWidth",
                on_action=bkt.Callback(LinkedShapes.size_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="Match formatting",
                image_mso="FormatPainter",
                on_action=bkt.Callback(LinkedShapes.format_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="Match text",
                image_mso="TextBoxInsert",
                on_action=bkt.Callback(LinkedShapes.text_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                label="Bring to front",
                image_mso="ObjectBringToFront",
                on_action=bkt.Callback(LinkedShapes.linked_shapes_tofront, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="Send to back",
                image_mso="ObjectSendToBack",
                on_action=bkt.Callback(LinkedShapes.linked_shapes_toback, shapes=True, context=True),
            ),
            bkt.ribbon.MenuSeparator(),
            bkt.ribbon.Button(
                label="Delete others",
                image_mso="HyperlinkRemove",
                on_action=bkt.Callback(LinkedShapes.delete_linked_shapes, shapes=True, context=True),
            ),
            bkt.ribbon.Button(
                label="Replace others with this one",
                image_mso="HyperlinkCreate",
                on_action=bkt.Callback(LinkedShapes.replace_with_this, shapes=True, context=True),
            ),
                    ]
            )