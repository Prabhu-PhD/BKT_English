# -*- coding: utf-8 -*-
'''
Created on 01.08.2022

@author: fstallmann
'''



import bkt

from ..models import text_menu as text

class ContextTextShapes(object):
    @staticmethod
    def get_buttons(shapes):
        if len(shapes) == 1:
            return [
                ### Text operations
                bkt.ribbon.MenuSeparator(title="Text operations"),
                bkt.ribbon.Button(
                    id = 'text_on_shape-context',
                    label = "Split text onto shapes",
                    supertip="Transfers the text content of each selected shape into a separate text shape.",
                    image_mso = "TableCellCustomMarginsDialog",
                    on_action=bkt.Callback(text.TextOnShape.textOutOfShape, shapes=True, slide=True),
                    get_enabled = bkt.Callback(text.TextOnShape.is_outable, shape=True),
                ),
                bkt.ribbon.Button(
                    id = 'decompose_text-context',
                    label = "Split shape text",
                    supertip="Split the selected shapes into several shapes based on the text paragraphs. One shape with the corresponding text is created per paragraph.",
                    image_mso = "TraceDependents",
                    on_action=bkt.Callback(text.SplitTextShapes.splitShapesByParagraphs, shapes=True, context=True),
                    get_enabled = bkt.Callback(text.SplitTextShapes.is_splitable, shape=True),
                ),
            ]
        else:
            return [
                ### Text operations
                bkt.ribbon.MenuSeparator(title="Text operations"),
                bkt.ribbon.Button(
                    id = 'text_in_shape-context',
                    label = "Combine text in shape",
                    supertip="Copies the text of a text shape into the second selected shape and deletes the text shape.",
                    image_mso = "TextBoxInsert",
                    on_action=bkt.Callback(text.TextOnShape.textIntoShape, shapes=True, shapes_min=2),
                    # get_enabled = bkt.Callback(text.TextOnShape.is_mergable, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_min2_selected,
                ),
                bkt.ribbon.Button(
                    id = 'text_on_shape-context',
                    label = "Split text onto shapes",
                    supertip="Transfers the text content of each selected shape into a separate text shape.",
                    image_mso = "TableCellCustomMarginsDialog",
                    on_action=bkt.Callback(text.TextOnShape.textOutOfShape, shapes=True, slide=True),
                    # get_enabled = bkt.Callback(text.TextOnShape.is_outable, shape=True),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'decompose_text-context',
                    label = "Split shape text",
                    supertip="Split the selected shapes into several shapes based on the text paragraphs. One shape with the corresponding text is created per paragraph.",
                    image_mso = "TraceDependents",
                    on_action=bkt.Callback(text.SplitTextShapes.splitShapesByParagraphs, shapes=True, context=True),
                    # get_enabled = bkt.Callback(text.SplitTextShapes.is_splitable, shape=True),
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'compose_text-context',
                    label = "Merge shape text",
                    supertip="Merges the selected shapes into one shape. The text of all shapes is taken over and concatenated.",
                    image_mso = "TracePrecedents",
                    on_action=bkt.Callback(text.SplitTextShapes.joinShapesWithText, shapes=True, shapes_min=2),
                    # get_enabled = bkt.Callback(text.SplitTextShapes.is_joinable, shapes=True),
                    get_enabled = bkt.apps.ppt_shapes_min2_selected,
                ),
                bkt.ribbon.Button(
                    id = 'text_truncate-context',
                    label="Delete shape texts",
                    supertip="Merges the selected shapes into one shape. The text of all shapes is taken over and concatenated.",
                    image_mso='ReviewDeleteMarkup',
                    on_action=bkt.Callback(text.TextPlaceholder.text_truncate, shapes=True),
                    # get_enabled = bkt.Callback(text.SplitTextShapes.is_joinable, shapes=True), #reuse callback from SplitTextShapes
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Button(
                    id = 'text_replace-context',
                    label="Replace shape texts…",
                    supertip="Replace the text of all selected shapes with text entered in the dialog box.",
                    image_mso='ReplaceDialog',
                    on_action=bkt.Callback(text.TextPlaceholder.text_replace, shapes=True, presentation=True),
                    # get_enabled = bkt.Callback(text.SplitTextShapes.is_joinable, shapes=True), #reuse callback from SplitTextShapes
                    get_enabled = bkt.apps.ppt_shapes_or_text_selected,
                ),
            ]