# -*- coding: utf-8 -*-
'''
Created on 29.04.2021

@author: fstallmann
'''



import bkt

# from .. import arrange
# from .. import shapes as mod_shapes

# from ..models import processshapes
from ..models import arrange_menu
from ..models import shapes_menu

from .stateshapes import ContextStateShapes
from .linkshapes import ContextLinkedShapes
from .harvey import ContextHarveyShapes
from .text import ContextTextShapes
from .process import ContextProcessShapes


class ObjectsGroup(object):
    @staticmethod
    def get_children(shapes):

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=   ContextStateShapes.get_buttons(shapes) + 
                            ContextHarveyShapes.get_buttons(shapes) + 
                            ContextProcessShapes.get_buttons(shapes) + 
                            ContextLinkedShapes.get_buttons(shapes) + 
                            ContextTextShapes.get_buttons(shapes) + [
            # Grouping functions
            bkt.ribbon.MenuSeparator(title="Grouping"),
            bkt.ribbon.Button(
                id='add_into_group-context',
                label="Insert into group",
                supertip="If the first or last selected shape is a group, all other shapes are inserted into this group. Otherwise all shapes are grouped.",
                image_mso="ObjectsRegroup",
                on_action=bkt.Callback(arrange_menu.GroupsMore.add_into_group, shapes=True),
                get_enabled = bkt.Callback(arrange_menu.GroupsMore.visible_add_into_group, shapes=True),
            ),
            bkt.ribbon.Button(
                id='remove_from_group-context',
                label="Detach from group",
                supertip="The selected shapes are detached from the current group without changing the group.",
                image_mso="ObjectsUngroup",
                on_action=bkt.Callback(arrange_menu.GroupsMore.remove_from_group, shapes=True),
                get_visible = bkt.Callback(arrange_menu.GroupsMore.visible_remove_from_group, shapes=True),
            ),
        ]
        )


class ShapeFreeform(object):
    @staticmethod
    def get_children(shape):
        shapes = [shape]
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children= [
        ### Shape connectors
        bkt.ribbon.Button(
            id = 'connector_update-context',
            label = "Reconnect connection area",
            supertip="Update the connection area after the connected shapes have changed.",
            image = "ConnectorUpdate",
            on_action=bkt.Callback(shapes_menu.ShapeConnectors.update_connector_shape, context=True, shape=True),
            get_enabled = bkt.Callback(shapes_menu.ShapeConnectors.is_connector, shape=True),
        ),
        ] + ContextLinkedShapes.get_buttons(shapes) + 
            ContextProcessShapes.get_buttons(shapes) + 
            ContextTextShapes.get_buttons(shapes)
        )




class Shape(object):
    @staticmethod
    def get_children(shape):
        shapes = [shape]
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=   ContextStateShapes.get_buttons(shapes) + 
                            ContextLinkedShapes.get_buttons(shapes) + 
                            ContextTextShapes.get_buttons(shapes) + [
            # Grouping functions
            bkt.ribbon.MenuSeparator(title="Grouping"),
            bkt.ribbon.Button(
                id='remove_from_group-context',
                label="Detach from group",
                supertip="The selected shapes are detached from the current group without changing the group.",
                image_mso="ObjectsUngroup",
                on_action=bkt.Callback(arrange_menu.GroupsMore.remove_from_group, shapes=True),
                get_visible = bkt.Callback(arrange_menu.GroupsMore.visible_remove_from_group, shapes=True),
            ),
        ]
        )




class Picture(object):
    @staticmethod
    def get_children(shape):
        shapes = [shape]
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children= ContextLinkedShapes.get_buttons(shapes)
        )