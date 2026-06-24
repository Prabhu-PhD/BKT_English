# -*- coding: utf-8 -*-
'''
Created on 29.04.2021

@author: fstallmann
'''



import bkt

from ..models import slides_menu


class ContextSlides(object):
    @staticmethod
    def get_children():
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None, 
                children=[
            bkt.ribbon.Button(
                label='Save',
                image_mso='SaveSelectionToTextBoxGallery',
                supertip="Saves the selected slides in a new presentation.",
                on_action=bkt.Callback(slides_menu.SlideMenu.save_slides_dialog)
            ),
            bkt.ribbon.Button(
                label='Send',
                image_mso='FileSendAsAttachment',
                supertip="Sends the selected slides as an e-mail attachment.",
                on_action=bkt.Callback(slides_menu.SlideMenu.send_slides_dialog)
            ),
                    ]
            )