# -*- coding: utf-8 -*-
'''
Created on 14.03.2018

@author: fstallmann
'''



# import math
# import logging

import bkt
# from bkt.library import visio


class InnerMargin(bkt.ribbon.RoundingSpinnerBox):
    ### Instance initialization
    attr = 'tmargin'
    
    def __init__(self, **kwargs):
        '''
        attr examples: tmargin, bmargin, rmargin, lmargin
        '''
        
        my_kwargs = dict(
            size_string = '###',
            round_cm = True,
            # convert = 'pt_to_cm',
            get_enabled = bkt.Callback(self.get_enabled, shapes=True)
        )
        my_kwargs.update(kwargs)
        
        super(InnerMargin, self).__init__(**my_kwargs)

    ### Spinner Box callbacks ###

    def get_enabled(self, shapes):
        return len(shapes) > 0
    
    def get_text(self, shapes):
        if len(shapes) > 0:
            return round(getattr(shapes[0], self.attr),2)
        
    def on_change(self, shapes, value):
        for shape in shapes:
            setattr(shape, self.attr, value)
            if InnerMargin.all_equal:
                shape.tmargin = value
                shape.bmargin = value
                shape.rmargin = value
                shape.lmargin = value
    
    ### class methods ###
    
    all_equal = False

    @classmethod
    def toggle_all_equal(cls, pressed):
        cls.all_equal = pressed
    
    ### set margin to 0
    
    @classmethod
    def set_to_0(cls, shapes, context):
        for shape in shapes:
            shape.tmargin = 0
            shape.bmargin = 0
            shape.rmargin = 0
            shape.lmargin = 0


inner_margin_top    = InnerMargin(attr="tmargin",  id='textFrameMargin-top-2',    show_label=False, imageMso='FillDown' , label="Inner margin top",   screentip='Inner margin top',   supertip='Change the top inner margin of the text field to the specified value.')
inner_margin_bottom = InnerMargin(attr="bmargin",  id='textFrameMargin-bottom-2', show_label=False, imageMso='FillUp'   , label="Inner margin bottom",  screentip='Inner margin bottom',  supertip='Change the bottom inner margin of the text field to the specified value.')
inner_margin_left   = InnerMargin(attr="lmargin",  id='textFrameMargin-left-2',   show_label=False, imageMso='FillRight', label="Inner margin left",  screentip='Inner margin left',  supertip='Change the left inner margin of the text field to the specified value.')
inner_margin_right  = InnerMargin(attr="rmargin",  id='textFrameMargin-right-2',  show_label=False, imageMso='FillLeft' , label="Inner margin right", screentip='Inner margin right', supertip='Change the right inner margin of the text field to the specified value.')

innenabstand_gruppe = bkt.ribbon.Group(
    label="Text-field inner margin",
    image_mso='ObjectNudgeRight',
    children=[
    bkt.ribbon.Box(children=[
        bkt.ribbon.LabelControl(label='         \u200b'),
        inner_margin_top,
        bkt.ribbon.LabelControl(label='   \u200b'),
        bkt.ribbon.Button(
            id='textFrameMargin-zero',
            label="=\u202F0",
            screentip="Inner margin to zero",
            supertip="Change the inner margin of the text field on all sides to zero.",
            on_action=bkt.Callback( InnerMargin.set_to_0 )
        )
    ]),
    bkt.ribbon.Box(children=[
        inner_margin_left,
        inner_margin_right,
    ]),
    bkt.ribbon.Box(children=[
        bkt.ribbon.LabelControl(label='         \u200b'),
        inner_margin_bottom,
        bkt.ribbon.LabelControl(label='   \u200b'),
        bkt.ribbon.ToggleButton(
            id='textFrameMargin-equal',
            label="==",
            screentip="Uniform inner margin",
            supertip="When the text-box inner margin of one side is changed, the inner margin of all sides is changed.",
            on_toggle_action=bkt.Callback( InnerMargin.toggle_all_equal)
        )
    ]),
    bkt.ribbon.DialogBoxLauncher(idMso='TextDialog')
])