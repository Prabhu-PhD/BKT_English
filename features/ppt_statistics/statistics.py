# -*- coding: utf-8 -*-
'''

@author: fstallmann
'''



import bkt


class Statistics(object):
    
    @staticmethod
    def show_dialog(context):
        from .dialog import StatisticsWindow
        # StatisticsWindow.create_and_show_dialog(context)
        dialog = StatisticsWindow(context)
        dialog.show_dialog(False)


statistik_gruppe = bkt.ribbon.Group(
    id="bkt_statistics_group",
    label='Statistics',
    supertip="Enables displaying simple statistics for a quick check of number-heavy slides. The `ppt_statistics` feature must be installed.",
    image_mso='RecordsTotals',
    children = [
        bkt.ribbon.Button(
            label="Load statistics",
            image_mso='RecordsTotals',
            size="large",
            supertip="Opens a dialog showing the number of selected shapes, numbers, sum of the numbers, number of characters, words, lines and paragraphs.",
            on_action=bkt.Callback(Statistics.show_dialog),
        ),
    ]
)


bkt.powerpoint.add_tab(bkt.ribbon.Tab(
    id="bkt_powerpoint_toolbox_extensions",
    insert_before_mso="TabHome",
    label='Toolbox 3/3',
    # get_visible defaults to False during async-startup
    get_visible=bkt.Callback(lambda:True),
    children = [
        statistik_gruppe
    ]
), extend=True)

