# -*- coding: utf-8 -*-
'''
Created on 06.07.2016

@author: rdebeerst
'''

import bkt
import bkt.library.powerpoint as pplib



class Characters(object):
    @staticmethod
    def symbol_insert(context):
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            Characters.add_protected_hyphen(context.app.ActiveWindow.Selection)
        elif bkt.get_key_state(bkt.KeyCodes.CTRL):
            Characters.add_protected_space(context.app.ActiveWindow.Selection)
        else:
            context.app.commandbars.ExecuteMso("SymbolInsert")

    
    ### TYPOGRAPHY ###
    @staticmethod
    def add_protected_hyphen(selection):
        selection.TextRange2.text='\xad'

    @staticmethod
    def add_protected_space(selection):
        selection.TextRange2.text='\xa0'

    @staticmethod
    def add_protected_narrow_space(selection):
        selection.TextRange2.text='\u202F'

    ### TYPOGRAPHY ###
    typography = [
        (None, '\xbb', "Linkes Guillemets"),
        (None, '\xab', "Rechtes Guillemets"),
        (None, '\xb6', "Paragraph"),
        (None, '\u2026', "Auslassungspunkte (Ellipse)", "Auslassungspunkte sind drei kurz aufeinanderfolgende Punkte. Meistens zeigen diese eine Ellipse (Auslassung eines Textteils) an."),
        (None, '\u2013', "Gedankenstrich (Halbgeviertstrich/En-Dash)", "Ein Gedankenstrich (sog. Halbgeviertstrich) wie er von Office teilweise automatisch gesetzt wird. Verwendet als Bis-Strich oder Streckenstrich."),
        (None, '\u2014', "Waagerechter Strich (Geviertstrich/Em-Dash)"),
        (None, '\u2020', "Kreuz"),
        (None, '\u2021', "Doppelkreuz"),
        (None, '\u25A0', "Schwarzes Quadrat"),
        (None, '\u25A1', "Weißes Quadrat"),
        (None, '\u2423', "Leerzeichen-Symbol"),
        (None, '\xa9',   "Copyright"),
        (None, '\xae',   "Registered Trade Mark"),
        (None, '\u2122', "Trade Mark"),
        (None, '\u2030', "Per mil"),
        (None, '\u20AC', "Euro"),
        (None, '\u1E9E', "Großes Eszett"),
    ]

    ### MATH ###
    math = [
        (None, '\xb1',   "Plus-Minus-Zeichen", "Ein Plus-Minus-Zeichen einfügen."),
        (None, '\u2212', "Echtes Minuszeichen", "Ein echtes Minuszeichen (kein Bindestrich) einfügen."),
        (None, '\xd7',   "Echtes Malzeichen (Kreuz)", "Ein echtes Kreuz-Multiplikatorzeichen einfügen."),
        (None, '\u22c5', "Echtes Malzeichen (Punkt)", "Ein echtes Punkt-Multiplikatorzeichen einfügen."),
        (None, '\u2044', "Echter Bruchstrich", "Einen echten Bruchstrich (ähnlich Schrägstrich) einfügen."),
        (None, '\u2248', "Ungefähr Gleich", "Ein Ungefähr Gleich Zeichen einfügen."),
        (None, '\u2260', "Ungleich", "Ein Ungleich-Zeichen einfügen."),
        (None, '\u2206', "Delta", "Ein Deltazeichen einfügen."), #alt: \u0394 griechisches Delta
        (None, '\u2300', "Mittelwert/Durchmesser", "Ein Durchmesserzeichen bzw. Durchschnittszeichen einfügen."), #alt: \xD8 leere menge
        (None, '\u2211', "Summenzeichen", "Ein Summenzeichen einfügen."),
        (None, '\u221A', "Wurzelzeichen", "Ein Wurzelzeichen einfügen."),
        (None, '\u221E', "Unendlich-Zeichen", "Ein Unendlich-Zeichen (liegende Acht) einfügen."),
        (None, '\u2264', "Kleiner-Gleich", "Ein kleiner oder gleich Zeichen einfügen."),
        (None, '\u2265', "Größer-Gleich", "Ein größer oder gleich Zeichen einfügen."),
    ]

    ### LIST ###
    lists = [
        (None, '\u2022', "Aufzählungszeichen (Kreis)", "Ein Aufzählungszeichen (schwarzer Punkt) einfügen."),
        (None, '\u2023', "Aufzählungszeichen (Dreieck)", "Ein Aufzählungszeichen (schwarzes Dreieck) einfügen."),
        (None, '\u25AA', "Aufzählungszeichen (Quadrat)", "Ein Aufzählungszeichen (schwarzes Quadrat) einfügen."),
        (None, '\u2043', "Aufzählungszeichen (Strich)", "Ein Aufzählungszeichen (Bindestrich) einfügen."),
        (None, '\u2212', "Echtes Minuszeichen", "Ein echtes Minuszeichen (kein Bindestrich) einfügen."),
        (None, '\x2b',   "Pluszeichen", "Ein Pluszeichen einfügen."),
        (None, '\u2610', "Box leer"),
        (None, '\u2611', "Box Häkchen"),
        (None, '\u2612', "Box Kreuzchen"),
        ("Wingdings", 'J', "Wingdings Smiley gut"),
        ("Wingdings", 'K', "Wingdings Smiley neutral"),
        ("Wingdings", 'L', "Wingdings Smiley schlecht"),
        (None, '\u2713', "Häkchen", "Ein Häkchen-Symbol einfügen."),
        (None, '\u2714', "Häkchen fett", "Ein fettes Häkchen-Symbol einfügen."),
        (None, '\u2717', "Kreuzchen geschwungen", "Ein geschwungenes Kreuzchen (passend zu Häkchen) einfügen."),
        (None, '\u2718', "Kreuzchen geschwungen fett", "Ein fettes geschwungenes Kreuzchen (passend zu Häkchen) einfügen."),
        (None, '\u2715', "Kreuzchen symmetrisch", "Ein symmetrisches Kreuzchen (ähnlich Malzeichen) einfügen."),
        (None, '\u2716', "Kreuzchen symmetrisch fett", "Ein fettes symmetrisches Kreuzchen (ähnlich Malzeichen) einfügen."),
        (None, '\u2605', "Stern schwarz"),
        (None, '\u2606', "Stern weiß"),
        (None, '\u261B', "Zeigefinger schwarz"),
        (None, '\u261E', "Zeigefinger weiß"),
        ("Wingdings", 'C', "Wingdings Thumbs-Up"),
        ("Wingdings", 'D', "Wingdings Thumbs-Down"),
        ### Default list symbol:
        # ("Arial",       u'\u2022', "Arial Bullet"),
        ("Courier New", 'o', "Courier New Kreis"),
        ("Wingdings",   '\xa7', "Wingdings Rechteck"),
        ("Symbol",      '-', "Symbol Strich"),
        ("Wingdings",   'v', "Wingdings Stern"),
        ("Wingdings",   '\xd8', "Wingdings Pfeil"),
        ("Wingdings",   '\xfc', "Wingdings Häckchen"),
    ]

    ### ARROWS ###
    arrows = [
        (None, '\u2192', "Pfeil rechts"),
        (None, '\u2190', "Pfeil links"),
        (None, '\u2191', "Pfeil oben"),
        (None, '\u2193', "Pfeil unten"),
        (None, '\u2194', "Pfeil links und rechts"),
        (None, '\u21C4', "Pfeil links und rechts"),
        (None, '\u2197', "Pfeil rechts oben"),
        (None, '\u2196', "Pfeil links oben"),
        (None, '\u2198', "Pfeil rechts unten"),
        (None, '\u2199', "Pfeil links unten"),
        (None, '\u2195', "Pfeil oben und unten"),
        (None, '\u21C5', "Pfeil oben und unten"),
        (None, '\u21E8', "Weißer Pfeil rechts"),
        (None, '\u21E6', "Weißer Pfeil links"),
        (None, '\u21E7', "Weißer Pfeil oben"),
        (None, '\u21E9', "Weißer Pfeil unten"),
        (None, '\u21AF', "Blitz"),
        (None, '\u21BA', "Kreispfeil gegen den Uhrzeigersinn"),
    ]

    @staticmethod
    def text_selection(selection):
        return selection.Type == 3


    @classmethod
    def get_text_fontawesome(cls):
        from .fontawesome import Fontawesome

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=Fontawesome.get_symbol_galleries()
            )

    @classmethod
    def get_text_fontawesome_exclusion(cls):
        from .fontawesome import Fontawesome

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=Fontawesome.get_exclusions()
            )
        
    @classmethod
    def get_text_unicodefont(cls):
        def _unicode_font_button(font):
            return bkt.ribbon.ToggleButton(
                label=font,
                screentip="Unicode font"+font,
                supertip=font+" als Unicode-Schriftart verwenden.",
                on_toggle_action=bkt.Callback(lambda pressed: pplib.PPTSymbolsSettings.switch_unicode_font(font)),
                get_pressed=bkt.Callback(lambda: pplib.PPTSymbolsSettings.unicode_font == font),
                get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image(font, "\u2192"))
            )

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=[
                            bkt.ribbon.ToggleButton(
                                label='Theme font (default)',
                                screentip="Unicode font matches theme font",
                                supertip="No special Unicode font is used, but the theme's default font.",
                                on_toggle_action=bkt.Callback(lambda pressed: pplib.PPTSymbolsSettings.switch_unicode_font(None)),
                                get_pressed=bkt.Callback(lambda: pplib.PPTSymbolsSettings.unicode_font is None),
                            ),
                        ] + [
                            _unicode_font_button(font)
                            for font in pplib.PPTSymbolsSettings.UNICODE_FONTS
                        ]
            )

    @classmethod
    def get_text_menu(cls):
        recent_symbols = pplib.PPTSymbolsGalleryRecent(
            id="symbols_recent_gallery",
            label="Recently used",
        )

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                # id=None,
                id="symbols_splitbutton",
                label="Symbol menu",
                show_label=False,
                image_mso="SymbolInsert",
                screentip="Symbol",
                supertip="Opens a menu with various galleries for quickly inserting symbols and special characters.",
                children=[
                    bkt.mso.button.SymbolInsert,
                    bkt.ribbon.MenuSeparator(title="Recently used"),
                    recent_symbols.get_index_as_button(2),
                    recent_symbols.get_index_as_button(1),
                    recent_symbols.get_index_as_button(0),
                    bkt.ribbon.MenuSeparator(title="Symbols"),
                    bkt.ribbon.Button(
                        id='symbols_add_protected_hyphen',
                        label='Non-breaking hyphen',
                        supertip='An optional (conditional) hyphen is a symbol for optional hyphenation. The hyphen appears only at the line end and otherwise stays invisible.',
                        on_action=bkt.Callback(cls.add_protected_hyphen, selection=True),
                        get_enabled = bkt.Callback(cls.text_selection, selection=True),
                        get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image("Arial", "-"))
                    ),
                    bkt.ribbon.Button(
                        id='symbols_add_protected_space',
                        label='Non-breaking space',
                        supertip='A non-breaking space does not allow a line break.',
                        on_action=bkt.Callback(cls.add_protected_space, selection=True),
                        get_enabled = bkt.Callback(cls.text_selection, selection=True),
                        get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image("Arial", "\u23B5")) #alt: 2423
                    ),
                    bkt.ribbon.Button(
                        id='symbols_add_protected_narrow_space',
                        label='Narrow non-breaking space',
                        supertip='A narrow non-breaking space does not allow a line break and is used e.g. between letters of abbreviations.',
                        on_action=bkt.Callback(cls.add_protected_narrow_space, selection=True),
                        get_enabled = bkt.Callback(cls.text_selection, selection=True),
                        get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image("Arial", "\u02FD"))
                    ),

                    pplib.PPTSymbolsGallery(
                        id="symbols_typo_gallery",
                        label="Typography symbols",
                        supertip="Insert various typography symbols",
                        symbols = cls.typography,
                    ),
                    bkt.ribbon.MenuSeparator(),

                    pplib.PPTSymbolsGallery(
                        id="symbols_math_gallery",
                        label="Math symbols",
                        supertip="Insert various math symbols",
                        symbols = cls.math,
                    ),
                    pplib.PPTSymbolsGallery(
                        id="symbols_lists_gallery",
                        label="List symbols",
                        supertip="Insert various list symbols",
                        symbols = cls.lists,
                    ),
                    pplib.PPTSymbolsGallery(
                        id="symbols_arrow_gallery",
                        label="Arrows",
                        supertip="Insert various arrows",
                        symbols = cls.arrows,
                    ),
                # ] + fontawesome.symbol_galleries + [
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.DynamicMenu(
                        id="symbols_icon_fonts",
                        label="Icon fonts",
                        supertip="Shows icons for available icon fonts that can be inserted as a text symbol or graphic.\n\nNote: The icon fonts must be installed on the computer.",
                        image_mso="Call",
                        get_content = bkt.Callback(cls.get_text_fontawesome)
                    ),
                    bkt.ribbon.MenuSeparator(title="Settings"),
                    bkt.ribbon.DynamicMenu(
                        label="Choose Unicode font",
                        image_mso='FontDialogPowerPoint',
                        supertip="Unicode characters can be inserted either with the default font or with a special Unicode font. This can be selected here.",
                        get_content = bkt.Callback(cls.get_text_unicodefont)
                    ),
                    bkt.ribbon.DynamicMenu(
                        label="Exclude icon fonts",
                        # image_mso='FontDialogPowerPoint',
                        supertip="Do not show certain icon fonts and exclude them from the search.",
                        get_content = bkt.Callback(cls.get_text_fontawesome_exclusion)
                    ),
                    bkt.ribbon.ToggleButton(
                        label='Insert as text (default)',
                        image_mso='TextTool',
                        screentip="Toggle insert as text",
                        supertip='If no text is selected and this option is enabled, the symbol is inserted as a Unicode character. This is the default when no key is pressed.',
                        on_toggle_action=bkt.Callback(pplib.PPTSymbolsSettings.switch_convert_into_text),
                        get_pressed=bkt.Callback(pplib.PPTSymbolsSettings.convert_into_text), #convert into text is a function!
                    ),
                    bkt.ribbon.ToggleButton(
                        label='Insert as shapes [Shift]',
                        image_mso='TextEffectTransformGallery',
                        screentip="Toggle insert as shape",
                        supertip='If no text is selected and this option is enabled, the symbol is converted into a shape. This also works when clicking a symbol while holding the Shift key.',
                        on_toggle_action=bkt.Callback(pplib.PPTSymbolsSettings.switch_convert_into_shape),
                        get_pressed=bkt.Callback(lambda: pplib.PPTSymbolsSettings.convert_into_shape),
                    ),
                    bkt.ribbon.ToggleButton(
                        label='Insert as image [Ctrl]',
                        image_mso='PictureRecolorBlackAndWhite',
                        screentip="Toggle insert as image",
                        supertip='If no text is selected and this option is enabled, the symbol is inserted as a raster graphic. This also works when clicking a symbol while holding the Ctrl key.',
                        on_toggle_action=bkt.Callback(pplib.PPTSymbolsSettings.switch_convert_into_bitmap),
                        get_pressed=bkt.Callback(lambda: pplib.PPTSymbolsSettings.convert_into_bitmap),
                    ),
                ]
            )


#TODO: Use MouseKeyHook to register Strg+-/Space key combination in order to add special chars

#OPTION 1: Dynamic Menu - Cons: Buttons (e.g. hyphen) cannot be added to quick access toolbar
# symbol_insert_splitbutton = bkt.ribbon.DynamicMenu(
#     id="symbols_splitbutton",
#     label="Symbol menu",
#     show_label=False,
#     image_mso="SymbolInsert",
#     screentip="Symbol",
#     supertip="Opens a menu with various galleries for quickly inserting symbols and special characters.",
#     get_content = bkt.Callback(
#         Characters.get_text_menu
#     ),
# )

#OPTION 2: Splitbutton with regular menu - Cons: Splitbutton is not intuitive and not compatible with dynamic menu
# symbol_insert_splitbutton = bkt.ribbon.SplitButton(
#     id="symbols_splitbutton",
#     show_label=False,
#     children=[
#         bkt.ribbon.Button(
#             label="Symbol",
#             image_mso="SymbolInsert",
#             screentip="Symbol",
#             supertip="Opens the dialog for inserting symbols.\n\nWith the Shift key held, a non-breaking hyphen is inserted directly.\n\nWith the Ctrl key held, a non-breaking space is inserted.",
#             on_action=bkt.Callback(Characters.symbol_insert, context=True),
#             get_enabled=bkt.Callback(lambda context: context.app.commandbars.GetEnabledMso("SymbolInsert"), context=True),
#         ),
#         #bkt.mso.button.SymbolInsert,
#         # character_menu
#         Characters.get_text_menu()
#     ]
# )

#OPTION 3: Regular menu with dynamic menu only for icons fonts
symbol_insert_splitbutton = Characters.get_text_menu()





class InnerMargin(pplib.TextframeSpinnerBox):
    
    ### class methods ###
    
    all_equal = False

    @classmethod
    def toggle_all_equal(cls, pressed):
        cls.all_equal = pressed

    @classmethod
    def get_all_equal(cls):
        return cls.all_equal
    
    ### set margin to 0
    
    @classmethod
    def set_to_0(cls, shapes, context):
        for textframe in pplib.iterate_shape_textframes(shapes):
            textframe.MarginTop    = 0
            textframe.MarginBottom = 0
            textframe.MarginLeft   = 0
            textframe.MarginRight  = 0


    ### Setter methods ###
    
    def set_attr_for_textframe(self, textframe, value):
        setattr(textframe, self.attr, value)
        if InnerMargin.all_equal:
            textframe.MarginTop    = value
            textframe.MarginBottom = value
            textframe.MarginLeft   = value
            textframe.MarginRight  = value



inner_margin_top    = InnerMargin(attr="MarginTop",    id='textFrameMargin-top-2',    image_button=False, show_label=False, image_mso='FillDown' , label="Inner margin top",   screentip='Inner margin top',   supertip='Change the top inner margin of the text field to the specified value (in cm).')
inner_margin_bottom = InnerMargin(attr="MarginBottom", id='textFrameMargin-bottom-2', image_button=False, show_label=False, image_mso='FillUp'   , label="Inner margin bottom",  screentip='Inner margin bottom',  supertip='Change the bottom inner margin of the text field to the specified value (in cm).')
inner_margin_left   = InnerMargin(attr="MarginLeft",   id='textFrameMargin-left-2',   image_button=False, show_label=False, image_mso='FillRight', label="Inner margin left",  screentip='Inner margin left',  supertip='Change the left inner margin of the text field to the specified value (in cm).')
inner_margin_right  = InnerMargin(attr="MarginRight",  id='textFrameMargin-right-2',  image_button=False, show_label=False, image_mso='FillLeft' , label="Inner margin right", screentip='Inner margin right', supertip='Change the right inner margin of the text field to the specified value (in cm).')



innenabstand_gruppe = bkt.ribbon.Group(
    id="bkt_innermargin_group",
    label="Text-field inner margin",
    image_mso='ObjectNudgeRight',
    children=[
    bkt.ribbon.Box(id='textFrameMargin-box-top', children=[
        bkt.ribbon.LabelControl(id='textFrameMargin-label-top', label='         \u200b'),
        #create_margin_spinner('MarginTop', imageMso='ObjectNudgeDown'),
        inner_margin_top,
        bkt.ribbon.LabelControl(label='   \u200b'),
        bkt.ribbon.Button(
            id='textFrameMargin-zero',
            label="=\u202F0",
            screentip="Inner margin to zero",
            supertip="Change the inner margin of the text field on all sides to zero.",
            on_action=bkt.Callback( InnerMargin.set_to_0, shapes=True, context=True )
        )
    ]),
    bkt.ribbon.Box(id='textFrameMargin-box-LR', children=[
        #create_margin_spinner('MarginLeft',  imageMso='ObjectNudgeRight'),
        #create_margin_spinner('MarginRight', imageMso='ObjectNudgeLeft')
        inner_margin_left,
        #bkt.ribbon.LabelControl(label=u' '),
        inner_margin_right,
    ]),
    bkt.ribbon.Box(id='textFrameMargin-box-bottom', children=[
        bkt.ribbon.LabelControl(id='textFrameMargin-label-bottom', label='         \u200b'),
        #create_margin_spinner('MarginBottom', imageMso='ObjectNudgeUp'),
        inner_margin_bottom,
        bkt.ribbon.LabelControl(label='   \u200b'),
        bkt.ribbon.ToggleButton(
            id='textFrameMargin-equal',
            label="==",
            screentip="Uniform inner margin",
            supertip="When the text-box inner margin of one side is changed, the inner margin of all sides is changed.",
            on_toggle_action=bkt.Callback( InnerMargin.toggle_all_equal ),
            get_pressed=bkt.Callback( InnerMargin.get_all_equal )
        )
    ]),
    bkt.ribbon.DialogBoxLauncher(idMso='TextAlignMoreOptionsDialog')
    #bkt.ribbon.DialogBoxLauncher(idMso='WordArtFormatDialog')
])



class ParSpaceBefore(pplib.ParagraphFormatSpinnerBox):
    attr = 'SpaceBefore'
    _attributes = dict(
        label="Paragraph spacing above",
        image_mso='WordOpenParaAbove',
        screentip="Top paragraph spacing",
        supertip="Change the paragraph spacing before the paragraph to the specified value (either in number of lines or in pt).",
    )

class ParSpaceAfter(pplib.ParagraphFormatSpinnerBox):
    attr = 'SpaceAfter'
    _attributes = dict(
        label="Paragraph spacing below",
        image_mso='WordOpenParaBelow',
        screentip="Bottom paragraph spacing",
        supertip="Change the paragraph spacing after the paragraph to the specified value (either in number of lines or in pt).",
    )

class ParSpaceWithin(pplib.ParagraphFormatSpinnerBox):
    attr = 'SpaceWithin'
    _attributes = dict(
        label="Row spacing",
        image_mso='LineSpacing',
        screentip="Row spacing",
        supertip="Change the line spacing (either in number of lines or in pt).",
        fallback_value = 1,
    )

class ParFirstLineIndent(pplib.ParagraphFormatSpinnerBox):
    attr = 'FirstLineIndent'
    _attributes = dict(
        label="Indent first line",
        image='first_line_indent',
        screentip="Special indent",
        supertip="Change the special indent (first line, hanging) to the specified value (in cm).",
    )

class ParLeftIndent(pplib.ParagraphFormatSpinnerBox):
    attr = 'LeftIndent'
    _attributes = dict(
        label="Indent left",
        image_mso='ParagraphIndentLeft',
        screentip="Paragraph indent left",
        supertip="Change the left paragraph indent to the specified value (in cm).",
    )

class ParRightIndent(pplib.ParagraphFormatSpinnerBox):
    attr = 'RightIndent'
    _attributes = dict(
        label="Indent right",
        image_mso='ParagraphIndentRight',
        screentip="Paragraph indent right",
        supertip="Change the right paragraph indent to the specified value (in cm).",
    )


class Absatz(object):

    @staticmethod
    def set_word_wrap(shapes, pressed):
        for textframe in pplib.iterate_shape_textframes(shapes):
            try:
                textframe.WordWrap = -1 if pressed else 0
            except:
                continue

    @staticmethod
    def get_word_wrap(shapes):
        for textframe in pplib.iterate_shape_textframes(shapes):
            try:
                return (textframe.WordWrap == -1) # msoTrue
            except:
                continue
        return None


    @staticmethod
    def set_auto_size(shapes, pressed):
        for textframe in pplib.iterate_shape_textframes(shapes):
            try:
                textframe.AutoSize = 1 if pressed else 0
                # 1 = ppAutoSizeShapeToFitText
                # 0 = ppAutoSizeNone
            except:
                continue

    @staticmethod
    def get_auto_size(shapes):
        for textframe in pplib.iterate_shape_textframes(shapes):
            try:
                return (textframe.AutoSize == 1)
            except:
                continue
        return None

    # def set_par_indent(self, shapes, value):
    #     # pt_value = cm_to_pt(value)
    #     # delta = pt_value - shapes[0].TextFrame.Ruler.Levels(1).LeftMargin
    #     for shape in shapes:
    #         shape.TextFrame.Ruler.Levels(1).LeftMargin = cm_to_pt(value)
    #         # shape.TextFrame.Ruler.Levels(1).LeftMargin = pt_value
    #         # shape.TextFrame.Ruler.Levels(1).LeftMargin  = shp.TextFrame.Ruler.Levels(1).LeftMargin + delta
    #
    # def get_par_indent(self, shapes):
    #     return round(pt_to_cm(shapes[0].TextFrame.Ruler.Levels(1).LeftMargin), 2)

    # @staticmethod
    # def set_par_sep_before(shapes, selection, value):
    #     value = max(0,value)
    #     if selection.Type == 2:
    #         # shapes selected
    #         for shape in shapes:
    #             # distance in points, not in number of lines
    #             shape.TextFrame.TextRange.ParagraphFormat.LineRuleBefore = 0
    #             # set distance
    #             shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = value
    #     elif selection.Type == 3:
    #         # text selected
    #         selection.TextRange2.ParagraphFormat.LineRuleBefore = 0
    #         selection.TextRange2.ParagraphFormat.SpaceBefore = value 

    # @staticmethod
    # def get_par_sep_before(shapes, selection):
    #     if selection.Type == 2:
    #         # shapes selected
    #         return shapes[0].TextFrame.TextRange.ParagraphFormat.SpaceBefore
    #     elif selection.Type == 3:
    #         # text selected
    #         try:
    #             # produces error if no text is selected
    #             return selection.TextRange2.Paragraphs(1,1).ParagraphFormat.SpaceBefore
    #         except:
    #             return selection.TextRange2.ParagraphFormat.SpaceBefore


    # @staticmethod
    # def set_par_sep_after(shapes, selection, value):
    #     value = max(0,value)
    #     if selection.Type == 2:
    #         # shapes selected
    #         for shape in shapes:
    #             # distance in points, not in number of lines
    #             shape.TextFrame.TextRange.ParagraphFormat.LineRuleAfter = 0
    #             # set distance
    #             shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = value
    #     elif selection.Type == 3:
    #         # text selected
    #         selection.TextRange2.ParagraphFormat.LineRuleAfter = 0
    #         selection.TextRange2.ParagraphFormat.SpaceAfter = value 

    # @staticmethod
    # def get_par_sep_after(shapes, selection):
    #     if selection.Type == 2:
    #         # shapes selected
    #         return shapes[0].TextFrame.TextRange.ParagraphFormat.SpaceAfter
    #     elif selection.Type == 3:
    #         # text selected
    #         try:
    #             # produces error if no text is selected
    #             return selection.TextRange2.Paragraphs(1,1).ParagraphFormat.SpaceAfter
    #         except:
    #             return selection.TextRange2.ParagraphFormat.SpaceAfter
    
    
    # @staticmethod
    # def set_left_indent(shapes, selection, value):
    #     # FIXME: apply text-selection-logic from set_par_sep_after
    #     if type(value) == str:
    #         value = float(value.replace(',', '.'))
    #     value = float(value) / pt_to_cm_factor
        
    #     if selection.Type == 2:
    #         # shapes selected
    #         for shape in shapes:
    #             shape.TextFrame2.TextRange.ParagraphFormat.LeftIndent = value
    #     elif selection.Type == 3:
    #         # text selected
    #         selection.TextRange2.ParagraphFormat.LeftIndent = value


    # @staticmethod
    # def get_left_indent(shapes, selection):
    #     if selection.Type == 2:
    #         # shapes selected
    #         value = shapes[0].TextFrame2.TextRange.ParagraphFormat.LeftIndent
    #     elif selection.Type == 3:
    #         # text selected
    #         try:
    #             # produces error if no text is selected
    #             value = selection.TextRange2.Paragraphs(1,1).ParagraphFormat.LeftIndent 
    #         except:
    #             value = selection.TextRange2.ParagraphFormat.LeftIndent 

    #     return round(value * pt_to_cm_factor, 2)
    
    
    # @staticmethod
    # def set_first_line_indent(shapes, selection, value):
    #     if type(value) == str:
    #         value = float(value.replace(',', '.'))
    #     value = float(value) / pt_to_cm_factor
        
    #     if selection.Type == 2:
    #         # shapes selected
    #         for shape in shapes:
    #             shape.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = value
    #     elif selection.Type == 3:
    #         # text selected
    #         selection.TextRange2.ParagraphFormat.FirstLineIndent = value

    # @staticmethod
    # def get_first_line_indent(shapes, selection):
    #     if selection.Type == 2:
    #         # shapes selected
    #         value = shapes[0].TextFrame2.TextRange.ParagraphFormat.FirstLineIndent
    #     elif selection.Type == 3:
    #         # text selected
    #         try:
    #             # produces error if no text is selected
    #             value = selection.TextRange2.Paragraphs(1,1).ParagraphFormat.FirstLineIndent 
    #         except:
    #             value = selection.TextRange2.ParagraphFormat.FirstLineIndent 
        
    #     return round(value * pt_to_cm_factor, 2)



text_menu = bkt.ribbon.Menu(
    label="Draw text field menu",
    supertip="Insert stickers, match bullet points, as well as further text-related functions",
    children=[
        bkt.ribbon.MenuSeparator(title="Insert text shapes"),
        bkt.mso.control.TextBoxInsert,
        bkt.ribbon.DynamicMenu(
            id="sticker_splitbutton", #2023-01-26 not a splitbutton anymore to make it dynamic
            label="Sticker menu",
            supertip="Insert various stickers",
            get_content=bkt.CallbackLazy("toolbox.models.text_menu", "sticker_menu")
        ),
        bkt.ribbon.Button(
            id = 'underlined_textbox',
            label = "Underlined text box",
            image = "underlined_textbox",
            screentip="Insert underlined text box",
            supertip="Insert a text box with a line at the bottom of the shape on the current slide.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextShapes", "addUnderlinedTextbox", slide=True, presentation=True)
        ),
        bkt.ribbon.MenuSeparator(title="Bullets"),
        bkt.ribbon.Button(
            id="bullet_fixing",
            label="Fix bullets",
            supertip="Bullets are corrected. The style is taken from the text placeholder on the master slide. Affects: symbol, symbol/text color, paragraph indent/spacing",
            image_mso='MultilevelListGallery',
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "fix_bullet_style", shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.ColorGallery(
            id = 'bullet_color',
            label='Change color',
            screentip="Change bullet point color",
            supertip="Changes the color of the selected bullet points.",
            on_rgb_color_change = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "set_bullet_color_rgb", selection=True, shapes=True),
            on_theme_color_change = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "set_bullet_theme_color", selection=True, shapes=True),
            get_selected_color = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "get_bullet_color_rgb", selection=True, shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
            children=[
                bkt.ribbon.Button(
                    id="bullet_color_auto",
                    label="Automatic",
                    screentip="Determine bullet point color automatically",
                    supertip="The bullet point color is determined automatically based on the text color.",
                    on_action=bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "set_bullet_color_auto", selection=True, shapes=True),
                    image_mso="ColorBlack",
                ),
            ]
        ),
        bkt.ribbon.SymbolsGallery(
            id="bullet_symbol",
            label="Change symbol",
            screentip="Change bullet point symbol",
            supertip="Changes the symbol of the selected bullet points.",
            symbols = Characters.lists,
            on_symbol_change = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "set_bullet_symbol", selection=True, shapes=True),
            get_selected_symbol = bkt.CallbackLazy("toolbox.models.text_menu", "BulletStyle", "get_bullet_symbol", selection=True, shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected
        ),
        bkt.ribbon.MenuSeparator(title="Text operations"),
        bkt.ribbon.Button(
            id = 'text_in_shape',
            label = "Text in shape",
            image_mso = "TextBoxInsert",
            screentip="Combine text in shape",
            supertip="Copies the text of a text shape into the second selected shape and deletes the text shape.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextOnShape", "textIntoShape", shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.Button(
            id = 'text_on_shape',
            label = "Text onto shape",
            image_mso = "TableCellCustomMarginsDialog",
            screentip="Split text onto shapes",
            supertip="Transfers the text content of each selected shape into a separate text shape.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextOnShape", "textOutOfShape", shapes=True, slide=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id = 'decompose_text',
            label = "Split shape text",
            image_mso = "TraceDependents",
            supertip="Split the selected shapes into several shapes based on the text paragraphs. One shape with the corresponding text is created per paragraph.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "SplitTextShapes", "splitShapesByParagraphs", shapes=True, context=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.Button(
            id = 'compose_text',
            label = "Merge shape text",
            image_mso = "TracePrecedents",
            supertip="Merges the selected shapes into one shape. The text of all shapes is taken over and concatenated.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "SplitTextShapes", "joinShapesWithText", shapes=True, shapes_min=2),
            get_enabled = bkt.apps.ppt_shapes_min2_selected,
        ),
        bkt.ribbon.MenuSeparator(),
        bkt.ribbon.Button(
            id = 'text_truncate',
            label="Delete shape texts",
            image_mso='ReviewDeleteMarkup',
            supertip="Delete the text of all selected shapes.",
            on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_truncate", shapes=True),
            get_enabled = bkt.apps.ppt_shapes_or_text_selected,
        ),
        bkt.ribbon.SplitButton(
            id = 'text_replace_splitbutton',
            get_enabled=bkt.apps.ppt_shapes_or_text_selected,
            children=[
                bkt.ribbon.Button(
                    id = 'text_replace',
                    label="Replace shape texts…",
                    image_mso='ReplaceDialog',
                    supertip="Replace the text of all selected shapes with text entered in the dialog box.",
                    on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_replace", shapes=True, presentation=True),
                    get_enabled=bkt.apps.ppt_shapes_or_text_selected,
                ),
                bkt.ribbon.Menu(label="Replace shape texts menu", supertip="Replace text with predefined standard placeholders", children=[
                    bkt.ribbon.Button(
                        id = 'text_tbd',
                        label="… with »tbd«",
                        image_mso='TextDialog',
                        screentip="Replace text with »tbd«",
                        supertip="Replace the text of all selected shapes with 'tbd'.",
                        on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_tbd", shapes=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'text_lorem',
                        label="… with Lorem ipsum",
                        image_mso='TextDialog',
                        screentip="Replace text with Lorem ipsum",
                        supertip="Replace the text of all selected shapes with long 'Lorem ipsum' placeholder text.",
                        on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_lorem", shapes=True),
                    ),
                    bkt.ribbon.Button(
                        id = 'text_counter',
                        label="… with numbering",
                        image_mso='TextDialog',
                        screentip="Replace text with numbering",
                        supertip="Replace the text of all selected shapes with numbering.",
                        on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_counter", shapes=True),
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'text_replace2',
                        label="… with custom text",
                        image_mso='ReplaceDialog',
                        screentip="Replace text with custom input",
                        supertip="Replace the text of all selected shapes with text entered in the dialog box.",
                        on_action=bkt.CallbackLazy("toolbox.models.text_menu", "TextPlaceholder", "text_replace", shapes=True, presentation=True),
                        get_enabled=bkt.apps.ppt_shapes_or_text_selected,
                    ),
                ])
            ]
        ),
    ]
)


class TextBox(object):
    @staticmethod
    def textbox_insert(context, pressed):
        from .models.text_menu import TextShapes
        if bkt.get_key_state(bkt.KeyCodes.SHIFT):
            TextShapes.addUnderlinedTextbox(context.slide, context.presentation)
        elif bkt.get_key_state(bkt.KeyCodes.CTRL):
            TextShapes.addSticker(context.slide, context.presentation)
        else:
            # NOTE: idMso is different on some machines, see: https://answers.microsoft.com/en-us/msoffice/forum/msoffice_powerpoint-msoffice_custom-mso_2007/powerpoint-2007-textboxinsert-vs/52f12b52-7e1c-4d7c-86a7-bded312437b0
            try:
                context.app.commandbars.ExecuteMso("TextBoxInsert")
            except:
                context.app.commandbars.ExecuteMso("TextBoxInsertHorizontal")
    
    @staticmethod
    def textbox_enabled(context):
        try:
            return context.app.commandbars.GetEnabledMso("TextBoxInsert")
        except:
            return context.app.commandbars.GetEnabledMso("TextBoxInsertHorizontal")
    
    @staticmethod
    def textbox_pressed(context):
        try:
            return context.app.commandbars.GetPressedMso("TextBoxInsert")
        except:
            return context.app.commandbars.GetPressedMso("TextBoxInsertHorizontal")


text_splitbutton = bkt.ribbon.SplitButtonFixed(
    id="textbox_insert_splitbutton",
    show_label=False,
    children=[
        bkt.ribbon.ToggleButton(
            id="textbox_insert",
            label="Draw text field",
            image_mso="TextBoxInsert",
            supertip="Draw a text field at any position.\n\nWith the Shift key held, an underlined text box is inserted.\n\nWith the Ctrl key held, a sticker is inserted.",
            on_toggle_action=bkt.Callback(TextBox.textbox_insert, context=True),
            get_pressed=bkt.Callback(TextBox.textbox_pressed, context=True),
            get_enabled=bkt.Callback(TextBox.textbox_enabled, context=True),
        ),
        # bkt.mso.toggleButton.TextBoxInsert,
        text_menu
    ]
)


paragraph_group = bkt.ribbon.Group(
    id="bkt_paragraph_group",
    label = "Paragraph",
    image_mso='FormattingMarkDropDown',
    children = [
        bkt.ribbon.Menu(
            label="Settings",
            imageMso="FormattingMarkDropDown",
            supertip="Change settings for the text box",
            children = [
                bkt.ribbon.ToggleButton(
                    id = 'wordwrap',
                    label="WordWrap",
                    image_mso="FormattingMarkDropDown",
                    screentip="Wrap text in shape",
                    supertip="Configure the text option to 'Wrap text in shape'.",
                    on_toggle_action=bkt.Callback(Absatz.set_word_wrap , shapes=True),
                    get_pressed=bkt.Callback(Absatz.get_word_wrap , shapes=True),
                    get_enabled = bkt.apps.ppt_selection_contains_textframe,
                ),
                bkt.ribbon.ToggleButton(
                    id = 'autosize',
                    label="AutoSize",
                    image_mso="SmartArtLargerShape",
                    screentip="Adjust the size of the shape",
                    supertip="Configure the text option to 'Resize shape to fit text' or 'Do not autofit'.",
                    on_toggle_action=bkt.Callback(Absatz.set_auto_size , shapes=True),
                    get_pressed=bkt.Callback(Absatz.get_auto_size , shapes=True),
                    get_enabled = bkt.apps.ppt_selection_contains_textframe,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.mso.control.TextAlignMoreOptionsDialog
            ]
        ),
        ParSpaceBefore(
            id = 'par_sep_top',
            show_label=False,
            size_string = '##',
            # label=u"Paragraph spacing above",
            # image_mso='WordOpenParaAbove',
            # screentip="Top paragraph spacing",
            # supertip="Change the paragraph spacing before the paragraph to the specified value (in pt).",
            #attr='SpaceBefore'
        ),
        ParSpaceAfter(
            id = 'par_sep_bottom',
            show_label=False,
            size_string = '##',
            # label=u"Paragraph spacing below",
            # image_mso='WordOpenParaBelow',
            # screentip="Bottom paragraph spacing",
            # supertip="Change the paragraph spacing after the paragraph to the specified value (in pt).",
            #attr='SpaceAfter'
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='PowerPointParagraphDialog')
    ]
)

paragraph_indent_group = bkt.ribbon.Group(
    id="bkt_paragraph_adv_group",
    label = "Paragraph indent",
    image_mso='ViewRulerPowerPoint',
    #ViewRulerPowerPoint
    children = [
        ParFirstLineIndent(
            id = 'first_line_indent',
            show_label=False,
            # label=u"Indent first line",
            # image='first_line_indent',
            # screentip="Special indent",
            # supertip="Change the special indent (first line, hanging) to the specified value (in cm).",
            # attr='FirstLineIndent',
            # big_step = 0.25,
            # small_step = 0.125,
            # rounding_factor = 0.125,
            # size_string = '-###',
        ),
        ParLeftIndent(
            id = 'left_indent',
            show_label=False,
            # label=u"Indent left",
            # image_mso='IndentClassic',
            # screentip="Paragraph indent",
            # supertip="Change the paragraph indent to the specified value (in cm).",
            # attr='LeftIndent',
            # big_step = 0.25,
            # small_step = 0.125,
            # rounding_factor = 0.125,
            # size_string = '-###',
        ),
        ParRightIndent(
            id = 'right_indent',
            show_label=False,
            # label=u"Indent left",
            # image_mso='IndentClassic',
            # screentip="Paragraph indent",
            # supertip="Change the paragraph indent to the specified value (in cm).",
            # attr='LeftIndent',
            # big_step = 0.25,
            # small_step = 0.125,
            # rounding_factor = 0.125,
            # size_string = '-###',
        ),
        ParSpaceWithin(
            id = 'par_sep_within',
            show_label=False,
            # label=u"Row spacing",
            # image_mso='LineSpacing',
            # screentip="Row spacing",
            # supertip="Change the line spacing (either in number of lines or in pt).",
            # attr='SpaceWithin',
            # size_string = '-###',
            # fallback_value = 1,
        ),
        bkt.ribbon.CheckBox(
            id = 'wordwrap2',
            label="WordWrap",
            # image_mso="FormattingMarkDropDown",
            screentip="Wrap text in shape",
            supertip="Configure the text option to 'Wrap text in shape'.",
            on_toggle_action=bkt.Callback(Absatz.set_word_wrap , shapes=True, require_text=True),
            get_pressed=bkt.Callback(Absatz.get_word_wrap , shapes=True, require_text=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.CheckBox(
            id = 'autosize2',
            label="AutoSize",
            # image_mso="SmartArtLargerShape",
            screentip="Adjust the size of the shape",
            supertip="Configure the text option to 'Resize shape to fit text' or 'Do not autofit'.",
            on_toggle_action=bkt.Callback(Absatz.set_auto_size , shapes=True, require_text=True),
            get_pressed=bkt.Callback(Absatz.get_auto_size , shapes=True, require_text=True),
            get_enabled = bkt.get_enabled_auto,
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='PowerPointParagraphDialog')
    ]
)


compact_font_group = bkt.ribbon.Group(
    id="bkt_compact_font_group",
    label = "Font",
    image_mso='GroupFont',
    children = [
        #NOTE: horizontal box layout leads to spacing between Font and FontSize ComboBox!
        bkt.mso.comboBox.Font(sizeString="WWWWWWWI", show_label=False),
        bkt.ribbon.ButtonGroup(children=[
            bkt.mso.control.Bold,
            bkt.mso.control.Italic,
            bkt.mso.control.Underline,
            # bkt.mso.control.Shadow,
            bkt.mso.control.Strikethrough,
        ]),
        bkt.ribbon.ButtonGroup(children=[
            bkt.mso.control.CharacterSpacingGallery,
            bkt.mso.control.ChangeCaseGallery,
            bkt.mso.control.ClearFormatting,
        ]),

        bkt.mso.control.FontSize,
        bkt.ribbon.ButtonGroup(children=[
            bkt.mso.control.FontSizeIncrease,
            bkt.mso.control.FontSizeDecrease,
        ]),
        bkt.ribbon.ButtonGroup(children=[
            bkt.mso.control.Superscript,
            bkt.mso.control.Subscript,
        ]),
        bkt.ribbon.DialogBoxLauncher(idMso='FontDialogPowerPoint')
    ]
)

compact_paragraph_group = bkt.ribbon.Group(
    id="bkt_compact_paragraph_group",
    label = "Paragraph",
    image_mso='GroupParagraph',
    children = [
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.ButtonGroup(children=[
                bkt.mso.control.BulletsGallery,
                bkt.mso.control.NumberingGallery,
            ]),
            bkt.ribbon.ButtonGroup(children=[
                bkt.mso.control.IndentDecrease,
                bkt.mso.control.IndentIncrease,
            ]),
            # bkt.mso.control.ConvertToSmartArt,
        ]),
        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.ribbon.ButtonGroup(children=[
                bkt.mso.control.AlignLeft,
                bkt.mso.control.AlignCenter,
                bkt.mso.control.AlignRight,
                bkt.mso.control.AlignJustify,
                bkt.mso.control.AlignJustifyMenu,
                bkt.mso.control.TableColumnsGallery,
            ]),
            # bkt.mso.control.ParagraphDistributed,
            # bkt.mso.control.AlignJustifyThai,
            # bkt.mso.control.TextDirectionLeftToRight,
            # bkt.mso.control.TextDirectionRightToLeft,
        ]),

        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.mso.control.LineSpacingGalleryPowerPoint,
            # bkt.mso.control.FontColorPicker,
            bkt.mso.control.TextDirectionGallery,
            bkt.mso.control.TextAlignGallery,
        ]),
        bkt.ribbon.DialogBoxLauncher(idMso='PowerPointParagraphDialog')
    ]
)