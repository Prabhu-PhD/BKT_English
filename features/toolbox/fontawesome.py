# -*- coding: utf-8 -*-
'''
Created on 10.02.2017

@author: rdebeerst
'''



import importlib
from collections import namedtuple

import bkt
from bkt.library.powerpoint import PPTSymbolsGallery


FontSymbol = namedtuple("FontSymbol", "module fontlabel fontname unicode label keywords")

class Fontawesome(object):
    installed_fonts = None
    fontsettings = [
            # module-name,      font-filename,                  suppress-font-not-installed-message, label
            ('fabricmdl2',      'Fabric MDL2 Assets',           True,  'Fabric MDL2'),
            ('fontawesome4',    'FontAwesome',                  True,  'Font Awesome 4'),
            ('fontawesome5',    'Font Awesome 5 Free Regular',  True, 'Font Awesome 5'),
            ('fontawesome6',    'Font Awesome 6 Free Regular',  True, 'Font Awesome 6'),
            ('fontawesome7',    'Font Awesome 7 Free Regular',  False, 'Font Awesome 7'),
            ('icomoon',         'IcoMoon-Free',                 True, 'IcoMoon'),
            ('materialicons',   'Material Icons',               True,  'Material Icons'),
            ('materialsymbols', 'Material Symbols Sharp',       False, 'Material Symbols'),
            ('segoemdl2',       'Segoe MDL2 Assets',            True,  'Segoe MDL2'),
            ('segoefluent',     'Segoe Fluent Icons',           True,  'Segoe Fluent'),
            ('segoeui',         'Segoe UI',                     False, 'Segoe UI'),
            ('wingdings',       'Wingdings',                    True,  'Wingdings'),
            ('unicodes',        'Segoe UI',                     True,  'Unicode Symbols'),
            # ('foobar', 'Non-existing test font', True, 'Foobar Test Font'),
        ]
    search_engine = None
    searchable_fonts = []
    exclusion = bkt.settings.get("toolbox.fonts_excluded", ["fontawesome4", "segoemdl2"])

    @classmethod
    def get_installed_fonts(cls):
        if not cls.installed_fonts:
            # Method 1 (returns Font Awesome 5 Free Regular)
            import System.Drawing.Text
            font_collection = System.Drawing.Text.InstalledFontCollection()
            cls.installed_fonts = [font.Name for font in font_collection.Families]
            # Method 2 (return Font Awesome 5 Free)
            # import System.Windows.Media
            # font_families = System.Windows.Media.Fonts.SystemFontFamilies
            # cls.installed_fonts = [font.Source for font in font_families]
        return cls.installed_fonts
    
    # helper to check system-fonts
    @classmethod
    def font_exists(cls, fontname):
        return fontname in cls.get_installed_fonts()

    @classmethod
    def get_symbol_galleries(cls):
        symbol_galleries = []
        for font_module, font_name, suppress_hint, font_label in cls.fontsettings:
            # check if font exists and is not excluded
            if font_module in cls.exclusion:
                continue
            elif cls.font_exists(font_name):
                # import the corresponding font-symbol-module from 'fontsymbols'-folder
                fontsymbolmodule = importlib.import_module('toolbox.fontsymbols.%s' % font_module)
                
                if not hasattr(fontsymbolmodule, 'menus'):
                    continue

                # add menu seperator with title
                if fontsymbolmodule.menu_title:
                    symbol_galleries += [
                        bkt.ribbon.MenuSeparator(title="" + fontsymbolmodule.menu_title),
                    ]
                else:
                    symbol_galleries += [
                        bkt.ribbon.MenuSeparator(title=font_label),
                    ]
                
                # add font-symbol-galleries
                symbol_galleries += fontsymbolmodule.menus
            elif not suppress_hint:
                symbol_galleries += [
                    bkt.ribbon.MenuSeparator(title=font_label),
                    bkt.ribbon.Button(
                        label="Font not installed",
                        enabled=False
                    )
                ]
        return symbol_galleries
    
    @classmethod
    def clear_search_engine(cls):
        cls.search_engine = None
        cls.searchable_fonts = []

    @classmethod
    def get_search_engine(cls, context):
        if cls.search_engine:
            return cls.search_engine

        from bkt.library.search import get_search_engine
        cls.search_engine = get_search_engine("fonticons", FontSymbol)
        # initialize search index on first use of engine
        # cls.update_search_index(cls.search_engine)
        
        def loop(worker):
            worker.ReportProgress(1, "Lege Suchindex an...")
            try:
                cls.update_search_index(cls.search_engine)
            except Exception as e:
                bkt.message.error("Fehler beim erstellen des Suchindex: {}".format(e), "BKT: Font-Icons")

        bkt.ui.execute_with_progress_bar(loop, context, indeterminate=True)
        return cls.search_engine
    
    @classmethod
    def update_search_index(cls, engine=None):
        engine = engine or cls.search_engine
        for font_module, font_name, _, font_label in cls.fontsettings:
            # check if font exists and is not excluded
            if font_module not in cls.exclusion and cls.font_exists(font_name):
                # import the corresponding font-symbol-module from 'fontsymbols'-folder
                fontsymbolmodule = importlib.import_module('toolbox.fontsymbols.%s' % font_module)
                try:
                    fontsymbolmodule.update_search_index(engine)
                    # cls.searchable_fonts.append(fontsymbolmodule.menu_title)
                    cls.searchable_fonts.append(font_label)
                except AttributeError:
                    continue

    @classmethod
    def get_text_fontawesome(cls):
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=cls.get_symbol_galleries()
            )
    
    @classmethod
    def toggle_exclusion(cls, current_control, pressed):
        module = current_control["tag"]
        if module in cls.exclusion:
            cls.exclusion.remove(module)
        else:
            cls.exclusion.append(module)
        bkt.settings["toolbox.fonts_excluded"] = cls.exclusion
    
    @classmethod
    def pressed_exclusion(cls, current_control):
        return current_control["tag"] in cls.exclusion

    @classmethod
    def get_exclusions(cls):
        def _toggle_button(font_module, font_label):
            return bkt.ribbon.ToggleButton(
                    label=font_label,
                    # screentip="Unicode font matches theme font",
                    # supertip="No special Unicode font is used, but the theme's default font.",
                    tag=font_module,
                    on_toggle_action=bkt.Callback(cls.toggle_exclusion),
                    get_pressed=bkt.Callback(cls.pressed_exclusion),
                )
        
        return [
                _toggle_button(font_module, font_label)
                for font_module, _, _, font_label in cls.fontsettings
            ]

    @classmethod
    def get_exclusion_menu(cls):
        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=cls.get_exclusions()
            )


class FontSearch(object):
    search_term = ""
    search_results = None
    search_exact = bkt.settings.get("bkt.symbols.search_exact", True)
    search_and = True #True = AND-search, False=OR-search

    search_fallback_font = bkt.settings.get("bkt.symbols.search_fallback_font", "Segoe UI Emoji")

    _cache_menu_infos = None

    @classmethod
    def _perform_search(cls, context):
        if len(cls.search_term) > 0:
            if len(cls.search_term) < 3:
                cls.search_exact = True
            cls.search_results = None
            engine = cls.get_search_engine(context)
            with engine.searcher() as searcher:
                if cls.search_exact:
                    cls.search_results = searcher.search_exact(cls.search_term, cls.search_and)
                else:
                    cls.search_results = searcher.search(cls.search_term, cls.search_and)
        else:
            cls.search_results = None

    @classmethod
    def set_rendering_font(cls, font_name):
        cls._cache_menu_infos = None
        # Fontawesome.clear_search_engine()
        cls.search_fallback_font = font_name
        bkt.settings["bkt.symbols.search_fallback_font"] = font_name

    @classmethod
    def toggle_search_exact(cls, pressed, context):
        cls.search_exact = not cls.search_exact
        bkt.settings["bkt.symbols.search_exact"] = cls.search_exact
        cls._perform_search(context)

    @classmethod
    def checked_search_exact(cls):
        return cls.search_exact

    @classmethod
    def set_search_term(cls, value, context):
        cls.search_term = value
        cls._perform_search(context)

    @classmethod
    def get_search_term(cls):
        return cls.search_term

    @classmethod
    def get_search_engine(cls, context):
        return Fontawesome.get_search_engine(context)
    
    @classmethod
    def get_symbol_galleries(cls, context):
        if not cls.search_results or len(cls.search_results) == 0:
            fontmodules = [
                bkt.ribbon.Button(
                    label="No results for »{}«".format(cls.search_term),
                    enabled=False
                )
            ]
        
        else:
            fontmodules = [
                bkt.ribbon.MenuSeparator(title="{} results for »{}«".format(len(cls.search_results), cls.search_term))
            ]
            for fontlabel, icons in cls.search_results.groupedby("fontlabel"):
                len_icons = len(icons)
                if len_icons > 999:
                    icons = icons[:999]
                    label = "{} (999 of {})".format(fontlabel, len_icons)
                else:
                    label = f"{fontlabel} ({len_icons})"
                
                fontmodules.append(
                    PPTSymbolsGallery(
                        label=label,
                        symbols=[
                            (
                                ico.fontname or cls.search_fallback_font,
                                ico.unicode,
                                # unichr(int(ico.unicode, 16)),
                                ico.label,
                                ', '.join(sorted(ico.keywords))
                            ) for ico in icons
                        ],
                        columns=16
                    )
                )

        fontmodules.extend(cls._get_symbol_galleries_infos(context))
        
        return bkt.ribbon.Menu(
                    xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                    id=None,
                    children=fontmodules
                )
    
    @classmethod
    def _get_symbol_galleries_infos(cls, context):
        if cls._cache_menu_infos:
            return cls._cache_menu_infos

        engine = cls.get_search_engine(context)
        cls._cache_menu_infos = [
            bkt.ribbon.MenuSeparator(title="Information"),
            bkt.ribbon.Button(
                label="Indexed icons: {}".format(engine.count_documents()),
                enabled=False,
            ),
            bkt.ribbon.Button(
                label="Indexed keywords: {}".format(engine.count_keywords()),
                enabled=False,
            ),
            bkt.ribbon.Button(
                label="Searchable fonts: {}".format(len(Fontawesome.searchable_fonts)),
                enabled=False,
                supertip=", ".join(Fontawesome.searchable_fonts)
            ),
        ]
        return cls._cache_menu_infos
    
    @classmethod
    def get_enabled_results(cls):
        return cls.search_results is not None
    
    @classmethod
    def get_results_label(cls):
        if cls.search_results is not None:
            return "{} Icons".format(len(cls.search_results))
        else:
            return "Ergebnis"
    
    @classmethod
    def get_unicode_settings_menu(cls):
        def _unicode_font_button(font):
            return bkt.ribbon.ToggleButton(
                label=font,
                screentip="Unicode font"+font,
                supertip=font+" als Unicode-Schriftart verwenden.",
                on_toggle_action=bkt.Callback(lambda pressed: cls.set_rendering_font(font)),
                get_pressed=bkt.Callback(lambda: cls.search_fallback_font == font),
                get_image=bkt.Callback(lambda:bkt.ribbon.SymbolsGallery.create_symbol_image(font, "\u2194"))
            )

        return bkt.ribbon.Menu(
                xmlns="http://schemas.microsoft.com/office/2009/07/customui",
                id=None,
                children=[
                            _unicode_font_button(font)
                            for font in bkt.library.powerpoint.PPTSymbolsSettings.UNICODE_FONTS
                        ]
            )


# Font search
fontsearch_gruppe = bkt.ribbon.Group(
    id="bkt_fontsearch_group",
    label="Icon search",
    image_mso='GroupSearch',
    children=[
        # bkt.ribbon.DynamicMenu(
        #     id="fontsearch_all_symbols",
        #     label="All icons",
        #     size="large",
        #     supertip="Shows icons for available icon fonts that can be inserted as a text symbol or graphic.\n\nNote: The icon fonts must be installed on the computer.",
        #     image_mso="Call",
        #     get_content = bkt.Callback(Fontawesome.get_text_fontawesome)
        # ),
        # bkt.ribbon.Separator(),
        bkt.ribbon.Label(
            label="Search term:",
        ),
        # bkt.ribbon.EditBox(
        bkt.ribbon.ComboBox(
            label="Search term",
            show_label=False,
            sizeString = '######',
            get_text = bkt.Callback(FontSearch.get_search_term),
            on_change = bkt.Callback(FontSearch.set_search_term, context=True),
            supertip="Enter search term and press ENTER",
            get_item_count=bkt.Callback(lambda context: FontSearch.get_search_engine(context).count_recent_searches(), context=True),
            get_item_label=bkt.Callback(lambda index, context: FontSearch.get_search_engine(context).get_recent_searches()[index], context=True),
        ),
        bkt.ribbon.Menu(
            label="Config",
            screentip="Settings",
            supertip="Icon search settings",
            children=[
                bkt.ribbon.ToggleButton(
                    label="Toggle exact search",
                    supertip="When exact search is disabled, 'person' also finds 'personality', 'impersonal', etc.",
                    on_toggle_action=bkt.Callback(FontSearch.toggle_search_exact, context=True),
                    get_pressed=bkt.Callback(FontSearch.checked_search_exact),
                ),
                bkt.ribbon.DynamicMenu(
                    label="Exclude icon fonts",
                    supertip="Do not show certain icon fonts and exclude them from the search.",
                    get_content = bkt.Callback(Fontawesome.get_exclusion_menu),
                ),
                bkt.ribbon.DynamicMenu(
                    label="Font for Unicode symbols",
                    supertip="Select the font for rendering the Unicode symbols. The default is Segoe Emoji.",
                    # get_content = bkt.CallbackLazy("toolbox.fontsymbols.unicodes", "get_settings_menu"),
                    get_content = bkt.Callback(FontSearch.get_unicode_settings_menu),
                ),
            ]
        ),
        bkt.ribbon.Separator(),
        bkt.ribbon.DynamicMenu(
            get_label=bkt.Callback(FontSearch.get_results_label),
            get_content=bkt.Callback(FontSearch.get_symbol_galleries, context=True),
            get_enabled=bkt.Callback(FontSearch.get_enabled_results),
            screentip="Search results",
            size="large",
            image_mso="Call",
            supertip="Shows the search results of the icon search for the desired search term",
        ),
        # bkt.ribbon.Box(children=[
        #     bkt.ribbon.Button(
        #         label="sync",
        #         on_action=bkt.Callback(FontSearch.sync_cache),
        #     ),
        #     bkt.ribbon.Button(
        #         label="clear",
        #         on_action=bkt.Callback(FontSearch.clear_cache),
        #     ),
        # ]),
    ]
)