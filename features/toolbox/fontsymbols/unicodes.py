# -*- coding: utf-8 -*-

import unicodedata
import logging

def update_search_index(search_engine):
    search_writer = search_engine.writer()

    # Index only the genuinely useful "symbol" code points -- NOT the entire
    # Unicode plane. The original code iterated range(0x30000) (~196k code
    # points, incl. all CJK/Hangul/historic scripts), producing a ~150k-entry
    # search index. That bloat caused the icon search to crash PowerPoint with a
    # native stack overflow (exception 0xc00000fd). These blocks cover the
    # symbols people actually want: Latin/Greek/Cyrillic letters, punctuation,
    # currency, arrows, math, technical, geometric shapes, misc symbols,
    # dingbats, and emoji/pictographs.
    symbol_ranges = (
        (0x0000, 0x3000),    # incl. arrows, math, symbols, dingbats, etc.
        (0x1F000, 0x1FB00),  # emoji & pictographs
    )

    # Generate a list of useful Unicode character names
    unicode_characters = []
    for start, end in symbol_ranges:
        for codepoint in range(start, end):
            try:
                char_name = unicodedata.name(chr(codepoint))
                unicode_characters.append((char_name, codepoint))
            except ValueError:
                # Skip characters without a name
                continue
    
    logging.info(f"Found {len(unicode_characters)} Unicode characters.")

    # unicode_fonts = [
    #     ("Segoe UI", 0x0000, 0x0780)
    #     ("Segoe UI Emoji", 0x1f600, 0x1f650)
    # ]

    # def _get_font(code):
    #     for font, start, end in unicode_fonts:
    #         if start <= code < end:
    #             return font
    #     return "Segoe UI Symbol"

    for label, code in unicode_characters:
        search_writer.add_document(
            module="unicodes",
            fontlabel="Unicode Symbols",
            # fontname=UnicodeSymbols.rendering_font,
            fontname=None,
            unicode=chr(code),
            label=f"{label} (U+{code:04X})",
            keywords=label.lower().split()
        )

    search_writer.commit()
