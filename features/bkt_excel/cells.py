# -*- coding: utf-8 -*-
'''
Created on 2017-07-18
@author: Florian Stallmann
'''

import logging

# import System.Convert as Converter #required for convert to double
# import System.DateTime as DateTime #required for parse as date and time
from System import DateTime, Array

import bkt
import bkt.library.excel.helpers as xllib
import bkt.library.excel.constants as xlcon

import bkt.dotnet as dotnet
Forms = dotnet.import_forms() #required to copy text to clipboard

class CellsOps(object):
    # hidden_columns = {}
    # hidden_rows = {}

    last_formula = "*100"
    last_prepend = "ID-"
    last_append = "...!?"
    last_slice_pos = "2:"
    last_slice_text = "/"

    #regex match count
    last_regex_match_pattern = r"([A-Z])\w+"

    #regex split
    last_regex_split_pattern = r"[;,\. ]+"
    last_regex_split_mode = 0

    #regex replace
    last_regex_sub_pattern = r"\sand\s"
    last_regex_sub_repl = r" & "

    @staticmethod
    def _set_hidden_name(key, sheet, rng):
        try:
            #sheet.Names(key).RefersToLocal = "=" + address
            sheet.Names(key).RefersTo = rng
        except:
            #sheet.Names.Add(Name=key, RefersToLocal="=" + address, Visible=False)
            sheet.Names.Add(Name=key, RefersTo=rng, Visible=False)

    @staticmethod
    def _get_hidden_name(key, sheet, delete=True):
        try:
            # rng = sheet.Names(key).RefersToRange -> Cuts off too long range strings, below method is better
            addr = sheet.Names(key).RefersToLocal[1:]
            pos = addr.find("!")+1
            addr = addr.replace(addr[0:pos], "") #remove sheet names to support much longer range strings
            if delete:
                sheet.Names(key).Delete()
            return sheet.Range(addr)
        except:
            return None

    # @staticmethod
    # def _del_hidden_name(key, sheet):
    #     try:
    #         name = sheet.Names(key).Delete()
    #     except:
    #         pass

    @classmethod
    def prepend_text(cls, cells, application):
        input_text = bkt.ui.show_user_input("Text eingeben, der vor alle Zellen geschrieben werden soll. Existierende Formeln werden überschrieben und durch Werte ersetzt.\n\nMögliche Platzhalter: [counter], [row], [column].", "Text voranstellen", cls.last_prepend)
        if not input_text:
            return

        if not xllib.confirm_no_undo(): return

        cls.last_prepend = input_text
        number_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])

        counter = 1
        for cell in cells:
            input_text_local = input_text.replace("[counter]", str(counter)).replace("[row]", str(cell.Row)).replace("[column]", str(cell.Column))
            cell.Value = input_text_local + cell.Text
            cell.NumberFormatLocal = number_format
            counter += 1

    @classmethod
    def append_text(cls, cells, application):
        input_text = bkt.ui.show_user_input("Text eingeben, der hinter alle Zellen geschrieben werden soll. Existierende Formeln werden überschrieben und durch Werte ersetzt.\n\nMögliche Platzhalter: [counter], [row], [column].", "Text anhängen", cls.last_append)
        if not input_text:
            return

        if not xllib.confirm_no_undo(): return

        cls.last_append = input_text
        number_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])

        counter = 1
        for cell in cells:
            input_text_local = input_text.replace("[counter]", str(counter)).replace("[row]", str(cell.Row)).replace("[column]", str(cell.Column))
            cell.Value = cell.Text + input_text_local
            cell.NumberFormatLocal = number_format
            counter += 1

    @classmethod
    def slice_text(cls, cells, application):
        def _get_slicer(pos_text):
            input_params = pos_text.strip(' \t\n\r[]').split(":")
            
            #Extract single character
            if len(input_params) == 1:
                start = int(input_params[0])
                stop = None if start < 0 else start + 1
            
            #Extract string with start and stop
            elif len(input_params) == 2:
                start = 0 if not input_params[0] else int(input_params[0])
                stop = None if not input_params[1] else int(input_params[1])

            else:
                raise ValueError('invalid number of parameters')
            
            if stop is not None and ((start > 0 and stop > 0) or (start < 0 and stop < 0)) and start >= stop:
                raise ValueError('no text remains as start is after stop')

            return slice(start, stop)

        preview_cell = application.ActiveCell
        def _preview(sender, e):
            try:
                if text.Text == '':
                    txt_preview.Text = ''
                else:
                    s = _get_slicer(text.Text)
                    txt_preview.Text = preview_cell.Text[s]
            except:
                txt_preview.Text = "FEHLER"

        explanation = '''Start- und Stopp-Position zum Schneiden getrennt mit ":" eingeben. Ist keine Start-Position gegeben, wird diese auf 0 gesetzt. Ist keine Stopp-Position gegeben, wird diese bis Textende gesetzt. Eine negative Position bedeutet, dass diese vom Textende berechnet wird.

  Beispiel für "ABCDEF":
  [2:]   = CDEF  Entferne die beiden ersten Zeichen
  [:2]   = AB    Entferne alles nach dem zweiten Zeichen
  [-2:]  = EF    Entferne alles bis zum vorletzten Zeichen
  [2:-2] = CD    Entferne 2 Zeichen an Anfang und Ende'''

        user_form = bkt.ui.UserInputBox(explanation, "Text anhand Position schneiden")
        text = user_form._add_textbox("text", cls.last_slice_pos)
        text.TextChanged += _preview
        
        user_form._add_label("Vorschau für aktive Zelle:")
        txt_preview = user_form._add_textbox("preview")
        txt_preview.ReadOnly = True
        _preview(None, None)

        form_return = user_form.show()
        if len(form_return) == 0 or form_return["text"] == '':
            return
        
        if not xllib.confirm_no_undo(): return

        number_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        cls.last_slice_pos = form_return["text"]

        try:
            s = _get_slicer(form_return["text"])
        except:
            bkt.message("Ungültige Eingabe!")
            return

        for cell in cells:
            cell.Value = cell.Text[s]
            cell.NumberFormatLocal = number_format

    @classmethod
    def find_and_slice_text(cls, cells, application):
        def _slice_text(initial_text, search_text, find_method, rslice):
            pos = find_method(initial_text, search_text)
            if pos == -1:
                return initial_text
            start = pos+len(search_text) if rslice else 0
            stop = None if rslice else pos
            s = slice(start, stop)
            return initial_text[s]

        preview_cell = application.ActiveCell
        def _preview(sender, e):
            try:
                find_method = str.rfind if cb_rfind.Checked else str.find
                txt_preview.Text = _slice_text(preview_cell.Text, text.Text, find_method, cb_rslice.Checked)
            except:
                txt_preview.Text = "FEHLER"

        user_form = bkt.ui.UserInputBox("Gibt Zelleninhalt von Beginn bis zum eingegebenen Text zurück. Wird der Text nicht gefunden, bleibt der Zelleninhalt unverändert.", "Text anhand Zeichen schneiden")
        text = user_form._add_textbox("text", cls.last_slice_text)
        text.TextChanged += _preview
        cb_rslice = user_form._add_checkbox("rslice", "Rechten Teil zurückgeben (ab eingegebenem Text bis Ende)")
        cb_rslice.CheckedChanged += _preview
        cb_rfind = user_form._add_checkbox("rfind", "Von rechts anfangen zu suchen")
        cb_rfind.CheckedChanged += _preview
        
        user_form._add_label("Vorschau für aktive Zelle:")
        txt_preview = user_form._add_textbox("preview")
        txt_preview.ReadOnly = True
        _preview(None, None)

        form_return = user_form.show()
        if len(form_return) == 0 or form_return["text"] == '':
            return

        if not xllib.confirm_no_undo(): return

        find_method = str.rfind if form_return["rfind"] else str.find
        
        number_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        cls.last_slice_text = form_return["text"]

        for cell in cells:
            try:
                cell.Value = _slice_text(cell.Text, form_return["text"], find_method, form_return["rslice"])
                cell.NumberFormatLocal = number_format
            except:
                pass


    @classmethod
    def apply_formula(cls, cells, application):
        from string import Template

        dec_sep = application.International(xlcon.XlApplicationInternational["xlDecimalSeparator"])
        active_cell = application.ActiveCell
        preview_cell_format = active_cell.NumberFormatLocal
        preview_cell_formula = active_cell.FormulaLocal.lstrip("=")
        preview_cell_address = active_cell.AddressLocal(False, False)
        try:
            preview_cell_value = str(active_cell.Value2).replace('.', dec_sep)
        except:
            preview_cell_value = -2146826273 #"#Value!"

        def _preview(sender, e):
            try:
                create_formulas = text.Text[0] == "="
                formula = text.Text if create_formulas or "$cell" in text.Text else "$cell" + text.Text
                template = Template(formula)
                formula = template.safe_substitute({"cell": preview_cell_address, "cellvalue": preview_cell_value})

                if create_formulas:
                    formula = template.safe_substitute({"cell": preview_cell_formula, "cellvalue": preview_cell_value})
                    # formula = formula.replace("[cell]", preview_cell_formula)
                    formula = formula[1:]
                else:
                    formula = template.safe_substitute({"cell": preview_cell_address, "cellvalue": preview_cell_value})
                    # formula = formula.replace("[cell]", preview_cell_value)
                    # formula = formula.replace("[cell]", preview_cell_address)
                    # formula = template.safe_substitute({"cell": preview_cell_address, "cellvalue": preview_cell_value})

                txt_preview.Text = "=" + formula
                txt_preview2.Text = str(xllib.xls_evaluate(formula, dec_sep, preview_cell_format))

            except:
                txt_preview.Text = "FEHLER"
                txt_preview2.Text = "FEHLER"

        user_form = bkt.ui.UserInputBox("Hier kann eine Formel auf alle markierten Zellen angewendet werden. Soll der Zelleninhalt nicht am Anfang stehen, können Sie mit dem Platzhalter $cell und $cellvalue arbeiten. Standardmäßig wird der resultierende Wert eingefügt (sofern die Formel nicht fehlerhaft ist). Wenn Ihre Eingabe mit '=' beginnt, wird eine Formel erstellt. In der Auswahlbox finden Sie Beispiele für mögliche Eingaben.", "Formel anwenden")
        text = user_form._add_combobox("formula", cls.last_formula, ["*100", "/100", "*(-1)", "+A1", "/SUMME(A1:A3)", "ABS($cell)", "RUNDEN($cell;2)", "ABRUNDEN($cell;2)", "AUFRUNDEN($cell;2)", "KÜRZEN($cell)", "1/$cell", "=($cell)*100", "GROSS(\"$cellvalue\")"])
        text.TextChanged += _preview
        user_form._add_checkbox("skip_existing_formulas", "Bestehende Formeln überspringen und nicht verändern")
        
        user_form._add_label("Vorschau für aktive Zelle:")
        txt_preview = user_form._add_textbox("preview")
        txt_preview.ReadOnly = True

        txt_preview2 = user_form._add_textbox("preview2")
        txt_preview2.ReadOnly = True

        _preview(None, None)

        form_return = user_form.show()
        if len(form_return) == 0:
            return

        if not xllib.confirm_no_undo(): return

        err_counter = 0
        formula = form_return["formula"]
        cls.last_formula = formula

        create_formulas = formula[0] == "="
        formula = formula if create_formulas or "$cell" in formula else "($cell)" + formula
        template = Template(formula)

        for cell in cells:
            if cell.Value2 is None or (cell.HasFormula and form_return["skip_existing_formulas"]):
                continue

            try:
                cell_value = str(cell.Value2).replace('.', dec_sep)
            except:
                cell_value = -2146826273 #"#Value!"

            # if cell.HasFormula and create_formulas:
            #     new_formula = formula.replace("[cell]", cell.FormulaLocal[1:])
            # else:
            #     new_formula = formula.replace("[cell]", cell.FormulaLocal)

            try:
                if create_formulas:
                    if cell.HasFormula:
                        cell.FormulaLocal = template.safe_substitute({"cell": cell.FormulaLocal[1:], "cellvalue": cell_value})
                    else:
                        cell.FormulaLocal = template.safe_substitute({"cell": cell.FormulaLocal, "cellvalue": cell_value})
                else:
                    new_formula = template.safe_substitute({"cell": cell.AddressLocal(False, False), "cellvalue": cell_value})
                    cell.FormulaLocal = application.Evaluate(new_formula)
                    # cell.FormulaLocal = "=" + new_formula
                    # #On error do not replace with int value of error
                    # if not application.WorksheetFunction.IsError(cell):
                    #     cell.Value = cell.Value()
            except:
                err_counter += 1
                cell.AddComment("Error applying formula")
                #bkt.helpers.exception_as_message(str(cell.AddressLocal()))

        if err_counter > 0:
            bkt.message("Fehler! Formel war auf " + str(err_counter) + " Zelle(n) nicht anwendbar.")


    @classmethod
    def regex_match_count(cls, cells, application):
        import re
        from itertools import chain
        
        preview_cell_value = application.ActiveCell.Text

        def _get_flags():
            flags=0
            if ignorecase.Checked:
                flags = re.IGNORECASE
            return flags

        def _do_regex(regex, string):
            if regex.groups > 0:
                return sum(res is not None for res in bkt.helpers.flatten(m.groups() for m in regex.finditer(string)))
            else:
                # return len(list(m.group() for m in regex.finditer(string)))
                return len(list(regex.finditer(string)))

        def _preview(sender, e):
            try:
                regex = re.compile(text.Text, _get_flags())
                txt_preview.Text = str(_do_regex(regex, preview_cell_value))
            except Exception as e:
                txt_preview.Text = "FEHLER: " + str(e)

        user_form = bkt.ui.UserInputBox("Hier kann ein regulärer Ausdruck in allen markierten Zellen gesucht und die Anzahl der Funde gezählt werden. In der Auswahlbox finden Sie Beispiele für mögliche Eingaben.", "RegEx anwenden")
        text = user_form._add_combobox("regex", cls.last_regex_match_pattern, [r"[;,\. ]+", r"([A-Z])\w+", r"([+-]?[\d\.]+,*[0-9]*)", r"[\w\.-]+@[\w\.-]+\.\w{2,4}"])
        text.TextChanged += _preview

        ignorecase = user_form._add_checkbox("ignorecase", "Groß-/Kleinschreibung ignorieren")
        ignorecase.CheckedChanged += _preview
        
        user_form._add_label("Vorschau für aktive Zelle:")
        txt_preview = user_form._add_textbox("preview")
        txt_preview.ReadOnly = True
        _preview(None, None)

        form_return = user_form.show()
        if len(form_return) == 0:
            return
        
        try:
            regex = re.compile(form_return["regex"], _get_flags())
        except Exception as e:
            bkt.message("Fehler! RegEx kann nicht kompiliert werden: "+str(e))
            return

        if not xllib.confirm_no_undo(): return

        err_counter = 0
        cls.last_regex_match_pattern = form_return["regex"]

        for cell in cells:
            if cell.Value2 is None:
                continue

            try:
                cell.Value = _do_regex(regex, cell.Text)
            except:
                err_counter += 1
                #bkt.helpers.exception_as_message(str(cell.AddressLocal()))

        if err_counter > 0:
            bkt.message("Fehler! RegEx war auf " + str(err_counter) + " Zelle(n) nicht anwendbar.")

    @classmethod
    def regex_split_to_columns(cls, cells, application):
        import re
        from itertools import chain
        
        preview_cell_value = application.ActiveCell.Text

        def _get_flags():
            flags=0
            if ignorecase.Checked:
                flags = re.IGNORECASE
            return flags

        def _get_mode():
            for radio in mode_radios:
                if radio.Checked:
                    return radio.Text

        def _do_regex(regex, string, join=None):
            current_mode = _get_mode()
            if current_mode.startswith("Split"):
                regex_result = regex.split(string)
            else:
                if regex.groups > 0:
                    regex_result = list(bkt.helpers.flatten(m.groups("") for m in regex.finditer(string)))
                else:
                    regex_result = list(m.group() for m in regex.finditer(string))
            
            if join is None:
                return regex_result
            else:
                return join.join(res or "" for res in regex_result)

        def _preview(sender, e):
            try:
                regex = re.compile(text.Text, _get_flags())
                txt_preview.Text = _do_regex(regex, preview_cell_value, "\t")
            except Exception as e:
                txt_preview.Text = "FEHLER: " + str(e)

        user_form = bkt.ui.UserInputBox("Hier kann ein regulärer Ausdruck auf alle markierten Zellen angewendet werden. In der Auswahlbox finden Sie Beispiele für mögliche Eingaben.", "RegEx anwenden")
        text = user_form._add_combobox("regex", cls.last_regex_split_pattern, [r"[;,\. ]+", r"([A-Z])\w+", r"([+-]?[\d\.]+,*[0-9]*)", r"[\w\.-]+@[\w\.-]+\.\w{2,4}"])
        text.TextChanged += _preview

        ignorecase = user_form._add_checkbox("ignorecase", "Groß-/Kleinschreibung ignorieren")
        ignorecase.CheckedChanged += _preview

        radio_mode_values = ["Find: Aufteilung je RegEx Übereinstimmung", "Split: RegEx definiert Trennzeichen"]
        _, mode_radios = user_form._add_radio_buttons("mode", "Modus", radio_mode_values, cls.last_regex_split_mode)
        for radio in mode_radios:
            radio.CheckedChanged += _preview
        
        user_form._add_label("Vorschau für aktive Zelle (Gruppen mit Tabs getrennt):")
        txt_preview = user_form._add_textbox("preview")
        txt_preview.ReadOnly = True
        _preview(None, None)

        form_return = user_form.show()
        if len(form_return) == 0:
            return
        
        try:
            regex = re.compile(form_return["regex"], _get_flags())
        except Exception as e:
            bkt.message("Fehler! RegEx kann nicht kompiliert werden: "+str(e))
            return

        if not xllib.confirm_no_undo(): return

        err_counter = 0
        cls.last_regex_split_pattern = form_return["regex"]
        cls.last_regex_split_mode = radio_mode_values.index(form_return["mode"])

        for cell in cells:
            if cell.Value2 is None:
                continue

            try:
                values_split = _do_regex(regex, cell.Text)
                values = Array.CreateInstance(object, 1, len(values_split))
                for i,col in enumerate(values_split):
                    values[0,i] = col or ""
                new_area = xllib.resize_areas([cell], cols=len(values_split))[0]
                new_area.Value = values
            except:
                err_counter += 1
                # bkt.helpers.exception_as_message(str(cell.AddressLocal()))

        if err_counter > 0:
            bkt.message("Fehler! RegEx war auf " + str(err_counter) + " Zelle(n) nicht anwendbar.")

    @classmethod
    def regex_replace(cls, cells, application):
        import re
        
        preview_cell_value = application.ActiveCell.Text

        def _get_flags():
            flags=0
            if ignorecase.Checked:
                flags = re.IGNORECASE
            return flags

        def _preview(sender, e):
            try:
                txt_preview.Text = re.sub(pattern.Text, repl.Text, preview_cell_value, flags=_get_flags())
            except Exception as e:
                txt_preview.Text = "FEHLER: " + str(e)

        user_form = bkt.ui.UserInputBox("Hier kann ein regulärer Ausdruck in allen markierten Zellen gesucht und ersetzt werden. In der Auswahlbox finden Sie Beispiele für mögliche Eingaben.", "RegEx anwenden")
        pattern = user_form._add_combobox("pattern", cls.last_regex_sub_pattern, [r"\sand\s", r" {2,}", r"([A-Z]+)/([a-z]+)"])
        pattern.TextChanged += _preview

        repl = user_form._add_combobox("repl", cls.last_regex_sub_repl, [r" & ", r" ", r"\2/\1"])
        repl.TextChanged += _preview

        ignorecase = user_form._add_checkbox("ignorecase", "Groß-/Kleinschreibung ignorieren")
        ignorecase.CheckedChanged += _preview
        
        user_form._add_label("Vorschau für aktive Zelle:")
        txt_preview = user_form._add_textbox("preview")
        txt_preview.ReadOnly = True
        _preview(None, None)

        form_return = user_form.show()
        if len(form_return) == 0:
            return
        
        try:
            regex_pattern = re.compile(form_return["pattern"], _get_flags())
        except Exception as e:
            bkt.message("Fehler! RegEx kann nicht kompiliert werden: "+str(e))
            return

        if not xllib.confirm_no_undo(): return

        err_counter = 0
        cls.last_regex_sub_pattern = form_return["pattern"]
        cls.last_regex_sub_repl = form_return["repl"]

        for cell in cells:
            if cell.Value2 is None:
                continue

            try:
                cell.Value = regex_pattern.sub(form_return["repl"], cell.Text)
            except:
                err_counter += 1
                #bkt.helpers.exception_as_message(str(cell.AddressLocal()))

        if err_counter > 0:
            bkt.message("Fehler! RegEx war auf " + str(err_counter) + " Zelle(n) nicht anwendbar.")


    @staticmethod
    def merge_cells(selection, cells, join="\r\n"):
        if not xllib.confirm_no_undo(): return
        target_content = join.join([cell.Text for cell in cells])
        selection.ClearContents()
        selection.Cells[1].Value = target_content

        # target_cell = next(cells)
        # for cell in cells:
        #     target_cell.Value = target_cell.Value() + join + cell.Value()
        #     cell.Value = None

    @staticmethod
    def merge_area_rows(areas, join="\r\n"):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            values = Array.CreateInstance(object, 1, area.columns.count)
            for i,col in enumerate(area.columns):
                values[0,i] = join.join([cell.Text for cell in col.rows])
            area.ClearContents()
            area.Rows[1].Value = values

    @staticmethod
    def merge_area_cols(areas, join=", "):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            values = Array.CreateInstance(object, area.rows.count, 1)
            for i,row in enumerate(area.rows):
                values[i,0] = join.join([cell.Text for cell in row.columns])
            area.ClearContents()
            area.Columns[1].Value = values

    @staticmethod
    def split_to_cols(cells, sep=","):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            values_split = cell.Text.split(sep)
            values = Array.CreateInstance(object, 1, len(values_split))
            for i,col in enumerate(values_split):
                values[0,i] = col.strip()
            new_area = xllib.resize_areas([cell], cols=len(values_split))[0]
            new_area.Value = values

    @staticmethod
    def split_to_rows(cells, sep=None):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            if sep is None:
                values_split = cell.Text.splitlines()
            else:
                values_split = cell.Text.split(sep)
            values = Array.CreateInstance(object, len(values_split), 1)
            for i,row in enumerate(values_split):
                values[i,0] = row.strip()
            new_area = xllib.resize_areas([cell], rows=len(values_split))[0]
            new_area.Value = values


    @staticmethod
    def formula_to_values(areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            area.Value = area.Value()

    @staticmethod
    def values_to_showntext(areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            for cell in iter(area.Cells):
                if cell.Value2 is None:
                    continue
                cell.Value = "'" + cell.Text
            area.NumberFormat = "@" #Text

    @staticmethod
    def text_to_numbers(areas, application):
        if not xllib.confirm_no_undo(): return
        general_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        for area in areas:
            #area.NumberFormatLocal = general_format
            #area.Value = application.WorksheetFunction.NumberValue( area )
            for cell in iter(area.Cells):
                if cell.HasFormula or cell.Value2 is None:
                    continue
                if cell.NumberFormat == "@": #Text
                    cell.NumberFormatLocal = general_format
                try:
                    # cell.Value = Converter.ToDouble(cell.Value())
                    cell.Value = application.WorksheetFunction.NumberValue( cell )
                except:
                    cell.Value = cell.Value()

    @staticmethod
    def numbers_to_text(areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            area.NumberFormat = "@" #Text
            for cell in iter(area.Cells):
                if cell.Value2 is None:
                    continue
                cell.Value = "'" + cell.Text

    @staticmethod
    def text_to_datetime(areas, application):
        if not xllib.confirm_no_undo(): return
        general_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        for area in areas:
            for cell in iter(area.Cells):
                if cell.HasFormula or cell.Value2 is None or isinstance(cell.Value(), DateTime):
                    continue
                if cell.NumberFormat == "@": #Text
                    cell.NumberFormatLocal = general_format
                try:
                    cell.Value = DateTime.Parse( cell.Text )
                except:
                    pass

    @staticmethod
    def text_to_formula(areas, application):
        if not xllib.confirm_no_undo(): return
        general_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        for area in areas:
            #area.NumberFormatLocal = general_format
            #area.FormulaLocal = area.Value()
            for cell in iter(area.Cells):
                if cell.Text[0] != "=":
                    continue
                cell.NumberFormatLocal = general_format
                cell.FormulaLocal = cell.Value()

    @staticmethod
    def formula_to_text(areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            #area.NumberFormat = "@" #Text
            #area.Value = area.FormulaLocal
            for cell in iter(area.Cells):
                if not cell.HasFormula:
                    continue
                cell.NumberFormat = "@" #Text
                cell.Value = "'" + cell.FormulaLocal

    @staticmethod
    def formula_to_absolute(cells, application):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            if cell.HasFormula:
                cell.Formula = application.ConvertFormula(cell.Formula, 1, 1, 1) #xlA1, xlA1, xlAbsolute

    @staticmethod
    def formula_to_relative(cells, application):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            if cell.HasFormula:
                cell.Formula = application.ConvertFormula(cell.Formula, 1, 1, 4) #xlA1, xlA1, xlRelative
    
    @staticmethod
    def local_formula_to_english_text(cells):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            if cell.HasFormula:
                cell.Value = "'" + cell.Formula
    
    @staticmethod
    def english_text_to_local_formula(cells, application):
        if not xllib.confirm_no_undo(): return
        general_format = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
        for cell in cells:
            if cell.Text[0] != "=":
                continue
            cell.NumberFormatLocal = general_format
            cell.Formula = cell.Value()

    @staticmethod
    def prohibit_duplicates(areas, application):
        if not xllib.confirm_no_undo("Dies überschreibt bestehende Datenüberprüfungen und kann nicht rückgängig gemacht werden. Ausführen?"): return
        for area in areas:
            vali_form = "=COUNTIF(" + area.Address(True, True, 1) + "," + area.Cells(1).Address(False, False, 1) + ")=1" #xlA1
            vali_form = xllib.formula_int2local(vali_form)
            area.Validation.Delete()
            area.Validation.Add(7, 1, 1, vali_form) #xlValidateCustom, xlValidAlertStop, xlBetween
            #area.Validation.ShowError = True
            #area.Validation.ErrorTitle = "Duplicate Value"
            #area.Validation.ErrorMessage = "This value was already entered. All values must be unique. Please try again."

    @staticmethod
    def subtotal(application, selection, func="Sum"):
        try:
            selection = selection.SpecialCells(xlcon.XlCellType["xlCellTypeVisible"])
            value = str( getattr(application.WorksheetFunction, func)(selection) )
            #value = str(application.WorksheetFunction.Subtotal(xlcon.subtotalFunction[func], selection))

            value = value.replace('.', application.International(xlcon.XlApplicationInternational["xlDecimalSeparator"]))
            Forms.Clipboard.SetText(value)
        except:
            bkt.message('Fehler beim Kopieren!')
        #bkt.message('Kopiert: ' + value)

    @staticmethod
    def enabled_subtotal(application, selection):
        try:
            #count number of cells that contain numbers
            return application.WorksheetFunction.Count(selection) > 0
            #application.WorksheetFunction.Subtotal(xlcon.subtotalFunction["AVG"], selection)
            #return True
        except:
            return False

    @staticmethod
    def trim(application, areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            area.Value = application.WorksheetFunction.Trim(area)

    @staticmethod
    def clean(application, areas):
        if not xllib.confirm_no_undo(): return
        for area in areas:
            area.Value = application.WorksheetFunction.Clean(area)

    @staticmethod
    def trim_python(application, cells):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            cell.Value = cell.Text.strip()

    @staticmethod
    def fill_down(cells, application):
        if not xllib.confirm_no_undo(): return

        # to_be_filled = None
        for cell in cells:
            if cell.Row == 1:
                continue
            if cell.Value2 is None:
                try:
                    cell.Value = cell.Offset(-1,0).MergeArea(1).Value()
                    # to_be_filled = xllib.range_union(to_be_filled, cell, application)
                except:
                    pass

        # if to_be_filled is not None:
        #     to_be_filled.FormulaR1C1 = "=R[-1]C"
        #     to_be_filled.Value = to_be_filled.Value()
       
        # for area in areas:
        #     empty_cells = area.SpecialCells(4) #xlCellTypeBlanks
        #     empty_cells.FormulaR1C1 = "=R[-1]C"
        #     area.Value = area.Value()

    @staticmethod
    def undo_fill_down(cells, application):
        if not xllib.confirm_no_undo(): return

        to_be_deleted = None
        for cell in cells:
            if cell.Row == 1:
                continue
            try:
                if cell.Value() == cell.Offset(-1,0).MergeArea(1).Value():
                    to_be_deleted = xllib.range_union(to_be_deleted, cell)
            except:
                pass

        if to_be_deleted is not None:
            to_be_deleted.Value = None

    @classmethod
    def toggle_hidden_columns(cls, sheet, application, selection):
        area = sheet.UsedRange

        #Restore hidden columns if sheet is the same
        hidden_cols = cls._get_hidden_name("BKT_HIDDEN_COLS", sheet)
        #if sheet.Name in cls.hidden_columns:
        if hidden_cols is not None:
            #sheet.Range(cls.hidden_columns[sheet.Name]).EntireColumn.Hidden = True
            #del cls.hidden_columns[sheet.Name]
            #sheet.Range(hidden_cols).EntireColumn.Hidden = True
            hidden_cols.EntireColumn.Hidden = True
            #cls._del_hidden_name("BKT_HIDDEN_COLS", sheet)

        #Show hidden columns and store them
        else:
            #hidden_cols = None
            for i in range(1,area.Columns.Count+1):
                if area.Columns(i).EntireColumn.Hidden:
                    hidden_cols = xllib.range_union(hidden_cols, area.Columns(i).EntireColumn)

            if hidden_cols is not None:
                hidden_cols.EntireColumn.Hidden = False
                #cls.hidden_columns[sheet.Name] = hidden_cols.AddressLocal(False, False)
                #cls._set_hidden_name("BKT_HIDDEN_COLS", sheet, hidden_cols.AddressLocal(True, True))
                cls._set_hidden_name("BKT_HIDDEN_COLS", sheet, hidden_cols)
            
            #If entire rows are selected hide them
            elif selection.Address() == selection.EntireColumn.Address():
                selection.EntireColumn.Hidden = True
            
            else:
                bkt.message("Keine ausgeblendeten Spalten im genutzten Bereich gefunden.")


    @classmethod
    def toggle_hidden_rows(cls, sheet, application, selection):
        area = sheet.UsedRange

        #Restore hidden rows if sheet is the same
        hidden_rows = cls._get_hidden_name("BKT_HIDDEN_ROWS", sheet)
        #if sheet.Name in cls.hidden_rows:
        if hidden_rows is not None:
            #sheet.Range(cls.hidden_rows[sheet.Name]).EntireRow.Hidden = True
            #del cls.hidden_rows[sheet.Name]
            #sheet.Range(hidden_rows).EntireRow.Hidden = True
            hidden_rows.EntireRow.Hidden = True
            #cls._del_hidden_name("BKT_HIDDEN_ROWS", sheet)

        #Show hidden rows and store them
        else:
            #hidden_rows = None
            for i in range(1,area.Rows.Count+1):
                if area.Rows(i).EntireRow.Hidden:
                    hidden_rows = xllib.range_union(hidden_rows, area.Rows(i).EntireRow)

            if hidden_rows is not None:
                hidden_rows.EntireRow.Hidden = False
                #cls.hidden_rows[sheet.Name] = hidden_rows.AddressLocal(False, False)
                #cls._set_hidden_name("BKT_HIDDEN_ROWS", sheet, hidden_rows.AddressLocal(True, True))
                cls._set_hidden_name("BKT_HIDDEN_ROWS", sheet, hidden_rows)
            
            #If entire rows are selected hide them
            elif selection.Address() == selection.EntireRow.Address():
                selection.EntireRow.Hidden = True
            
            else:
                bkt.message("Keine ausgeblendeten Zeilen im genutzten Bereich gefunden.")


    @classmethod
    def remove_hidden_cols(cls, sheet):
        if not xllib.confirm_no_undo(): return

        xllib.freeze_app()

        deleted = 0
        try:
            area = sheet.UsedRange
            for i in range(1,area.Columns.Count+1):
                if area.Columns(i).EntireColumn.Hidden:
                    area.Columns(i).EntireColumn.Delete()
                    deleted += 1
        finally:
            xllib.unfreeze_app()
        
        bkt.message("Es wurden %s Spalten gelöscht" % deleted)



    @classmethod
    def remove_hidden_rows(cls, sheet):
        area = sheet.UsedRange
        deleted = 0
        for i in range(1,area.Rows.Count+1):
            if area.Rows(i).EntireRow.Hidden:
                area.Rows(i).EntireRow.Delete()
                deleted += 1
        bkt.message("Es wurden %s Zeilen gelöscht" % deleted)

    @staticmethod
    def show_all_cells(sheet):
        sheet.Columns.EntireColumn.Hidden = False
        sheet.Rows.EntireRow.Hidden = False

    @staticmethod
    def hide_unused_areas(sheet):
        selection = xllib.get_unused_ranges(sheet)

        for rng in selection:
            rng.Hidden = True


    @staticmethod
    def paste_on_visible(application, sheet, cell, pasteType=xlcon.XlPasteType["xlPasteAll"]):
        if not xllib.confirm_no_undo(): return

        xllib.freeze_app(disable_display_alerts=True)
        temporary_sheet = xllib.create_temp_sheet()

        try:
            with cell: #otherwise SystemError: Ein COM-Objekt, das vom zugrunde liegenden RCW getrennt wurde, kann nicht verwendet werden.
                temporary_sheet.Cells(cell.Row, cell.Column).PasteSpecial(pasteType)
                rows = temporary_sheet.UsedRange.Rows.Count
                cols = temporary_sheet.UsedRange.Columns.Count
                
                ### METHOD 1: COPY CELL BY CELL ###
                #FIXME: cache area of visible columns once determined in first loop
                # cur_cell = cell
                # for i in range(1,rows+1):
                #     for j in range(1,cols+1):
                #         temporary_sheet.UsedRange.Cells(i, j).Copy()
                #         cur_cell.PasteSpecial(pasteType)
                #         if j < cols:
                #             cur_cell = xllib.get_next_visible_cell(cur_cell, 'right')
                #     if i < rows:
                #         cur_cell = sheet.Cells(cur_cell.Row, cell.Column)
                #         cur_cell = xllib.get_next_visible_cell(cur_cell, 'bottom')
                # sheet.Range(cell, cur_cell).Select()
            
                ### METHOD 2: INSERT BLANKS AND PASTE USING SKIP BLANKS ###
                i = cell.Row
                rows_to_check = i+rows
                while i <= rows_to_check:
                    if sheet.Cells(i,1).EntireRow.Hidden:
                        temporary_sheet.Cells(i,1).EntireRow.Insert()
                        rows_to_check += 1
                    i += 1

                i = cell.Column
                cols_to_check = i+cols
                while i <= cols_to_check:
                    if sheet.Cells(1,i).EntireColumn.Hidden:
                        temporary_sheet.Cells(1,i).EntireColumn.Insert()
                        cols_to_check += 1
                    i += 1

                temporary_sheet.UsedRange.Copy()
                cell.PasteSpecial(pasteType, SkipBlanks=True)
            
        except:
            bkt.message("Sorry, etwas ist schiefgelaufen!?")
            logging.exception("Error pasting on visible cells")

        temporary_sheet.Delete()
        xllib.unfreeze_app()


class Format(object):
    @staticmethod
    def hide_zero(cells, application, pressed):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            if pressed:
                formats = cell.NumberFormat.split(";")
                formats = formats + ['']*(3-len(formats))
                formats[2] = ''
                cell.NumberFormat = ";".join(formats)
                #cell.NumberFormat = '0;;;@'
            else:
                if cell.NumberFormat == '0;;;@':
                    cell.NumberFormatLocal = application.International(xlcon.XlApplicationInternational["xlGeneralFormatName"])
                    return

                formats = cell.NumberFormat.split(";")
                if len(formats) == 3:
                    del formats[2]
                elif len(formats) >= 4:
                    #(.*) (.*),(0*).* (.*)
                    formats[2] = "0"
                cell.NumberFormat = ";".join(formats)

    @staticmethod
    def hide_zero_pressed(cell):
        formats = cell.NumberFormat.split(";")
        return len(formats) >= 3 and formats[2] == ''
        #return cell.NumberFormat == '0;;;@'

    @staticmethod
    def hide_zero_simple(cells, application):
        if not xllib.confirm_no_undo(): return
        for cell in cells:
            cell.NumberFormat = '0;;;@'

    @staticmethod
    def number_in_thousand(cells):
        if not xllib.confirm_no_undo(): return
        #TODO: Make buttons smart: recognize number format and adjust it instead of replacing it
        for cell in cells:
            cell.NumberFormat = '_-* #.##0,0. "k"_-;-* #.##0,0. "k"_-;_-* "-"? "k"_-;_-@_-'

    @staticmethod
    def number_in_million(cells):
        if not xllib.confirm_no_undo(): return
        #TODO: Make buttons smart: recognize number format and adjust it instead of replacing it
        for cell in cells:
            cell.NumberFormat = '_-* #.##0,0.. "Mio."_-;-* #.##0,0.. "Mio."_-;_-* "-"? "Mio."_-;_-@_-'

    @staticmethod
    def merged_cells_to_center_across(cells):
        if not xllib.confirm_no_undo(): return

        for cell in cells:
            if cell.MergeCells and cell.MergeArea.Rows.Count == 1 and cell.MergeArea.HorizontalAlignment == -4108: #xlCenter
                area = cell.MergeArea
                cell.MergeCells = False
                area.HorizontalAlignment = 7 #xlCenterAcrossSelection

    @staticmethod
    def merged_cells_to_unmerged_filled(cells):
        if not xllib.confirm_no_undo(): return

        for cell in cells:
            if cell.MergeCells:
                area = cell.MergeArea
                cell.MergeCells = False
                if cell.HasFormula:
                    area.Formula = cell.Formula
                else:
                    area.Value = cell.Value()

    @staticmethod
    def horiz_align(selection, alignment, pressed):
        if not xllib.confirm_no_undo(): return
        
        if not pressed:
            selection.HorizontalAlignment = 1 #xlGeneral
        else:
            selection.HorizontalAlignment = alignment

    @staticmethod
    def horiz_align_pressed(selection, alignment):
        return selection.HorizontalAlignment == alignment


zellen_inhalt_gruppe = bkt.ribbon.Group(
    id="group_cell_contents",
    label="Cell contents",
    image_mso="Formula",
    children=[
        bkt.ribbon.Button(
            id = 'apply_formula',
            label="Apply formula…",
            show_label=True,
            size='large',
            image_mso='Formula',
            supertip="Apply a formula to all selected cells.",
            on_action=bkt.Callback(CellsOps.apply_formula, cells=True, application=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.Menu(
            id = 'apply_regex',
            label="Apply regex…",
            show_label=True,
            size='large',
            image_mso='ApplyFilter',
            supertip="Apply a regular expression to all selected cells.",
            children=[
                bkt.ribbon.Button(
                    id = 'regex_match',
                    label="Count/filter with regex",
                    supertip="Write the number of results/groups of a regular expression for selected cells into the respective cell.",
                    on_action=bkt.Callback(CellsOps.regex_match_count, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'regex_split',
                    label="Split into columns with regex",
                    supertip="Split all results/groups of a regular expression for the selected cells into columns.",
                    on_action=bkt.Callback(CellsOps.regex_split_to_columns, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'regex_replace',
                    label="Search and replace with regex",
                    supertip="Search and replace in selected cells with a regular expression.",
                    on_action=bkt.Callback(CellsOps.regex_replace, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        ),
        bkt.ribbon.Menu(
            label="Text tools",
            show_label=True,
            image_mso='FormControlEditBox',
            screentip="Various text manipulations",
            supertip="Add or cut text.",
            children=[
                bkt.ribbon.Button(
                    id = 'prepend_text',
                    label="Prepend text…",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Prepend a text to all selected cells.",
                    on_action=bkt.Callback(CellsOps.prepend_text, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'append_text',
                    label="Append text…",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Append a text to all selected cells.",
                    on_action=bkt.Callback(CellsOps.append_text, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.ribbon.Button(
                    id = 'slice_text',
                    label="Cut text by position…",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Truncate a text at the front or back at a given position.",
                    on_action=bkt.Callback(CellsOps.slice_text, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'find_and_slice_text',
                    label="Cut text by character…",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Truncate a text at the front or back at a given character.",
                    on_action=bkt.Callback(CellsOps.find_and_slice_text, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(),
                bkt.mso.control.ReplaceDialog(),
            ]
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.Button(
                    id = 'formula_to_values',
                    label="Formulas to values",
                    show_label=True,
                    image_mso='ShowFormulas',
                    supertip="Replace formulas in all selected cells with their respective values. Cells without formulas remain unchanged.",
                    on_action=bkt.Callback(CellsOps.formula_to_values, areas=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.MenuSeparator(title="Values/text"),
                    bkt.ribbon.Button(
                        id = 'formula_to_values2',
                        label="Formulas to values",
                        show_label=True,
                        image_mso='ShowFormulas',
                        supertip="Replace formulas in all selected cells with their respective values. Cells without formulas remain unchanged.",
                        on_action=bkt.Callback(CellsOps.formula_to_values, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'values_to_showntext',
                        label="To displayed text",
                        show_label=True,
                        #image_mso='PasteTextOnly',
                        supertip="Replace values in all selected cells with the actually displayed text. The cell format is changed to 'Text'.",
                        on_action=bkt.Callback(CellsOps.values_to_showntext, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(title="Numbers/dates"),
                    bkt.ribbon.Button(
                        id = 'numbers_to_text',
                        label="Numeric values to text",
                        show_label=True,
                        #image_mso='PasteTextOnly',
                        supertip="Converts numeric values (numbers, date, time) into numbers stored as text. The cell format is changed to 'Text'.",
                        on_action=bkt.Callback(CellsOps.numbers_to_text, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'text_to_numbers',
                        label="Text to numbers",
                        show_label=True,
                        #image_mso='PasteValues',
                        supertip="Converts numbers stored as text into real numbers. The cell format is changed to 'General'.",
                        on_action=bkt.Callback(CellsOps.text_to_numbers, areas=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'text_to_datetime',
                        label="Text to date/time",
                        show_label=True,
                        #image_mso='PasteTextOnly',
                        supertip="Converts date and time values stored as text into a real date, with time if applicable. The cell format is changed to 'General'.",
                        on_action=bkt.Callback(CellsOps.text_to_datetime, areas=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(title="Formulas"),
                    bkt.ribbon.Button(
                        id = 'text_to_formula',
                        label="Text to formulas",
                        show_label=True,
                        #image_mso='PasteFormulas',
                        supertip="Converts formulas stored as text into real formulas. The cell format is changed to 'General'.",
                        on_action=bkt.Callback(CellsOps.text_to_formula, areas=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'formula_to_text',
                        label="Formulas to text",
                        show_label=True,
                        #image_mso='PasteTextOnly',
                        supertip="Converts formulas into formulas stored as text. The cell format is changed to 'Text'. Cells without formulas remain unchanged.",
                        on_action=bkt.Callback(CellsOps.formula_to_text, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'formula_to_absolute',
                        label="Formulas A1 to $A$1",
                        show_label=True,
                        #image_mso='PasteFormulas',
                        supertip="Converts references in formulas to absolute references.",
                        on_action=bkt.Callback(CellsOps.formula_to_absolute, cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'formula_to_relative',
                        label="Formulas $A$1 to A1",
                        show_label=True,
                        #image_mso='PasteFormulas',
                        supertip="Converts references in formulas to relative references.",
                        on_action=bkt.Callback(CellsOps.formula_to_relative, cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'eng_text_to_formula',
                        label="English formulas to formulas",
                        show_label=True,
                        #image_mso='PasteFormulas',
                        supertip="Converts English formulas stored as text into real formulas. The cell format is changed to 'General'.",
                        on_action=bkt.Callback(CellsOps.english_text_to_local_formula, cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'formula_to_eng_text',
                        label="Formulas to English formulas",
                        show_label=True,
                        #image_mso='PasteTextOnly',
                        supertip="Converts formulas into English formulas stored as text. The cell format is changed to 'Text'. Cells without formulas remain unchanged.",
                        on_action=bkt.Callback(CellsOps.local_formula_to_english_text, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.Button(
                    id = 'cells_trim',
                    label="Trim",
                    show_label=True,
                    image_mso='TextDirectionContext',
                    supertip="Remove superfluous spaces at the start or end of all selected cells (like the Excel function TRIM).",
                    on_action=bkt.Callback(CellsOps.trim, application=True, areas=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'cells_trim2',
                        label="Trim/shorten (trim)",
                        show_label=True,
                        image_mso='TextDirectionContext',
                        supertip="Remove superfluous spaces at the start or end of all selected cells (like the Excel function TRIM).",
                        on_action=bkt.Callback(CellsOps.trim, application=True, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'cells_trim_python',
                        label="Advanced trim/shorten (trim)",
                        show_label=True,
                        # image_mso='TextDirectionContext',
                        supertip="Remove superfluous spaces at the start or end of all selected cells using Python's strip function.",
                        on_action=bkt.Callback(CellsOps.trim_python, application=True, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'cells_clean',
                        label="Clean/clean up (clean)",
                        show_label=True,
                        #image_mso='TextDirectionContext',
                        supertip="Remove non-printable characters in all selected cells (like the Excel function CLEAN).",
                        on_action=bkt.Callback(CellsOps.clean, application=True, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.Button(
                    id = 'cells_fill_down',
                    label="Fill empty cells downward",
                    show_label=True,
                    image_mso='FillDown',
                    supertip="Fill empty cells in the selected range each with the filled cell above.",
                    on_action=bkt.Callback(CellsOps.fill_down, cells=True, application=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'cells_fill_down2',
                        label="Fill empty cells downward",
                        show_label=True,
                        image_mso='FillDown',
                        supertip="Fill empty cells in the selected range each with the filled cell above.",
                        on_action=bkt.Callback(CellsOps.fill_down, cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'cells_undo_fill_down',
                        label="Empty cells filled downward again",
                        show_label=True,
                        image_mso='FillUp',
                        supertip="Delete repeating cell values so that only the topmost cell remains filled.",
                        on_action=bkt.Callback(CellsOps.undo_fill_down, cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'cells_merge',
                        label="Merge all cell contents",
                        show_label=True,
                        # image_mso='FillUp',
                        supertip="Inserts all cells into the active cell, separated by line breaks",
                        on_action=bkt.Callback(CellsOps.merge_cells, selection=True, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'cells_merge_cols',
                        label="Merge column-wise with comma",
                        show_label=True,
                        # image_mso='FillUp',
                        supertip="Inserts all columns (per selection range) into the first column, separated by commas",
                        on_action=bkt.Callback(CellsOps.merge_area_cols, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'cells_merge_rows',
                        label="Merge row-wise with line break",
                        show_label=True,
                        # image_mso='FillUp',
                        supertip="Inserts all rows (per selection range) into the first row, separated by line breaks",
                        on_action=bkt.Callback(CellsOps.merge_area_rows, areas=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'cells_split_cols',
                        label="Split comma-separated into columns",
                        show_label=True,
                        # image_mso='FillUp',
                        supertip="Split cell contents into columns at each comma",
                        on_action=bkt.Callback(CellsOps.split_to_cols, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'cells_split_rows',
                        label="Split line breaks into rows",
                        show_label=True,
                        # image_mso='FillUp',
                        supertip="Split cell contents into rows at each line break",
                        on_action=bkt.Callback(CellsOps.split_to_rows, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        #TODO: Zellen mit gleichen Werten verbinden
        #TODO: Zellen nicht mehr verbinden und Werte in einzelne Zellen füllen
        bkt.ribbon.SplitButton(
            get_enabled = bkt.Callback(CellsOps.enabled_subtotal, application=True, selection=True),
            children=[
                bkt.ribbon.Button(
                    id = 'selection_subtotal_sum',
                    label="Copy sum of selected cells",
                    show_label=True,
                    image_mso='Copy',
                    supertip="Copy the sum over the selected visible cells to the clipboard.",
                    on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection), application=True, selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'selection_subtotal_sum2',
                        label="Copy sum of selected cells",
                        show_label=True,
                        image_mso='Copy',
                        supertip="Copy the sum over the selected visible cells to the clipboard.",
                        on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection), application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'selection_subtotal_avg',
                        label="Copy average of selected cells",
                        show_label=True,
                        #image_mso='Copy',
                        supertip="Copy the average over the selected visible cells to the clipboard.",
                        on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection, "Average"), application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'selection_subtotal_min',
                        label="Copy minimum of selected cells",
                        show_label=True,
                        #image_mso='Copy',
                        supertip="Copy the minimum over the selected visible cells to the clipboard.",
                        on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection, "Min"), application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'selection_subtotal_max',
                        label="Copy maximum of selected cells",
                        show_label=True,
                        #image_mso='Copy',
                        supertip="Copy the maximum over the selected visible cells to the clipboard.",
                        on_action=bkt.Callback(lambda application, selection: CellsOps.subtotal(application, selection, "Max"), application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.SplitButton(
            get_enabled = bkt.Callback(lambda: Forms.Clipboard.ContainsText()),
            children=[
                bkt.ribbon.Button(
                    id = 'paste_on_visible_all',
                    label="Insert into visible cells",
                    show_label=True,
                    image_mso='PasteTableByOverwritingCells',
                    supertip="Inserts the clipboard content only into visible cells. Hidden or filtered-out cells are skipped.",
                    on_action=bkt.Callback(CellsOps.paste_on_visible, application=True, sheet=True, cell=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'paste_on_visible_all2',
                        label="Insert into visible cells",
                        show_label=True,
                        image_mso='PasteTableByOverwritingCells',
                        supertip="Inserts the clipboard content only into visible cells. Hidden or filtered-out cells are skipped.",
                        on_action=bkt.Callback(CellsOps.paste_on_visible, application=True, sheet=True, cell=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'paste_on_visible_values',
                        label="Insert values into visible cells",
                        show_label=True,
                        image_mso='PasteValues',
                        supertip="Inserts the clipboard content as values only into visible cells. Hidden or filtered-out cells are skipped.",
                        on_action=bkt.Callback(lambda application, sheet, cell: CellsOps.paste_on_visible(application, sheet, cell, xlcon.XlPasteType["xlPasteValues"]), application=True, sheet=True, cell=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'paste_on_visible_formulas',
                        label="Insert formulas into visible cells",
                        show_label=True,
                        image_mso='PasteFormulas',
                        supertip="Inserts the clipboard content as formulas only into visible cells. Hidden or filtered-out cells are skipped.",
                        on_action=bkt.Callback(lambda application, sheet, cell: CellsOps.paste_on_visible(application, sheet, cell, xlcon.XlPasteType["xlPasteFormulas"]), application=True, sheet=True, cell=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        #TODO: Upper/Lower/Proper-Case
        #TODO: Formatierung gezielt übertragen (Auswahl Zellenformat, Benutzerdefinierte Format., Datenvalidierung)
        #TODO: Benutzerdefinierte Formatierung konsolidieren (wenn farbe und typ identisch, range_union)
        #TODO: Unit/Currency Conversion
    ]
)


zellen_format_gruppe = bkt.ribbon.Group(
    id="group_cell_formats",
    label="Cell formats",
    image_mso="TableColumnsInsertLeftExcel",
    children=[
        bkt.ribbon.SplitButton(
            size="large",
            children=[
                bkt.ribbon.Button(
                    id = 'toggle_hidden_columns',
                    label="Show/hide columns",
                    show_label=True,
                    image_mso='TableColumnsInsertLeftExcel',
                    supertip="Toggle all hidden columns between hidden and shown.\n\nIf no hidden columns are stored or present in the sheet, and columns are selected, these are hidden.",
                    on_action=bkt.Callback(CellsOps.toggle_hidden_columns, sheet=True, application=True, selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'toggle_hidden_columns2',
                        label="Show/hide columns",
                        show_label=True,
                        image_mso='TableColumnsInsertLeftExcel',
                        supertip="Toggle all hidden columns between hidden and shown.\n\nIf no hidden columns are stored or present in the sheet, and columns are selected, these are hidden.",
                        on_action=bkt.Callback(CellsOps.toggle_hidden_columns, sheet=True, application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'toggle_hidden_rows',
                        label="Show/hide rows",
                        show_label=True,
                        image_mso='TableRowsInsertBelowExcel',
                        supertip="Toggle all hidden rows between hidden and shown.\n\nIf no hidden rows are stored or present in the sheet, and rows are selected, these are hidden.",
                        on_action=bkt.Callback(CellsOps.toggle_hidden_rows, sheet=True, application=True, selection=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'show_all_cells',
                        label="Show all columns and rows",
                        show_label=True,
                        #image_mso='TableInsertMultidiagonalCell',
                        supertip="Show all hidden columns and rows again.",
                        on_action=bkt.Callback(CellsOps.show_all_cells, sheet=True, require_worksheet=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'hide_unused_areas',
                        label="Hide unused range",
                        show_label=True,
                        #image_mso='ViewGridlinesToggleExcel',
                        supertip="Hide all columns and rows of the unused range.",
                        on_action=bkt.Callback(CellsOps.hide_unused_areas, sheet=True, require_worksheet=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.MenuSeparator(),
                    bkt.ribbon.Button(
                        id = 'remove_hidden_cols',
                        label="Delete hidden columns",
                        show_label=True,
                        #image_mso='TableInsertMultidiagonalCell',
                        supertip="Delete all hidden columns.",
                        on_action=bkt.Callback(CellsOps.remove_hidden_cols, sheet=True, require_worksheet=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'remove_hidden_rows',
                        label="Delete hidden rows",
                        show_label=True,
                        #image_mso='TableInsertMultidiagonalCell',
                        supertip="Delete all hidden rows.",
                        on_action=bkt.Callback(CellsOps.remove_hidden_rows, sheet=True, require_worksheet=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                ])
            ]
        ),
        bkt.ribbon.Button(
            id = 'prohibit_duplicates',
            label="Forbid duplicates",
            show_label=True,
            image_mso='DataValidation',
            supertip="Forbids duplicates within the respectively selected ranges via a data validation. Existing data validations are overwritten.",
            on_action=bkt.Callback(CellsOps.prohibit_duplicates, areas=True, application=True),
            get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
        ),
        bkt.ribbon.SplitButton(
            children=[
                bkt.ribbon.ToggleButton(
                    id = 'hide_zero',
                    label="Show/hide 0",
                    show_label=True,
                    image='hide_zero',
                    screentip="Hide zero values",
                    supertip="Hide and show 0 values via cell format. An attempt is made to detect the existing cell format and adjust it accordingly.",
                    on_toggle_action=bkt.Callback(Format.hide_zero, cells=True, application=True),
                    get_pressed=bkt.Callback(Format.hide_zero_pressed, cell=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Menu(children=[
                    bkt.ribbon.Button(
                        id = 'hide_zero_simple',
                        label="Hide 0 values",
                        show_label=True,
                        image='hide_zero',
                        screentip="Hide zero values",
                        supertip="Hide 0 values via cell format ('0;;;@'). The existing cell format is overwritten.",
                        on_action=bkt.Callback(lambda cells, application: Format.hide_zero_simple(cells, application), cells=True, application=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'number_in_thousand',
                        label="Thousands to 0.0 k",
                        show_label=True,
                        image='number_in_thousand',
                        screentip="Display thousand amounts clearly",
                        supertip="Display thousand amounts as x.x k via cell format. The existing cell format is overwritten.",
                        on_action=bkt.Callback(Format.number_in_thousand, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    ),
                    bkt.ribbon.Button(
                        id = 'number_in_million',
                        label="Million values to 0.0 M",
                        show_label=True,
                        image='number_in_million',
                        screentip="Display million amounts clearly",
                        supertip="Display million amounts as x.x million via cell format. The existing cell format is overwritten.",
                        on_action=bkt.Callback(Format.number_in_million, cells=True),
                        get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                    )
                ]),
            ]
        ),
        bkt.ribbon.Menu(
            label="Cells & alignment",
            show_label=True,
            image_mso='AlignJustify',
            screentip="Replace merged cells and use unusual text alignments",
            #supertip="Add or cut text.",
            children=[
                bkt.ribbon.Button(
                    id = 'merged_cells_to_center_across',
                    label="Replace merged cells with center across selection",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Replaces merged cells within the current selection with the horizontal alignment 'Center across selection', provided the merged cells consist of one row and were previously centered.",
                    on_action=bkt.Callback(Format.merged_cells_to_center_across, cells=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.Button(
                    id = 'merged_cells_to_unmerged_filled',
                    label="Unmerge cells and distribute contents",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Unmerges merged cells within the current selection and inserts the original cell content into all cells.",
                    on_action=bkt.Callback(Format.merged_cells_to_unmerged_filled, cells=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.MenuSeparator(),
                # bkt.mso.control.AlignJustify,
                # bkt.mso.control.ParagraphDistributed,
                bkt.ribbon.ToggleButton(
                    id = 'halign_justify',
                    label="Justified",
                    show_label=True,
                    image_mso='AlignJustify',
                    supertip="Justify selected cells.",
                    on_toggle_action=bkt.Callback(lambda selection, pressed: Format.horiz_align(selection, -4130, pressed), selection=True), #xlHAlignJustify
                    get_pressed=bkt.Callback(lambda selection: Format.horiz_align_pressed(selection, -4130), selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.ToggleButton(
                    id = 'halign_distributed',
                    label="Evenly distributed",
                    show_label=True,
                    image_mso='ParagraphDistributed',
                    supertip="Align selected cells horizontally distributed (extreme justification).",
                    on_toggle_action=bkt.Callback(lambda selection, pressed: Format.horiz_align(selection, -4117, pressed), selection=True), #xlHAlignDistributed
                    get_pressed=bkt.Callback(lambda selection: Format.horiz_align_pressed(selection, -4117), selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.ToggleButton(
                    id = 'halign_centeracross',
                    label="Center across selection",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="Align selected cells 'Center across selection', i.e. merged cells are simulated.",
                    on_toggle_action=bkt.Callback(lambda selection, pressed: Format.horiz_align(selection, 7, pressed), selection=True), #xlHAlignCenterAcrossSelection 
                    get_pressed=bkt.Callback(lambda selection: Format.horiz_align_pressed(selection, 7), selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
                bkt.ribbon.ToggleButton(
                    id = 'halign_fill',
                    label="Fill (repeat text to the end)",
                    show_label=True,
                    #image_mso='FormControlEditBox',
                    supertip="'Fill' selected cells, i.e. cell content is visually repeated across the entire cell width.",
                    on_toggle_action=bkt.Callback(lambda selection, pressed: Format.horiz_align(selection, 5, pressed), selection=True), #xlHAlignFill
                    get_pressed=bkt.Callback(lambda selection: Format.horiz_align_pressed(selection, 5), selection=True),
                    get_enabled = bkt.CallbackTypes.get_enabled.dotnet_name,
                ),
            ]
        ),
        bkt.ribbon.DialogBoxLauncher(idMso='CellAlignmentOptions')
    ]
)

comments_gruppe = bkt.ribbon.Group(
    id="group_cell_comments",
    label="Comments",
    image_mso="ReviewNewComment",
    children=[
        bkt.mso.control.ReviewNewComment(size="large"),

        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.mso.control.ShapeChangeShapeGallery(),
            bkt.mso.control.ShapeFillColorPicker(),
            bkt.mso.control.ShapeOutlineColorPicker(),
        ]),

        bkt.ribbon.Box(box_style="horizontal", children=[
            bkt.mso.control.ReviewPreviousComment(),
            bkt.mso.control.ReviewNextComment(show_label=True),
        ]),

        bkt.ribbon.Box(box_style="horizontal", children=[
            # bkt.mso.control.ReviewDeleteComment(),
            # bkt.mso.control.ReviewShowOrHideComment(),
            bkt.mso.control.ReviewShowAllComments(show_label=True, label="All on/off"),
        ]),

        bkt.ribbon.DialogBoxLauncher(idMso='ObjectFormatDialog')
    ]
)