"""Private JSON-lines worker. Owns one isolated LibreOffice process and document."""
import argparse
import copy
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import uuid

import uno
from com.sun.star.beans import PropertyValue


def props(**values):
    result = []
    for name, value in values.items():
        prop = PropertyValue()
        prop.Name, prop.Value = name, value
        result.append(prop)
    return tuple(result)


def url(path):
    return uno.systemPathToFileUrl(str(Path(path).resolve()))


def enum(name, value):
    return uno.Enum(name, value)


def rgb(value):
    if not value or value == "transparent":
        return -1
    value = value.lstrip("#")
    if len(value) == 3:
        value = "".join(ch * 2 for ch in value)
    return int(value[-6:], 16)


FORMATS = {"normal": "General", "text": "@", "number": "#,##0.00",
           "number_plain": "0.00", "percent": "0.00%", "usd": "$#,##0.00",
           "eur": "€#,##0.00", "rmb": "¥#,##0.00", "date": "yyyy/m/d",
           "time": "h:mm:ss", "datetime": "yyyy/m/d h:mm", "duration": "[h]:mm:ss"}
FILTERS = {"xlsx": "Calc MS Excel 2007 XML", "ods": "calc8",
           "xls": "MS Excel 97", "csv": "Text - txt - csv (StarCalc)",
           "tsv": "Text - txt - csv (StarCalc)"}


def style_of(sheet, cell):
    styles = sheet.get("styles", [])
    index = cell.get("style")
    return styles[index] if isinstance(index, int) and index < len(styles) else {}


def cells_of(sheet):
    return {(int(ri), int(ci)): cell
            for ri, row in sheet.get("rows", {}).items() if ri != "len" and isinstance(row, dict)
            for ci, cell in row.get("cells", {}).items() if not cell.get("nativeDerived")}


def content(cell):
    # Cached/displayed values are never treated as user edits.
    return cell.get("text", ""), bool(cell.get("formulaPreview"))


class Engine:
    def __init__(self, office, root, input_file):
        self.root = Path(root).resolve()
        self.office, self.process, self.desktop, self.doc = office, None, None, None
        self.model = None
        self.names = []
        self.checkpoint = self.root / "recovery.ods"
        profile = self.root / "profile"
        (profile / "user").mkdir(parents=True, exist_ok=True)
        (profile / "user" / "registrymodifications.xcu").write_text(
            '<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry">'
            '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
            '<prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>'
            '<item oor:path="/org.openoffice.Office.Calc/Content/Update">'
            '<prop oor:name="Link" oor:op="fuse"><value>1</value></prop></item></oor:items>', encoding="utf-8")
        self.pipe = "ahmed_" + uuid.uuid4().hex
        self.log = open(self.root / "office.log", "ab")
        command = [office, "-env:UserInstallation=" + url(profile), "--headless",
                   "--nologo", "--nodefault", "--norestore", "--nofirststartwizard",
                   "--accept=pipe,name=" + self.pipe + ";urp;StarOffice.ComponentContext"]
        self.process = subprocess.Popen(command, stdout=self.log, stderr=self.log,
                                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        resolver = uno.getComponentContext().ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", uno.getComponentContext())
        deadline = time.monotonic() + 30
        while True:
            try:
                context = resolver.resolve("uno:pipe,name=" + self.pipe + ";urp;StarOffice.ComponentContext")
                self.context = context
                self.desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
                break
            except Exception:
                if time.monotonic() >= deadline:
                    raise RuntimeError("LibreOffice UNO connection timed out")
                time.sleep(.1)
        self.load(input_file)
        self.recover()

    def load(self, filename):
        if self.doc is not None:
            self.doc.setModified(False)
            self.doc.close(True)
        self.doc = self.desktop.loadComponentFromURL(url(filename), "_blank", 0,
            props(Hidden=True, ReadOnly=False,
                  UpdateDocMode=uno.getConstantByName("com.sun.star.document.UpdateDocMode.NO_UPDATE"),
                  MacroExecutionMode=uno.getConstantByName("com.sun.star.document.MacroExecMode.NEVER_EXECUTE")))
        if self.doc is None or not self.doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            raise RuntimeError("LibreOffice could not open a spreadsheet")
        self.auto = self.doc.isAutomaticCalculationEnabled()
        self.doc.enableAutomaticCalculation(False)
        self.names = list(self.doc.Sheets.ElementNames)
        self.formatter = self.context.ServiceManager.createInstanceWithContext("com.sun.star.util.NumberFormatter", self.context)
        self.formatter.attachNumberFormatsSupplier(self.doc)

    def recover(self):
        # Native checkpoint contains the complete document, never a grid reconstruction.
        next_path = self.export("ods", "recovery-next.ods")
        os.replace(next_path, self.checkpoint)

    def export(self, ext="xlsx", name=None):
        if ext not in FILTERS:
            raise ValueError("Unsupported export format: " + ext)
        destination = self.root / (name or ("snapshot." + ext))
        options = dict(FilterName=FILTERS[ext], Overwrite=True)
        if ext in ("csv", "tsv"):
            options["FilterOptions"] = ("9" if ext == "tsv" else "44") + ",34,76,1"
        self.doc.enableAutomaticCalculation(self.auto)
        try:
            self.doc.storeToURL(url(destination), props(**options))
        finally:
            self.doc.enableAutomaticCalculation(False)
        return str(destination)

    def apply_style(self, cell, old, new, old_data, new_data):
        font_old, font = old.get("font", {}), new.get("font", {})
        mappings = {"name": ("CharFontName", "Calibri"), "size": ("CharHeight", 11.0),
                    "bold": ("CharWeight", 100.0), "italic": ("CharPosture", None)}
        for key, (prop, default) in mappings.items():
            if font_old.get(key) == font.get(key):
                continue
            value = font.get(key, default)
            if key == "bold":
                value = 150.0 if value else 100.0
            if key == "italic":
                value = enum("com.sun.star.awt.FontSlant", "ITALIC" if value else "NONE")
            setattr(cell, prop, value)
        for key, prop, default in (("color", "CharColor", -1), ("bgcolor", "CellBackColor", -1),
                                    ("textwrap", "IsTextWrapped", False),
                                    ("underline", "CharUnderline", 0), ("strike", "CharStrikeout", 0)):
            if old.get(key) == new.get(key):
                continue
            value = new.get(key, default)
            if key in ("color", "bgcolor"):
                value = rgb(value) if isinstance(value, str) else value
            elif key in ("underline", "strike"):
                value = 1 if value else 0
            setattr(cell, prop, value)
        if old.get("align") != new.get("align"):
            value = {"left": "LEFT", "center": "CENTER", "right": "RIGHT"}.get(new.get("align"), "STANDARD")
            cell.HoriJustify = enum("com.sun.star.table.CellHoriJustify", value)
        if old.get("valign") != new.get("valign"):
            value = {"top": "TOP", "middle": "CENTER", "bottom": "BOTTOM"}.get(new.get("valign"), "STANDARD")
            cell.VertJustify = enum("com.sun.star.table.CellVertJustify", value)
        if old.get("format") != new.get("format") or old_data.get("numFmt") != new_data.get("numFmt"):
            code = new_data.get("numFmt") if new_data.get("numFmtFormat") == new.get("format") else None
            code = code or FORMATS.get(new.get("format"), "General")
            locale = uno.createUnoStruct("com.sun.star.lang.Locale")
            locale.Language, locale.Country = "en", "US"
            key = self.doc.NumberFormats.queryKey(code, locale, True)
            if key < 0:
                key = self.doc.NumberFormats.addNew(code, locale)
            cell.NumberFormat = key
        for side in ("top", "bottom", "left", "right"):
            before, after = old.get("border", {}).get(side), new.get("border", {}).get(side)
            if before == after:
                continue
            border = uno.createUnoStruct("com.sun.star.table.BorderLine2")
            if after:
                border.Color = rgb(after[1]) if len(after) > 1 else 0
                kind = after[0]
                border.LineStyle = {"dotted": 2, "dashed": 1, "double": 3}.get(kind, 0)
                border.LineWidth = {"thin": 26, "medium": 53, "thick": 79, "double": 79}.get(kind, 26)
                border.OuterLineWidth = border.LineWidth
            setattr(cell, side.title() + "Border2", border)

    def sync(self, model):
        if self.model is None:
            raise RuntimeError("Reload native state before editing after an interrupted document operation")
        if [s.get("name") for s in model] != [s.get("name") for s in self.model]:
            raise RuntimeError("Sheet changes must be sent through the native document command")
        # Reject unsupported interface mutations before changing any document data.
        for old, new in zip(self.model, model):
            for field in ("images", "backgroundImage", "validations", "sheetProtection", "hyperlinks", "freeze", "autofilter"):
                if old.get(field) != new.get(field):
                    raise RuntimeError("Editing " + field + " is not connected to the native backend yet. Use the desktop app.")
        needed = False
        for old, new in zip(self.model, model):
            a, b = cells_of(old), cells_of(new)
            for key in a.keys() | b.keys():
                before, after = a.get(key, {}), b.get(key, {})
                if (content(before) != content(after) or style_of(old, before) != style_of(new, after)
                        or before.get("comment") != after.get("comment")):
                    needed = True
                    break
            for field, property_name in (("rows", "height"), ("cols", "width")):
                for key, value in new.get(field, {}).items():
                    if isinstance(value, dict) and value.get(property_name) is not None and value.get(property_name) != old.get(field, {}).get(key, {}).get(property_name):
                        needed = True
            if set(old.get("merges", [])) != set(new.get("merges", [])):
                needed = True
        if not needed:
            self.model = copy.deepcopy(model)
            return {"changed": False}
        changed = False
        undo = self.doc.UndoManager
        undo.enterUndoContext("Spreadsheet edit")
        self.doc.lockControllers()
        try:
            for index, (old, new) in enumerate(zip(self.model, model)):
                sheet = self.doc.Sheets.getByIndex(index)
                a, b = cells_of(old), cells_of(new)
                for position in a.keys() | b.keys():
                    before, after = a.get(position, {}), b.get(position, {})
                    old_style, new_style = style_of(old, before), style_of(new, after)
                    if content(before) == content(after) and old_style == new_style and before.get("comment") == after.get("comment"):
                        continue
                    cell = sheet.getCellByPosition(position[1], position[0])
                    if content(before) != content(after):
                        text = after.get("text", "")
                        if after.get("formulaPreview") or (not after.get("previewCell") and text.startswith("=")):
                            # The English Excel formula grammar used by the imported grid.
                            parser = self.doc.createInstance("com.sun.star.sheet.FormulaParser")
                            parser.FormulaConvention = uno.getConstantByName("com.sun.star.sheet.AddressConvention.XL_A1")
                            mapper = self.doc.createInstance("com.sun.star.sheet.FormulaOpCodeMapper")
                            parser.OpCodeMap = mapper.getAvailableMappings(uno.getConstantByName("com.sun.star.sheet.FormulaLanguage.OOXML"), uno.getConstantByName("com.sun.star.sheet.FormulaMapGroup.ALL_EXCEPT_SPECIAL"))
                            parser.CompileEnglish = True
                            cell.setTokens(parser.parseFormula(text, cell.CellAddress))
                        elif text == "":
                            cell.clearContents(23)
                        elif after.get("hasSavedValue") and after.get("savedType") == "b":
                            cell.setFormula("=TRUE()" if after["savedValue"] else "=FALSE()")
                        elif new_style.get("format") == "text" or after.get("numFmt") == "@":
                            cell.setString(text)
                        elif text.startswith("'"):
                            cell.setString(text[1:])
                        else:
                            try:
                                format_key = self.formatter.detectNumberFormat(cell.NumberFormat, text)
                                number = self.formatter.convertStringToNumber(format_key, text)
                            except Exception:
                                cell.setString(text)
                            else:
                                cell.setValue(number)
                                cell.NumberFormat = format_key
                        changed = True
                    if old_style != new_style:
                        self.apply_style(cell, old_style, new_style, before, after)
                        changed = True
                    if before.get("comment") != after.get("comment"):
                        cell.Annotation.String = after.get("comment", "")
                        changed = True
                for ri, row in new.get("rows", {}).items():
                    if ri == "len" or not isinstance(row, dict):
                        continue
                    previous = old.get("rows", {}).get(ri, {})
                    if row.get("height") != previous.get("height") and row.get("height") is not None:
                        sheet.Rows.getByIndex(int(ri)).Height = round(row["height"] * 2540 / 96)
                        changed = True
                for ci, column in new.get("cols", {}).items():
                    if ci == "len" or not isinstance(column, dict):
                        continue
                    previous = old.get("cols", {}).get(ci, {})
                    if column.get("width") != previous.get("width") and column.get("width") is not None:
                        sheet.Columns.getByIndex(int(ci)).Width = round(column["width"] * 2540 / 96)
                        changed = True
                for ref in set(old.get("merges", [])) - set(new.get("merges", [])):
                    sheet.getCellRangeByName(ref).merge(False)
                    changed = True
                for ref in set(new.get("merges", [])) - set(old.get("merges", [])):
                    sheet.getCellRangeByName(ref).merge(True)
                    changed = True
        except Exception:
            self.doc.unlockControllers()
            undo.leaveUndoContext()
            self.load(self.checkpoint)
            raise
        else:
            self.doc.unlockControllers()
            undo.leaveUndoContext()
            self.model = copy.deepcopy(model)
            if changed:
                self.recover()
        return {"changed": changed}

    def operation(self, command):
        action = command["action"]
        if action in ("undo", "redo"):
            manager = self.doc.UndoManager
            possible = manager.isUndoPossible() if action == "undo" else manager.isRedoPossible()
            if possible:
                getattr(manager, action)()
                self.recover()
            return {"changed": bool(possible)}
        undo = self.doc.UndoManager
        undo.enterUndoContext(action)
        self.doc.lockControllers()
        try:
            index = int(command.get("sheet", 0))
            sheet = self.doc.Sheets.getByIndex(index)
            if action in ("insertRows", "deleteRows", "insertColumns", "deleteColumns"):
                collection = sheet.Rows if "Rows" in action else sheet.Columns
                method = collection.insertByIndex if action.startswith("insert") else collection.removeByIndex
                method(int(command["index"]), int(command.get("count", 1)))
            elif action == "renameSheet":
                sheet.Name = command["name"]
            elif action == "addSheet":
                self.doc.Sheets.insertNewByName(command["name"], int(command.get("index", self.doc.Sheets.Count)))
            elif action == "deleteSheet":
                self.doc.Sheets.removeByName(sheet.Name)
            elif action == "duplicateSheet":
                self.doc.Sheets.copyByName(sheet.Name, command["name"], int(command["index"]))
            elif action == "moveSheet":
                self.doc.Sheets.moveByName(sheet.Name, int(command["index"]))
            else:
                raise RuntimeError("Native operation is not supported: " + action)
        except Exception:
            self.doc.unlockControllers()
            undo.leaveUndoContext()
            self.load(self.checkpoint)
            raise
        else:
            self.doc.unlockControllers()
            undo.leaveUndoContext()
            self.names = list(self.doc.Sheets.ElementNames)
            self.recover()
        return {"changed": True}

    def dispatch(self, request):
        method = request["method"]
        if method == "bind":
            self.model = request["model"]
            if [s.get("name") for s in self.model] != list(self.doc.Sheets.ElementNames):
                raise RuntimeError("Viewer and native document have different sheets")
            return {"bound": True}
        if method == "sync":
            return self.sync(request["model"])
        if method == "operation":
            return self.operation(request["command"])
        if method == "calculate":
            self.doc.calculateAll()
            self.recover()
            return {"path": self.export("xlsx")}
        if method == "export":
            # Save returns native workbook bytes. It never reconstructs a document from the grid.
            self.doc.calculateAll()
            self.recover()
            return {"path": self.export(request.get("ext", "xlsx"))}
        if method == "inspect":
            return {"macroMode": next((p.Value for p in self.doc.getArgs() if p.Name == "MacroExecutionMode"), None),
                    "sheets": list(self.doc.Sheets.ElementNames),
                    "names": list(self.doc.NamedRanges.ElementNames),
                    "undo": self.doc.UndoManager.isUndoPossible(),
                    "redo": self.doc.UndoManager.isRedoPossible()}
        raise RuntimeError("Unknown native worker method")

    def close(self):
        try:
            if self.doc:
                self.doc.setModified(False)
                self.doc.close(True)
            if self.desktop:
                self.desktop.terminate()
        finally:
            if self.process:
                try:
                    self.process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    self.process.kill()
            if hasattr(self, "log"):
                self.log.close()


def main():
    arguments = argparse.ArgumentParser()
    arguments.add_argument("--office", required=True)
    arguments.add_argument("--root", required=True)
    arguments.add_argument("--input", required=True)
    args = arguments.parse_args()
    engine = None
    try:
        engine = Engine.__new__(Engine)
        engine.__init__(args.office, args.root, args.input)
        print(json.dumps({"ready": True, "pid": os.getpid()}), flush=True)
        for line in sys.stdin:
            request = json.loads(line)
            if request.get("method") == "close":
                print(json.dumps({"id": request["id"], "result": {"closed": True}}), flush=True)
                break
            try:
                result = engine.dispatch(request)
                print(json.dumps({"id": request["id"], "result": result}, ensure_ascii=False), flush=True)
            except Exception as error:
                print(json.dumps({"id": request.get("id"), "error": str(error)}, ensure_ascii=False), flush=True)
    except Exception as error:
        print(json.dumps({"fatal": str(error)}, ensure_ascii=False), flush=True)
    finally:
        if engine:
            engine.close()


if __name__ == "__main__":
    main()
