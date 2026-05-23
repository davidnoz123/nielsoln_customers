"""word_tools.py -- COM-based Word automation (pywin32 / win32com.client).

Provides:
  - Auto-install of pywin32 if not present (_pywin32_install, _get_win32com_client)
  - WordDriver class: COM lifecycle (open, detach, close) + document writing helpers
  - Heading, paragraph, table insertion via the Word object model

Usage pattern
-------------
    driver = WordDriver()
    driver.open()                         # creates a new blank document
    driver.add_heading("My Doc", level=1)
    driver.add_paragraph("Some text here.")
    driver.add_table([["Col A", "Col B"], ["row1a", "row1b"]])
    driver.save("C:\\path\\to\\output.docx")
    driver.close()

Or attach to an existing file:
    driver = WordDriver(document_path="C:\\path\\to\\existing.docx")
    driver.open()
    ...

import runpy ; temp = runpy._run_module_as_main("word_tools")
"""

import importlib
import os
import sys


def _log(msg: str) -> None:
    print(msg, flush=True)


# ---------------------------------------------------------------------------
# pywin32 installer (needs post-install script + process re-exec)
# ---------------------------------------------------------------------------

_PYWIN32_RESTART_FLAG = "_PYWIN32_RESTARTED"
_win32com_client_cache = None


def _pywin32_install() -> None:
    import glob
    import subprocess
    _log("Installing pywin32 ...")
    subprocess.run(
        [sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "pywin32"],
        check=True,
    )
    scripts_dir = os.path.dirname(sys.executable)
    candidates = glob.glob(os.path.join(scripts_dir, "pywin32_postinstall.py"))
    if candidates:
        _log("Running pywin32 post-install ...")
        subprocess.run([sys.executable, candidates[0], "-install"])
    else:
        _log("WARNING: pywin32_postinstall.py not found -- win32 DLLs may not be registered")
    _log("Restarting Python to load pywin32 DLLs ...")
    os.environ[_PYWIN32_RESTART_FLAG] = "1"
    os.execv(sys.executable, [sys.executable] + sys.argv)


def _get_win32com_client():
    global _win32com_client_cache
    if _win32com_client_cache is not None:
        return _win32com_client_cache
    try:
        _win32com_client_cache = importlib.import_module("win32com.client")
    except ImportError:
        if os.environ.get(_PYWIN32_RESTART_FLAG):
            raise RuntimeError("pywin32 was installed but win32com.client is still not importable")
        _pywin32_install()  # re-execs; code below never reached
    return _win32com_client_cache


# ---------------------------------------------------------------------------
# Word style name constants (English locale defaults)
# ---------------------------------------------------------------------------

# wdStyleHeading1 .. wdStyleHeading9 = -2, -3, ..., -10  (Word built-in enum)
# Using string names is safer across locales.
_HEADING_STYLES = {
    1: "Heading 1",
    2: "Heading 2",
    3: "Heading 3",
    4: "Heading 4",
}
_STYLE_NORMAL      = "Normal"
_STYLE_LIST_BULLET = "List Bullet"
_STYLE_LIST_NUMBER = "List Number"

# wdContentControl type constants used by insert_text_cc / insert_checkbox_cc.
_CC_TEXT     = 1   # wdContentControlRichText (Range.Text is read/write; wdContentControlText=2 is COM-protected)
_CC_CHECKBOX = 8   # wdContentControlCheckBox
_CC_PICTURE  = 3   # wdContentControlPicture  — holds an embedded image

# wdContentControlAppearance constants for the appearance parameter of insert_text_cc.
_CC_APPEARANCE_BOX    = 1   # wdContentControlBoundingBox  — blue border (default)
_CC_APPEARANCE_TAGS   = 2   # wdContentControlTags          — start/end tag markers
_CC_APPEARANCE_HIDDEN = 3   # wdContentControlHidden        — no visual indicator


# ---------------------------------------------------------------------------
# WordDriver
# ---------------------------------------------------------------------------

class WordDriver:
    """pywin32 wrapper for creating and editing Word documents via COM.

    Lifecycle:
        driver = WordDriver()           # new blank doc
        driver = WordDriver(path)       # existing file
        driver.open()
        ...
        driver.save(path)               # optional: save to a specific path
        driver.close(save=True)         # or close with auto-save
    """

    def __init__(self, document_path: str | None = None) -> None:
        self._document_path = document_path
        self._word = None
        self._doc  = None
        self._doc_owned = True  # False when we attached to an already-open document

    # ------------------------------------------------------------------
    # COM lifecycle
    # ------------------------------------------------------------------

    def open(self) -> None:
        """Attach to an already-open document by FullName, or open/create from disk."""
        win32com_client = _get_win32com_client()
        try:
            self._word = win32com_client.GetActiveObject("Word.Application")
            _log("WordDriver.open: attached to running Word instance")
        except Exception:
            self._word = win32com_client.Dispatch("Word.Application")
            _log("WordDriver.open: launched new Word instance")

        self._word.Visible = True

        if self._document_path:
            norm = self._document_path.replace("/", "\\").lower()
            for doc in self._word.Documents:
                if doc.FullName.lower() == norm:
                    self._doc = doc
                    self._doc_owned = False
                    _log(f"WordDriver.open: attached to already-open '{doc.Name}'")
                    return
            if os.path.exists(self._document_path):
                self._doc = self._word.Documents.Open(self._document_path)
                _log(f"WordDriver.open: opened '{self._document_path}'")
            else:
                self._doc = self._word.Documents.Add()
                _log(f"WordDriver.open: created new document (will save to '{self._document_path}')")
        else:
            self._doc = self._word.Documents.Add()
            _log("WordDriver.open: created new blank document")

    def detach(self) -> None:
        """Release COM references without closing Word."""
        self._doc  = None
        self._word = None

    def close(self, save: bool = False) -> None:
        """Close the document (and quit Word if this instance opened it)."""
        if self._doc is not None:
            if self._doc_owned:
                # wdDoNotSaveChanges=0, wdSaveChanges=-1
                self._doc.Close(SaveChanges=-1 if save else 0)
            self._doc = None
        if self._word is not None:
            if self._doc_owned:
                self._word.Quit()
            self._word = None

    def save(self, path: str | None = None) -> None:
        """Save the document.  Uses *path* if given, else saves in place."""
        if path:
            # FileFormat 16 = wdFormatXMLDocument (.docx)
            self._doc.SaveAs2(FileName=path, FileFormat=16)
            _log(f"WordDriver.save: saved to '{path}'")
        else:
            self._doc.Save()
            _log("WordDriver.save: saved in place")

    # ------------------------------------------------------------------
    # Content helpers
    # ------------------------------------------------------------------

    def _end_range(self):
        """Return a Range collapsed to the very end of the document."""
        rng = self._doc.Content
        rng.Collapse(Direction=0)  # wdCollapseEnd = 0
        return rng

    def add_heading(self, text: str, level: int = 1) -> None:
        """Append a heading paragraph at the end of the document."""
        style = _HEADING_STYLES.get(level, f"Heading {level}")
        rng = self._end_range()
        rng.InsertParagraphAfter()
        rng.Collapse(Direction=0)
        rng.InsertAfter(text)
        rng.Select()
        self._word.Selection.Style = self._doc.Styles(style)
        self._word.Selection.Collapse(Direction=0)

    def add_paragraph(self, text: str, style: str = _STYLE_NORMAL) -> None:
        """Append a normal (or styled) paragraph at the end of the document."""
        rng = self._end_range()
        rng.InsertParagraphAfter()
        rng.Collapse(Direction=0)
        rng.InsertAfter(text)
        rng.Select()
        self._word.Selection.Style = self._doc.Styles(style)
        self._word.Selection.Collapse(Direction=0)

    def add_bullet(self, text: str) -> None:
        """Append a bullet-list paragraph."""
        self.add_paragraph(text, style=_STYLE_LIST_BULLET)

    def add_table(self, rows: list[list[str]], header_row: bool = True) -> None:
        """Append a table.  First row is formatted as a header if header_row=True.

        rows -- list of lists; all rows must have the same number of columns.
        """
        if not rows:
            return
        num_rows = len(rows)
        num_cols = len(rows[0])

        # Move to end and insert a blank paragraph to anchor the table
        rng = self._end_range()
        rng.InsertParagraphAfter()
        rng.Collapse(Direction=0)

        tbl = self._doc.Tables.Add(rng, num_rows, num_cols)

        for r_idx, row in enumerate(rows):
            for c_idx, cell_text in enumerate(row):
                cell = tbl.Cell(r_idx + 1, c_idx + 1)
                cell.Range.Text = str(cell_text)

        if header_row and num_rows > 0:
            # Bold the first row
            header = tbl.Rows(1)
            header.Range.Bold = True

        # Move cursor past the table
        rng_after = tbl.Range
        rng_after.Collapse(Direction=0)
        rng_after.Select()

    def add_page_break(self) -> None:
        """Insert a page break at the current end of the document."""
        rng = self._end_range()
        # wdPageBreak = 7
        rng.InsertBreak(Type=7)

    def clear(self) -> None:
        """Delete all content in the document."""
        self._doc.Content.Delete()

    # ------------------------------------------------------------------
    # Document inspection
    # ------------------------------------------------------------------

    def scan(self) -> dict:
        """Return the open document's structure as a dict.

        Keys:
            paragraphs      -- list of {style, text}
            tables          -- list of list-of-list-of-str (rows x cells)
            bookmarks       -- list of bookmark name strings
            content_controls-- list of {title, tag, type, text}
        """
        result = {
            "paragraphs": [],
            "tables": [],
            "bookmarks": [],
            "content_controls": [],
        }

        for para in self._doc.Paragraphs:
            result["paragraphs"].append({
                "style": para.Style.NameLocal,
                "text": para.Range.Text.strip(),
            })

        for tbl in self._doc.Tables:
            table_data = []
            for row in tbl.Rows:
                row_data = [cell.Range.Text.strip() for cell in row.Cells]
                table_data.append(row_data)
            result["tables"].append(table_data)

        for bm in self._doc.Bookmarks:
            result["bookmarks"].append(bm.Name)

        for cc in self._doc.ContentControls:
            result["content_controls"].append({
                "title": cc.Title,
                "tag": cc.Tag,
                "type": cc.Type,
                "text": cc.Range.Text.strip(),
            })

        return result

    # ------------------------------------------------------------------
    # Content control insertion and manipulation
    # ------------------------------------------------------------------

    def table_cell_range(self, table_index: int, row: int, col: int):
        """Return a Range collapsed to the start of a table cell's content.

        All indices are 1-based.  Collapsing to start avoids the degenerate
        zero-length range that results from trimming the cell terminator on an
        empty cell, which causes ContentControls.Add to return None.
        Pass the result directly to insert_text_cc or insert_checkbox_cc.
        """
        rng = self._doc.Tables(table_index).Cell(row, col).Range
        rng.Collapse(Direction=1)   # wdCollapseStart
        return rng

    def table_cell_content_range(self, table_index: int, row: int, col: int):
        """Return the full content Range of a table cell, excluding the cell terminator.

        Unlike table_cell_range, this returns the entire cell content (not
        collapsed), suitable for character searches via find_char_positions_in_range.
        All indices are 1-based.
        """
        rng = self._doc.Tables(table_index).Cell(row, col).Range
        rng.MoveEnd(Unit=1, Count=-1)   # wdCharacter=1; trim trailing cell mark
        return rng

    def insert_text_cc(self, rng, tag: str, title: str, placeholder: str = "",
                       appearance: int = _CC_APPEARANCE_BOX):
        """Insert a plain-text content control at *rng*.

        The content control is inserted at the collapsed position represented by
        *rng* (see table_cell_range).  Returns the new ContentControl COM object.
        Font size is captured from *rng* before insertion and restored afterwards
        to prevent the CC from expanding the table cell.

        *appearance*: one of _CC_APPEARANCE_BOX (default, blue border),
                      _CC_APPEARANCE_TAGS (start/end markers), or
                      _CC_APPEARANCE_HIDDEN (no visual indicator — CC looks like
                      plain text, sized exactly to its content).
        """
        try:
            font_size = rng.Font.Size
        except Exception:
            font_size = None
        cc = self._doc.ContentControls.Add(Type=_CC_TEXT, Range=rng)
        cc.Tag = tag
        cc.Title = title
        cc.Appearance = appearance
        if font_size and font_size > 0:
            cc.Range.Font.Size = font_size
        if placeholder:
            try:
                # PlaceholderText is a BuildingBlock with a Range property;
                # not all Word COM versions expose this reliably.
                pt = cc.PlaceholderText
                if pt is not None:
                    pt.Range.Text = placeholder
                    if font_size and font_size > 0:
                        pt.Range.Font.Size = font_size
            except Exception:
                # Fallback: set the CC's actual text content
                cc.Range.Text = placeholder
        return cc

    def insert_checkbox_cc(self, rng, tag: str, title: str, checked: bool = False):
        """Insert a checkbox content control at *rng*.

        Font size is captured from *rng* before insertion and restored on the
        CC range afterwards so the checkbox renders at the cell's native size.
        Returns the new ContentControl COM object.
        """
        try:
            font_size = rng.Font.Size
        except Exception:
            font_size = None
        cc = self._doc.ContentControls.Add(Type=_CC_CHECKBOX, Range=rng)
        cc.Tag = tag
        cc.Title = title
        cc.Checked = checked
        if font_size and font_size > 0:
            cc.Range.Font.Size = font_size
        return cc

    def find_char_positions_in_range(self, container_rng, char: str) -> list:
        """Return a list of int document positions for every occurrence of *char*
        that falls within *container_rng*.

        Positions are document-level character offsets suitable for use with
        self._doc.Range(pos, pos+1).  The list is in ascending (document) order.
        """
        container_end = container_rng.End
        positions = []
        rng = container_rng.Duplicate
        finder = rng.Find
        finder.ClearFormatting()
        finder.Text = char
        finder.MatchCase = True
        finder.Forward = True
        finder.Wrap = 0     # wdFindStop
        while finder.Execute() and rng.Start < container_end:
            positions.append(rng.Start)
            rng.Collapse(Direction=0)   # wdCollapseEnd — advance past found char
        return positions

    def replace_positions_with_checkbox_ccs(self, items: list) -> int:
        """Replace single characters at the given document positions with checkbox CCs.

        *items*: list of (position, tag, title) in DESCENDING position order.
        Processing end-to-start ensures earlier positions are not shifted by
        each insertion.  Returns the number of content controls inserted.
        """
        count = 0
        for pos, tag, title in items:
            rng = self._doc.Range(pos, pos + 1)
            # Capture font size from the character being replaced so the
            # checkbox CC renders at the same size as the surrounding text.
            try:
                font_size = rng.Font.Size
            except Exception:
                font_size = None
            rng.Delete()
            rng.Collapse(Direction=1)   # wdCollapseStart
            cc = self._doc.ContentControls.Add(Type=_CC_CHECKBOX, Range=rng)
            cc.Tag = tag
            cc.Title = title
            cc.Checked = False
            if font_size and font_size > 0:
                cc.Range.Font.Size = font_size
            count += 1
        return count

    def clear_para_suffix_for_cc(self, prefix_text: str):
        """Find the first paragraph whose text starts with *prefix_text*.

        Deletes everything after the prefix (e.g. trailing underscores) up to
        but not including the paragraph mark, then returns a Range collapsed
        to that insertion point — ready for insert_text_cc or insert_picture_cc.
        Returns None if no matching paragraph is found.
        """
        for para in self._doc.Paragraphs:
            if para.Range.Text.startswith(prefix_text):
                insert_at = para.Range.Start + len(prefix_text)
                content_end = para.Range.End - 1   # one position before paragraph mark
                if content_end > insert_at:
                    self._doc.Range(insert_at, content_end).Delete()
                rng = self._doc.Range(insert_at, insert_at)
                return rng
        return None

    def insert_picture_cc(self, rng, tag: str, title: str):
        """Insert a picture content control at *rng*.  Returns the CC."""
        cc = self._doc.ContentControls.Add(Type=_CC_PICTURE, Range=rng)
        cc.Tag = tag
        cc.Title = title
        return cc

    def set_cc_picture(self, tag: str, image_path: str) -> bool:
        """Embed an image into a picture content control identified by *tag*.

        Replaces any existing image in the CC.  The image is saved into the
        document (LinkToFile=False).  Returns True if the tag was found.

        Note: Word raises 'selection partially covers a plain text content
        control' if the document has an active cursor inside a text CC when
        this is called.  Use insert_picture_at_tag() instead for signature
        fields — it deletes the CC first, bypassing the guard entirely.
        """
        for cc in self._doc.ContentControls:
            if cc.Tag == tag:
                rng = cc.Range
                for i in range(rng.InlineShapes.Count, 0, -1):
                    rng.InlineShapes(i).Delete()
                rng.InlineShapes.AddPicture(
                    FileName=image_path,
                    LinkToFile=False,
                    SaveWithDocument=True,
                )
                return True
        return False

    def insert_picture_at_tag(self, tag: str, image_path: str,
                              max_width_pt: float = 515.3,
                              max_height_pt: float = 49.6) -> bool:
        """Insert a picture at a content control or bookmark identified by *tag*.

        Designed for signature fields where Word's AddPicture guard fires on
        picture CCs adjacent to text CCs.

        Strategy:
        - If a bookmark named *tag* already exists (re-sign): insert directly.
        - Otherwise find the picture CC, delete it (removes the guard), add a
          bookmark named *tag* at that position, then insert the picture.

        The image is scaled down to fit within *max_width_pt* x *max_height_pt*
        (points; 72 pt = 1 inch) while preserving aspect ratio.  It is never
        upscaled.  The bookmark persists after insertion, so re-signing works.
        Returns True if the tag was found (as CC or bookmark).
        """
        def _fit(shape):
            w, h = float(shape.Width), float(shape.Height)
            if w > 0 and h > 0:
                scale = min(max_width_pt / w, max_height_pt / h, 1.0)
                shape.LockAspectRatio = -1   # wdTrue
                shape.Width = w * scale
            shape.Range.InsertAfter("\r")    # new paragraph after signature

        # ── Re-sign: bookmark already present from a previous signature ──────
        if self._doc.Bookmarks.Exists(tag):
            bm_start = self._doc.Bookmarks(tag).Range.Start
            # Expand to the full paragraph so we catch the picture even if the
            # bookmark didn't grow to include it after the previous insertion.
            para_rng = self._doc.Range(bm_start, bm_start)
            para_rng.Expand(1)   # wdParagraph = 1
            for i in range(para_rng.InlineShapes.Count, 0, -1):
                para_rng.InlineShapes(i).Delete()
            # Refresh bookmark (deletion may have shifted it)
            fresh_rng = self._doc.Range(bm_start, bm_start)
            self._doc.Bookmarks.Add(tag, fresh_rng)
            shape = self._doc.Bookmarks(tag).Range.InlineShapes.AddPicture(
                FileName=image_path,
                LinkToFile=False,
                SaveWithDocument=True,
            )
            _fit(shape)
            return True

        # ── First sign: find picture CC, delete it, insert at plain range ────
        for cc in self._doc.ContentControls:
            if cc.Tag == tag:
                start = cc.Range.Start
                cc.Delete(DeleteContents=True)
                fresh_rng = self._doc.Range(start, start)
                self._doc.Bookmarks.Add(tag, fresh_rng)
                shape = self._doc.Bookmarks(tag).Range.InlineShapes.AddPicture(
                    FileName=image_path,
                    LinkToFile=False,
                    SaveWithDocument=True,
                )
                _fit(shape)
                return True

        return False

    def set_cc_value(self, tag: str, value) -> bool:
        """Set a content control's value by tag.

        *value*: bool for checkbox CCs, str for text CCs.
        Returns True if a matching tag was found and set, False otherwise.
        """
        for cc in self._doc.ContentControls:
            if cc.Tag == tag:
                if isinstance(value, bool):
                    cc.Checked = value
                else:
                    cc.Range.Text = str(value)
                return True
        return False

    def get_cc_values(self) -> dict:
        """Return {tag: value} for all tagged content controls.

        Checkbox CCs map to bool; text CCs map to str (stripped).
        Untagged controls are skipped.
        """
        result = {}
        for cc in self._doc.ContentControls:
            if not cc.Tag:
                continue
            if cc.Type == _CC_CHECKBOX:
                result[cc.Tag] = cc.Checked
            else:
                result[cc.Tag] = cc.Range.Text.strip()
        return result


# ---------------------------------------------------------------------------
# Module-level helpers
# ---------------------------------------------------------------------------

def make_writable_copy(src_path: str, dst_path: str) -> None:
    """Copy *src_path* to *dst_path* and strip the read-only attribute.

    Safe when *src_path* is marked read-only (e.g. a protected template).
    Uses shutil.copy2 to preserve timestamps, then ensures the destination
    is writable before returning.
    """
    import shutil
    import stat
    shutil.copy2(src_path, dst_path)
    mode = os.stat(dst_path).st_mode
    os.chmod(dst_path, mode | stat.S_IWRITE)
    _log(f"make_writable_copy: {os.path.basename(src_path)} → {dst_path}")


def scan_document(path: str) -> dict:
    """Open *path* in Word via COM, print its structure, and return the result dict.

    Prints every non-empty paragraph with its style name, all table cell contents,
    any bookmarks, and any content controls (title/tag/text).  Useful for
    inspecting a template before automating it.
    """
    abs_path = os.path.abspath(path)
    driver = WordDriver(document_path=abs_path)
    driver.open()
    try:
        result = driver.scan()
    finally:
        driver.detach()

    _log(f"\n--- Document scan: {os.path.basename(abs_path)} ---")
    _log(f"Paragraphs ({len(result['paragraphs'])}):")
    for p in result["paragraphs"]:
        if p["text"]:
            _log(f"  [{p['style']}]  {p['text'][:120]}")

    if result["tables"]:
        _log(f"\nTables ({len(result['tables'])}):")
        for i, tbl in enumerate(result["tables"]):
            cols = len(tbl[0]) if tbl else 0
            _log(f"  Table {i + 1}: {len(tbl)} rows x {cols} cols")
            for row in tbl:
                _log("    " + " | ".join(c[:50] for c in row))

    if result["bookmarks"]:
        _log(f"\nBookmarks: {result['bookmarks']}")

    if result["content_controls"]:
        _log(f"\nContent Controls ({len(result['content_controls'])}):")
        for cc in result["content_controls"]:
            _log(f"  title={cc['title']!r}  tag={cc['tag']!r}  "
                 f"type={cc['type']}  text={cc['text'][:60]!r}")

    _log("--- end scan ---\n")
    return result


# ---------------------------------------------------------------------------
# Standalone test / REPL entry point
# ---------------------------------------------------------------------------

def _safe_local_imports(g: dict) -> None:
    try:
        pass  # no non-stdlib imports needed in __main__ beyond this module itself
    except Exception:
        import traceback as _tb
        _log(f"FATAL: _safe_local_imports failed:\n{_tb.format_exc()}")
        raise


def main() -> int:
    import tempfile

    _log("word_tools.py: smoke test -- creating a sample document")
    out_path = os.path.join(tempfile.gettempdir(), "word_tools_test.docx")

    driver = WordDriver(document_path=out_path)
    driver.open()

    driver.clear()
    driver.add_heading("Task Summary", level=1)
    driver.add_heading("Overview", level=2)
    driver.add_paragraph(
        "This is a sample document created by word_tools.py to verify COM automation."
    )
    driver.add_heading("Requirements", level=2)
    driver.add_bullet("Requirement one")
    driver.add_bullet("Requirement two")
    driver.add_bullet("Requirement three")
    driver.add_heading("Reference Table", level=2)
    driver.add_table(
        [
            ["Field", "Value"],
            ["Project", "Test Project"],
            ["Pay Rate", "$25.00/hr"],
            ["Tasks", "10"],
        ],
        header_row=True,
    )
    driver.save(out_path)
    _log(f"word_tools.py: saved to {out_path}")
    driver.detach()

    return 0


if __name__ == "__main__":
    _safe_local_imports(globals())
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:
        import traceback
        _log(f"FATAL unhandled exception:\n{traceback.format_exc()}")
        raise
