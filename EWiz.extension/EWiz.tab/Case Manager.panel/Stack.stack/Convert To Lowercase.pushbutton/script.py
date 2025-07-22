# -*- coding: utf-8 -*-
from __future__ import unicode_literals     # Py2 – make all literals unicode
__title__   = "Convert To Lowercase"
__doc__     = """Version = 1.7.2
Date    = July 2025
========================================
Description:
A text-cleanup utility for Revit that iterates through every **TextNote**
in the active document, applies strict sentence‐case to each sentence,
and then restores any project, system or default exception literals to
their original capitalization.
========================================
How-To:
1. Click the button on the ribbon.
2. Wait for the script to process all TextNotes.
3. Review the updated notes.
4. If some instances are lowercased incorrectly - use "Configure Exceptions" script to add new exception.  
========================================
Author: AO"""

"""
Convert To Lowercase
===============

A text-clean-up tool for Revit.
It walks every **TextNote** in the current document, converts each sentence to
strict sentence-case, then restores any strings listed in the three-tier
exception system (default / system / project).

Key features
------------
* Sentence segmentation with safe handling of dotted abbreviations
  (e.g. “u.n.o.”, “T.B.C.”).
* Pre-restoration of single- and multi-word literals so their capitalisation
  survives the blanket case change.
* Domain-aware normalisations:
    • engineering units and superscripts
    • section-profile shorthand (e.g. “100x100x6 SHS”)
    • quantities like “4No.” (no extra spaces)
    • typographic apostrophes for five key possessives.
* Skips text notes inside model groups.
* Writes a rotating per-user / per-model log (token-level diff).
* IronPython 2.7 -- Revit 2024 compatible.
"""
from Autodesk.Revit.DB import *

import clr
clr.AddReference('System')
from System.Collections.Generic import List
app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document

import os
import sys
import re
import difflib
import atexit
import clr

# -----------------------------------------------------------------------------
# .NET / pyRevit imports
# -----------------------------------------------------------------------------
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons

clr.AddReference('RevitServices')
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Transaction,
    ElementId,
)

# -----------------------------------------------------------------------------
# Local library path
# -----------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR    = os.path.normpath(os.path.join(SCRIPT_DIR, '..', '..', '..', '..', 'lib'))
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

# -----------------------------------------------------------------------------
# Logging
# -----------------------------------------------------------------------------
from logging_util import get_logger

LOGGER = get_logger('ExceptionsManager', filename_override='ExceptionsManager')
SRC    = 'ChangeRegister'
LOGGER.info('note processed', extra={'src': SRC})

# -----------------------------------------------------------------------------
# Metadata
# -----------------------------------------------------------------------------
VERSION   = '1.7.2'
USERNAME  = os.environ.get('USERNAME', 'unknown_user')
try:
    _rvt_path  = __revit__.ActiveUIDocument.Document.PathName
    MODEL_NAME = os.path.splitext(os.path.basename(_rvt_path))[0] if _rvt_path else 'unsaved_doc'
except Exception:
    MODEL_NAME = 'unsaved_doc'

# Flush logs on exit
atexit.register(lambda: [h.flush() for h in LOGGER.handlers])

# -----------------------------------------------------------------------------
# Exception manager
# -----------------------------------------------------------------------------
import exception_manager as EM
from exception_manager import ExceptionManager

REGEX_MAP = ExceptionManager.get_single_regexps()

# -----------------------------------------------------------------------------
# Sentence segmentation helpers
# -----------------------------------------------------------------------------
_DOTTED_LITERALS = [
    re.escape(lit.rstrip('.'))
    for lit in ExceptionManager.get_single_literals_map().values()
    if lit.endswith('.')
]
_ABBR = '|'.join(_DOTTED_LITERALS) or r'__never__'

# Split on “.?! ” that are *not* part of a protected abbreviation
SENT_END = re.compile(
    r'(?<!\b(?:' + _ABBR + r')\.)[.?!]\s+',
    re.IGNORECASE,
)

NO_SENTENCE_DOT  = [u'u.n.o.', u't.b.c.']
_ABBR_SAFE_DOT   = u'\uFE52'                         # small Unicode dot

# -----------------------------------------------------------------------------
# Inline heuristics
# -----------------------------------------------------------------------------
def is_bulleted_or_numbered_line(line):
    """Return *True* for a leading digit+dot or bullet character."""
    return bool(re.match(r'^\s*(?:\d+\.\s+|[-*•]\s+)', line))


def split_into_sentences(text):
    """
    Yield sentence / delimiter pairs without duplicating the punctuation.

    Example:  "Hello. World\n" →
              [("Hello", ". "), ("World", "\n")]
    """
    pos, out = 0, []
    for match in SENT_END.finditer(text):
        out.append((text[pos:match.start(0)], match.group(0)))
        pos = match.end(0)
    out.append((text[pos:], ''))
    return out


def apply_strict_sentence_case(line):
    """Lower-case the sentence then title-case the first alphabetical token."""
    segments   = split_into_sentences(line)
    out_chunks = []

    for seg_text, delim in segments:
        tokens, found_alpha = re.split(r'(\s+)', seg_text), False
        new_tokens = []

        for tok in tokens:
            if not found_alpha and re.search(r'[A-Za-z]', tok):
                tok, found_alpha = tok[:1].upper() + tok[1:].lower(), True
            else:
                tok = tok.lower()
            new_tokens.append(tok)

        out_chunks.append(''.join(new_tokens) + delim)
    return ''.join(out_chunks)


def protect_abbrev(text):
    """Swap dots in dotted abbreviations for the safe dot so they survive split."""
    for abbr in NO_SENTENCE_DOT:
        pattern = re.compile(r'(?i)\b' + re.escape(abbr) + r'\b')
        text    = pattern.sub(lambda m: m.group(0).replace('.', _ABBR_SAFE_DOT), text)
    return text


def restore_abbrev(text):
    """Reverse :func:`protect_abbrev`."""
    return text.replace(_ABBR_SAFE_DOT, u'.')

# -----------------------------------------------------------------------------
# Tokeniser – keeps punctuation tokens separate so we can re-assemble unchanged
# -----------------------------------------------------------------------------
TOKEN_RE = re.compile(
    r'([A-Za-z](?:\.[A-Za-z])+\.?)|'     # dotted abbreviation (u.n.o.)
    r'([A-Za-z0-9&/]+)([=);:,]*)|'       # word with optional trailing chars
    r'(/m²)|'                            # explicit “/m²” token
    r'(\s+)|'                            # whitespace
    r'([^\w\s]+)'                        # single punctuation mark
)


def tokenize_with_punct(text):
    """Return a list of (is_word, token) pairs preserving all whitespace/punct."""
    out = []
    for m in TOKEN_RE.finditer(text):
        abbr, word, tail, slash_unit, ws, other = m.groups()
        if abbr:
            out.append((True, abbr))
        elif word:
            out.append((True, word))
            if tail:
                out.append((False, tail))
        elif slash_unit:
            out.append((False, slash_unit))
        elif ws:
            out.append((False, ws))
        elif other:
            out.append((False, other))
    return out

# -----------------------------------------------------------------------------
# Engineering-specific helpers
# -----------------------------------------------------------------------------
PROFILE_RE = re.compile(
    r'^(?:(?P<prefix>\d+(?:\.\d+)?(?:x\d+(?:\.\d+)?)*)(?P<lit>(chs|shs|rhs|ea|pfc|ubp?|ub|uc|ua|vb|wp|x|d))'
    r'(?P<suffix>(?:x\d+(?:\.\d+)?)*\d*(?:\.\d+)*)|'
    r'(?P<lit2>(chs|shs|rhs|ea|pfc|ubp?|ub|uc|ua|vb|wp|x|d))'
    r'(?P<number>\d+(?:\.\d+)?(?:x\d+(?:\.\d+)?)*))$',
    re.I,
)


def normalise_profile(token):
    """
    Convert steel-section shorthand to canonical form.

    Example:  100X100x6shs  →  100x100x6SHS
    """
    match = PROFILE_RE.match(token.replace(' ', ''))
    if not match:
        return None

    lit = (match.group('lit') or match.group('lit2')).upper()
    pre = match.group('prefix') or ''
    suf = match.group('suffix') or match.group('number') or ''
    return (pre + lit + suf).replace('X', 'x')


def convert_subscripts(text):
    """Replace 2/3 exponents in m, mm tokens with proper superscripts."""
    text = re.sub(
        r'(\d)(m)([23])\b',
        lambda m: m.group(1) + 'm' + (u'²' if m.group(3) == '2' else u'³'),
        text,
        flags=re.I,
    )
    for k, v in {'mm2': 'mm²', 'mm3': 'mm³', 'm2': 'm²', 'm3': 'm³'}.items():
        text = re.sub(r'(?i)\b' + k + r'\b', v, text)
    return text


def restore_literals(line):
    """Apply every replacement regex from *REGEX_MAP* then hard-code BS / EN."""
    for rx, canon in REGEX_MAP.items():
        line = rx.sub(canon, line)
    line = re.sub(r'\bbs\b', 'BS', line, flags=re.I)
    line = re.sub(r'\ben\b', 'EN', line, flags=re.I)
    return line

# -----------------------------------------------------------------------------
# Core “enforcer” for word tokens
# -----------------------------------------------------------------------------
class ExceptionApplier(object):
    """Apply all exception rules to individual word tokens."""

    def __init__(self):
        data          = EM.ExceptionManager.load_all()
        self.literals = {p for items in data.values() for p in items if ' ' not in p}

        # Unit canonical map
        self.unit_map = {
            'n': 'N',         'kn': 'kN',       'n/m': 'N/m',    'kn/m': 'kN/m',
            'n/m²': 'N/m²',   'n/mm²': 'N/mm²', 'kn/m²': 'kN/m²','kn/mm²': 'kN/mm²',
            'knm': 'kNm',
            'm': 'm', 'm²': 'm²', 'm³': 'm³',
            'mm': 'mm', 'mm²': 'mm²', 'mm³': 'mm³',
            'hz': 'Hz', 'mpa': 'MPa',
        }
        unit_pattern      = '|'.join(re.escape(u) for u in self.unit_map)
        self.numeric_re   = re.compile(r'^(?P<num>\d+(?:\.\d+)?)?(?P<unit>(' + unit_pattern + r'))$', re.I)

        # Possessive fix-ups (“architect’s” → “Architect’s” etc.)
        self.apostrophe_map = {
            "architect's": u"Architect’s",
            "contractor's": u"Contractor’s",
            "engineer's": u"Engineer’s",
            "manufacturer's": u"Manufacturer’s",
            "sub-contractor's": u"Sub-contractor’s",
        }
        # Also accept the curly apostrophe input
        self.apostrophe_map.update({k.replace("'", u"’"): v for k, v in self.apostrophe_map.items()})

        # Section-profile literals and misc
        codes                = data.get('Section Profiles & Steel Sections', [])
        self.sec_codes       = {c.lower(): c for c in codes}
        self.single_letter_codes = {'t', 'd'}

        # Detect surrounding punctuation
        self.bracket_pat = re.compile(
            r'^(?P<open>[\(\[]?)'
            r'(?P<core>[^\)\]\.,;:?!]+)?'
            r'(?P<close>[\)\]]?)'
            r'(?P<punct>[.,;:?!]*)?$'
        )

    # ---------------------------------------------------------------------
    # Public
    # ---------------------------------------------------------------------
    def enforce(self, token):
        """
        Return *token* after applying every domain rule.
        The surrounding punctuation / brackets are preserved verbatim.
        """
        token_conv = convert_subscripts(token)

        # Try *whole-token* steel-profile first (fast-exit)
        prof = normalise_profile(token_conv)
        if prof:
            return prof

        match = self.bracket_pat.match(token_conv)
        if not match:                                 # no brackets
            return self._enforce_core(token_conv)

        # Handle leading / trailing punctuation
        open_, core, close, punct = (
            match.group('open') or '',
            match.group('core') or '',
            match.group('close') or '',
            match.group('punct') or '',
        )
        enforced = self._enforce_core(core)
        return open_ + enforced + close + punct

    # ---------------------------------------------------------------------
    # Internal helpers
    # ---------------------------------------------------------------------
    def _enforce_core(self, tok):
        """Apply rules to the word *without* punctuation context."""
        tok_low = tok.lower()

        # Unit after number (20kN → 20kN, 8.4kN/m² → same canonical form)
        num_unit = self.numeric_re.match(tok)
        if num_unit:
            num   = num_unit.group('num') or ''
            canon = self.unit_map.get(num_unit.group('unit').lower(), '')
            return num + canon

        # “4No.” without nbsp
        if re.match(r'^\d+no\.?$', tok_low):
            n = re.match(r'^\d+', tok).group(0)
            dot = '.' if tok_low.endswith('no.') else ''
            return u'%sNo%s' % (n, dot)

        # Apostrophe possessives
        if tok_low in self.apostrophe_map:
            return self.apostrophe_map[tok_low]

        # Literal (single-word) overrides
        if tok_low in {l.lower() for l in self.literals}:
            return next(l for l in self.literals if l.lower() == tok_low)

        # Section profiles embedded at either end (e.g. “75x75EA”)
        for lc, rc in self.sec_codes.items():
            if tok_low.startswith(lc):
                return rc + tok[len(lc):]
            if tok_low.endswith(lc):
                return tok[:-len(lc)] + rc

        # No change
        return tok

    # ---------------------------------------------------------------------
    def apply_apostrophe_exceptions(self, text):
        """One final pass to ensure curly apostrophes are in place."""
        for raw, canon in self.apostrophe_map.items():
            text = re.sub(r'\b' + re.escape(raw) + r'\b', canon, text, flags=re.I)
        return text


# -----------------------------------------------------------------------------
# Note-level processing
# -----------------------------------------------------------------------------
def convert_text_note_text(text, applier):
    """Convert one note, return (new_text, change_pairs)."""
    text      = protect_abbrev(text)
    out_lines, changes = [], []

    for line in text.splitlines(True):
        if is_bulleted_or_numbered_line(line):
            out_lines.append(line)
            continue

        sent = apply_strict_sentence_case(line)
        sent = restore_literals(sent)

        orig_toks = tokenize_with_punct(line)
        sent_toks = tokenize_with_punct(sent)

        for (io, old), (iw, new) in zip(orig_toks, sent_toks):
            if io and iw and old != new:
                changes.append((old, new))

        rebuilt = []
        for is_word, tok in sent_toks:
            if is_word:
                new_tok = applier.enforce(tok)
                if new_tok != tok:
                    changes.append((tok, new_tok))
                rebuilt.append(new_tok)
            else:
                rebuilt.append(tok)

        out_lines.append(''.join(rebuilt))

    return restore_abbrev(''.join(out_lines)), changes


# -----------------------------------------------------------------------------
# Revit transaction wrapper
# -----------------------------------------------------------------------------
def update_text_notes_to_sentence_case(doc):
    """Main entry: run over every TextNote in the model."""
    notes = (FilteredElementCollector(doc)
             .OfCategory(BuiltInCategory.OST_TextNotes)
             .WhereElementIsNotElementType()
             .ToElements())

    applier = ExceptionApplier()
    updated = skipped = total_changes = 0

    txn = Transaction(doc, 'Change Register: Sentence-case')
    txn.Start()
    try:
        for note in notes:
            if note.GroupId != ElementId.InvalidElementId:
                skipped += 1
                continue

            old_text = note.Text
            new_text, changes = convert_text_note_text(old_text, applier)
            new_text = applier.apply_apostrophe_exceptions(new_text)

            if old_text != new_text:
                note.Text = new_text
                updated  += 1

                if changes:
                    total_changes += len(changes)
                    pairs = ['%s>%s' % (o, c) for o, c in changes]
                    LOGGER.info('note_id=%s changes=%s', note.Id.IntegerValue, ','.join(pairs))
                else:
                    for diff_line in difflib.unified_diff(
                            old_text.splitlines(),
                            new_text.splitlines(),
                            fromfile='old', tofile='new', lineterm=''):
                        LOGGER.info('note_id=%s diff: %s', note.Id.IntegerValue, diff_line)

        txn.Commit()
        LOGGER.info('completed run', extra={'note_cnt': updated,
                                            'skipped_cnt': skipped,
                                            'total_changes': total_changes})
        MessageBox.Show(
            u'Updated: {0}\nSkipped (grouped): {1}'.format(updated, skipped),
            'Change Register', MessageBoxButtons.OK)

    except Exception as exc:
        txn.RollBack()
        LOGGER.error('fatal_error=%s', exc)
        MessageBox.Show(u'Transaction rolled back:\n{0}'.format(exc),
                        'Change Register Error', MessageBoxButtons.OK)


# -----------------------------------------------------------------------------
# pyRevit entry-point
# -----------------------------------------------------------------------------
if __name__ == '__main__':
    try:
        update_text_notes_to_sentence_case(__revit__.ActiveUIDocument.Document)
    except Exception as exc:
        LOGGER.error('startup_error=%s', exc)
        MessageBox.Show(str(exc), 'Change Register Error', MessageBoxButtons.OK)

#from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
#kit_button_clicked(btn_name=__title__)                  # Display Default Print Message