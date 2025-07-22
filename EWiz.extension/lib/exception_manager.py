# -*- coding: utf-8 -*-
"""
Exception Manager
===============
Version   : 1.7.2
Date      : 2025-05-15
Author    : AO

A three-tier exception system used by the Change Register tool.

Tiers
-----
1. **Built-in defaults**      –  ``DEFAULT_EXCEPTIONS`` below
2. **System-wide overrides**  –  *system_exceptions.json* (per-PC)
3. **Project overrides**      –  *<model>_exceptions.json*
   (`Manage Exceptions` UI writes to this file)

Key features
------------
* Single-word and multi-word literals restored with original capitalisation.
* Unit normalisation (ASCII ↔ Unicode superscript variants).
* Section-profile & equals-sign literal handling.
* Punctuation and parentheses preserved.
* Compatible with IronPython 2.7 / Revit 2024.
"""

# ───────────────────────────── imports ────────────────────────────────────
import os
import json
import re
import atexit
import datetime
import getpass
from collections import OrderedDict

# WinForms UI (for error pop-ups)
from System.Windows.Forms import MessageBox, MessageBoxButtons

from logging_util import get_logger

# ────────────────────────── constants / paths ─────────────────────────────
LIB_DIR     = os.path.dirname(__file__)
SYSTEM_FILE = os.path.join(LIB_DIR, 'system_exceptions.json')

VERSION  = '1.7.2'
USER     = getpass.getuser()

# Current RVT model (for logging only – *not* for file rotation)
doc_path   = __revit__.ActiveUIDocument.Document.PathName   # noqa: F821
model_name = os.path.splitext(os.path.basename(doc_path))[0] if doc_path else 'unsaved_model'

# ────────────────────────────── logger ─────────────────────────────────────
SRC    = 'ExcepMgr'                      # log-source tag
logger = get_logger('ExceptionsManager', filename_override='ExceptionsManager')
LOG_KW = {'src': SRC, 'version': VERSION, 'user': USER}

logger.info('Exception manager initialised for model “%s”', model_name, extra=LOG_KW)

# ───────────────────── default exception catalogue ────────────────────────
# (trimmed for brevity – content unchanged)

DEFAULT_EXCEPTIONS = {
   # B2: Aggressive Chemical Environment (ACEC)
    'Aggressive Chemical Environment (ACEC)': [
        'ACEC',  # Prefix for aggressive chemical environment classification.
        'AC-1',  # ACEC class AC-1: mild chemical environment.
        'AC-1s',  # ACEC class AC-1s: mild with soluble salts present.
        'AC-2',  # ACEC class AC-2: moderate chemical environment.
        'AC-2s',  # ACEC class AC-2s: moderate with soluble salts present.
        'AC-2z',  # ACEC class AC-2z: moderate with elevated chloride levels.
        'AC-3s',  # ACEC class AC-3s: severe with soluble salts present.
        'AC-3z',  # ACEC class AC-3z: severe with elevated chloride levels.
        'AC-4',  # ACEC class AC-4: very severe chemical environment.
        'AC-4s',  # ACEC class AC-4s: very severe with soluble salts present.
        'AC-4z',  # ACEC class AC-4z: very severe with elevated chloride levels.
        'AC-5z',  # ACEC class AC-5z: extreme with elevated chloride levels.
    ],

    # B2: Bolt Sizes
    'Bolt Sizes': [
        'M4',  # Nominal diameter 4 mm coarse-thread bolt.
        'M6',  # Nominal diameter 6 mm coarse-thread bolt.
        'M8',  # Nominal diameter 8 mm coarse-thread bolt.
        'M10',  # Nominal diameter 10 mm coarse-thread bolt.
        'M12',  # Nominal diameter 12 mm coarse-thread bolt.
        'M16',  # Nominal diameter 16 mm coarse-thread bolt.
        'M20',  # Nominal diameter 20 mm coarse-thread bolt.
        'M24',  # Nominal diameter 24 mm coarse-thread bolt.
        'M30',  # Nominal diameter 30 mm coarse-thread bolt.
        'M36',  # Nominal diameter 36 mm coarse-thread bolt.
        'M42',  # Nominal diameter 42 mm coarse-thread bolt.
        'HD',  # Hexagonal head bolt.
        'MS',  # Mild steel bolt.
        'Hilti Hit-Hy',  # Chemical anchor and bolt system by Hilti.
    ],

    # B2: Coating & Surface-Finish Codes
    'Coating & Surface-Finish Codes': [
        'HDG', # Hot-dip galvanised.
        'PRE', # Pre-painted.
        'ZP',  # Zinc-phosphate primer.
    ],

    # B2: Concrete Grades
    'Concrete Grades': [
        'C10',     # Concrete grade C10.
        'C12/15',  # Concrete grade C12/15.
        'C15',     # Concrete grade C15.
        'C16/20',  # Concrete grade C16/20.
        'C20',     # Concrete grade C20.
        'C25',     # Concrete grade C25.
        'C28/35',  # Concrete grade C28/35.
        'C25/30',  # Concrete grade C28/35.
        'C32/40',  # Concrete grade C32/40.
        'C35',     # Concrete grade C35.
        'C35/45',  # Concrete grade C35/45.
        'C40',     # Concrete grade C40.
        'C45/55',  # Concrete grade C45/55.
        'C50/60',  # Concrete grade C50/60.
        'C55/67',  # Concrete grade C55/67.
        'C60/75',  # Concrete grade C60/75.
        'C70/85',  # Concrete grade C70/85.
        'C7/8',    # Concrete grade C7/8.
        'C80/95',  # Concrete grade C80/95.
        'C90/105', # Concrete grade C90/105.
        'C100/115',# Concrete grade C100/115.
    ],

    # B2: Condition & Process Abbreviations
    'Condition & Process Abbreviations': [
        'dp',     # Deep.
        'GALV',   # Galvanised.
        'FW',     # Fillet weld.
        'ss',     # Stainless steel.
    ],

    # B2: Concrete Types
    'Concrete Types': [
        'Gen0',  # General-purpose unreinforced concrete mix (e.g., blinding).
        'Gen1',  # Standard structural concrete mix for foundations and slabs.
        'Gen2',  # Reinforced concrete mix for beams and columns.
        'Gen3',  # High-strength structural concrete mix for load-bearing elements.
        'Gen4',  # Specialist high-performance concrete mix (e.g., precast or rapid-strength).
    ],

    # B2: Cost & Measurement Sums (per RICS NRM2)
    'Cost & Measurement Sums': [
        'DW',       # Daywork.
        'L.S.',     # Lump sum.
        'PC Sum',   # Prime cost sum.
        'Prov Sum', # Provisional sum.
        'Class',     # Plastic cross-section.
    ],

    # B2: Design Sulfate Classes (DC)
    'Design Sulfate Classes (DC)': [
        'DC-1',       # Sulfate class DC-1.
        'DC-2',       # Sulfate class DC-2.
        'DC-2z',      # Sulfate class DC-2z.
        'DC-3',       # Sulfate class DC-3.
        'DC-3z',      # Sulfate class DC-3z.
        'DC-4',       # Sulfate class DC-4.
        'DC-4z',      # Sulfate class DC-4z.
        'DC-4z+APM3', # Sulfate class DC-4z+APM3.
        'FND1',       # Foundation concrete mix
        'FND2',       # Foundation concrete mix
        'FND3',       # Foundation concrete mix
        'FND4',       # Foundation concrete mix
        'FND5',       # Foundation concrete mix
        'DS-1',       # Sulfate class DS-1.
        'DS-2',       # Sulfate class DS-2.
        'DS-3',       # Sulfate class DS-3.
        'DS-4',       # Sulfate class DS-4.
    ],

    # B2: Discipline & Coordination Tags
    'Discipline & Coordination Tags': [
        'CDM',    # Construction design & management.
        'DRG',    # Drawing.
        'Ex',     # Existing.
        'MEP',    # Mechanical/electrical/plumbing.
        'NTS',    # Not to scale.
        'SOP',    # Setting out point.
        'TBC',    # To be confirmed.
        'T.B.C.', # To be confirmed (variant).
    ],

    # B2: Drawing & Layout Abbreviations
    'Drawing & Layout Abbreviations': [
        '2D',  # Two-dimensional representation.
        '3D',  # Three-dimensional representation.
        'CL',  # Centre line.
        'c/c',  # Centre-to-centre spacing.
        'CRS',  # Centres (alternative abbreviation).
        'CRS.',  # Centres (punctuated variant).
        'Dim',  # Dimension annotation.
        'CJ',  # Construction joint.
        'DJ',  # Double joist.
        'CSK',  # Countersunk hole / countersink preparation.
        'H&S',  # Health & Safety reference.
        'VB1',  # Internal layout code VB1.
    ],

    # B2: Elevation & Level References
    'Elevation & Level References': [
        'ExGL', # Existing ground level.
        'FFL',  # Finished floor level.
        'GRD',  # Ground.
        'H/L',  # High level.
        'L/L',  # Low level.
        'SSL',  # Structural slab level.
        'ToB',  # Top of base.
        'ToC',  # Top of concrete.
        'ToS',  # Top of steel.
        'ToW',  # Top of wall.
        'U/S',  # Underside.
    ],

    # B2: Exposure & Durability Classes (per BS EN 206-1)
    'Exposure & Durability Classes': [
        'UDS', # Ultimate durability state.
        'XC1', # Corrosion induced by carbonation.
        'XC2', # Wet, rarely dry.
        'XC3', # Wet, dry and moderate humidity.
        'XC4', # Dry or permanently dry.
        'XD1', # Corrosion induced by chlorides, low moisture.
        'XD2', # Corrosion induced by chlorides, moderate moisture.
        'XD3', # Corrosion induced by chlorides, high moisture.
        'XS1', # Corrosion induced by seawater spray.
        'XS2', # Corrosion induced by seawater spray and tidal action.
        'XS3', # Corrosion induced by seawater immersion.
        'XF1', # Freeze/thaw attack, moderate water saturation.
        'XF2', # Freeze/thaw attack with de-icing agents.
        'XF3', # Sea/ground water spray.
        'XF4', # Sea water immersion.
        'XA1', # Chemical attack, mild.
        'XA2', # Chemical attack, moderate.
        'XA3', # Chemical attack, severe.
    ],

    # B2: General Note Abbreviations
    'General Note Abbreviations': [
        'CCTV',    # Closed-circuit television (inspection/survey).
        'EGL',     # Existing ground level.
        'Typ',     # Typical (applies throughout unless noted).
        'uno',     # Unless noted otherwise.
        'u.n.o.',  # Unless noted otherwise (punctuated variant).
    ],

    # B2: General Notes Headings (including two-word & possessive)
    'General Notes': [
        'Architect',
        'BIM',
        'BRE Special Digest',
        'British',
        'Building'
        'Regulations',
        'Client',
        'Contractor',
        'Coordinator',
        'Designer',
        'Engineer',
        'EWP',
        'Eurocodes',
        'London',
        'Underground',
        'Manufacturer',
        'Rothoblaas',
        'Sub-contractor',


        'Annex A',
        'As Built',
        'Elliott Wood',
        'Special Digest',
        'Conbextra GP',
        'Temporary Works',
    ],

    # B2: Hole Sizes
    'Hole Sizes': [
        'H10',  # Hole for M10 bolt.
        'H12',  # Hole for M12 bolt.
        'H16',  # Hole for M16 bolt.
        'H20',  # Hole for M20 bolt.
        'H24',  # Hole for M24 bolt.
        'H30',  # Hole for M30 bolt.
        'H36',  # Hole for M36 bolt.
        'H42',  # Hole for M42 bolt.
    ],

    # B2: Load Cases & Limit States
    'Load Cases & Limit States': [
        'FLS',  # Fatigue limit state.
        'SLS',  # Serviceability limit state.
        'ULS',  # Ultimate limit state.
    ],

    # B2: Manufacturer & Product References
    'Manufacturer & Product References': [
        'Aecom',                      # Consultancy and infrastructure services.
        'UKPN',                       # Utility reference: UK Power Networks.
        'Ancon',                      # Structural fixings manufacturer.
        'MDC',                        # Variant: Ancon multi-drilled connections.
        'Armco',                      # Steel piling and metalwork manufacturer.
        'Fosroc',
        'Ltd',                        # Construction chemicals manufacturer.
        'Furfix',                     # Mechanical anchor and fixing manufacturer.
        'Conbextra GP Free Flow',     # High-performance free-flowing grout.
        'Galvafroid',                 # Corrosion protection product range.
        'Weber Five Star',            # Premium render & plaster product range.
        'London Clay',                # Bentonite slurry reference for drilling.
        'WLCA',                       # Water Leakage Control Agent.
        'M&E',                        # Mechanical & Electrical services reference.
    ],

    # B2: Material & Element Abbreviations
    'Material & Element Abbreviations': [
        'AAC',   # Autoclaved aerated concrete.
        'CC',    # Concrete cased.
        'CLT',   # Cross laminated timber.
        'DPC',   # Damp proof course.
        'DPM',   # Damp proof membrane.
        'GLT',   # Glue laminated timber.
        'LVL',   # Laminated veneer lumber.
        'MC',    # Mass concrete.
        'ms',    # Mild steel.
        'OSB',   # Oriented strand board.
        'PC',    # Precast concrete.
        'PCC',   # Precast concrete variant.
        'RAAC',  # Reinforced autoclaved aerated concrete.
        'RC',    # Reinforced concrete.
        'RMC',   # Ready mixed concrete.
        'S/W',   # Softwood.
        'TB',    # Thermal break.
        'WBP',   # Water boiler proof.
        'WRC',   # Water resistant concrete.
    ],

    # B2: Moments, Forces & Reactions
    'Moments, Forces & Reactions': [
        'MEd',      # Design bending moment.
        'NEd',      # Design axial force.
        'TEd',      # Design torsional moment.
        'VEd',      # Design shear force.
        'UDL',      # Uniformly distributed load.
        'LL',       # Live (imposed) load.
        'DL',       # Dead load.
        'Mx',       # Bending moment about local x‑axis.
        'My',       # Bending moment about local y‑axis.
        'Mz',       # Bending moment about local z‑axis.
        'Fh',       # Horizontal force component.
        'Fv',       # Vertical force component.
        'Fx',       # Force in global x‑direction.
        'V',        # Vertical shear force.
        'H',        # Horizontal shear force.
        'C',        # Compression force.
        'T',        # Tension force.
        'Tm',       # Torsional moment (non‑design value).
    ],

        # B2: Spans
        'Spans': [
        'L/250',    # Maximum span deflection ratio 1/250.
        'L/300',    # Maximum span deflection ratio 1/300.
        'L/360',    # Maximum span deflection ratio 1/360.
        'L/500',    # Maximum span deflection ratio 1/500.
    ],

    # B2: Quantity & Limits Abbreviations
    'Quantity & Limits Abbreviations': [
        'MAX',  # Maximum.
        'MIN',  # Minimum.
        'MJ',   # Movement joint.
        'R',    # Reaction.
        'thk',  # Thick.
    ],

    # B2: Rebar Steel Grade Names
    'Rebar Steel Grade Names': [
        'B500A', # Rebar steel grade B500A.
        'B500B', # Rebar steel grade B500B.
        'B500C', # Rebar steel grade B500C.
    ],

    # B2: Reinforcement Mesh
    'Reinforcement Mesh': [
        'A98',   # Mesh grade A98.
        'A142',  # Mesh grade A142.
        'A193',  # Mesh grade A193.
        'A252',  # Mesh grade A252.
        'A393',  # Mesh grade A393.
        'B1131', # Mesh grade B1131.
        'B196',  # Mesh grade B196.
        'B283',  # Mesh grade B283.
        'B385',  # Mesh grade B385.
        'B503',  # Mesh grade B503.
        'B785',  # Mesh grade B785.
        'C283',  # Mesh grade C283.
        'C385',  # Mesh grade C385.
        'C503',  # Mesh grade C503.
        'C636',  # Mesh grade C636.
        'C785',  # Mesh grade C785.
        'D49',   # Mesh grade D49.
        'D98',   # Mesh grade D98.
    ],

    # B2: Section Properties (per SCI Blue Book App A-3)
    'Section Properties': [
        'Iy',  # Second moment of area about y-axis.
        'Iz',  # Second moment of area about z-axis.
        'Wx',  # Section modulus about x-axis.
        'Wy',  # Plastic section modulus about y-axis.
        'Wz',  # Plastic section modulus about z-axis.
        'Zx',  # Section modulus about x-axis.
        'Zy',  # Section modulus about y-axis.
    ],

    # B2: Section Profiles & Steel Sections
    'Section Profiles & Steel Sections': [
        'CHS',       # Circular hollow section.
        'EA',        # Equal angle.
        'PFC',       # Parallel flange channel.
        'RHS',       # Rectangular hollow section.
        'SHS',       # Square hollow section.
        'T-section', # Structural tee.
        'UB',        # Universal beam.
        'UBP',       # Universal bearing pile.
        'UC',        # Universal column.
        'UA',        # Unequal angle.
        'VB',        # Vertical bracing.
        'WP',        # Wind post.
        'X',         # Cross bracing.
        'D',
    ],

    # B2: Steel Grade Suffixes (per BS EN 10210)
    'Steel Grade Suffixes': [
        'J0',     # Impact grade J0.
        'J2',     # Impact grade J2.
        'K2',     # Impact grade K2.
    ],

    # B2: Steel Grades
    'Steel Grades': [
        'S275',     # Steel grade S275.
        'S355',     # Steel grade S355.
        'S420',     # Steel grade S420.
        'S460',     # Steel grade S460.
        'S355JOH',  # Structural steel grade S355J0H: impact tested at 0°C, controlled-rolled.
    ],

    # B2: Standards & Codes
    'Standards & Codes': [
        'ASTM',                      # American Society for Testing and Materials.
        'BS',                        # British Standards.
        'DD',                        # Design data for masonry wall ties.
        'EN',                        # European Norms prefix.
        'ISO',                       # International Organization for Standardization.
        'NA',                        # National Annex.
        'PD',                        # BSI Published Document series (supplementary guidance).
        'LDSA',                      # Internal load-design standard.
        'BES',                       # Company-specific engineering standard.
    ],

    # B2: Timber Grades (per BS EN 338)
    'Timber Grades': [
        'C14',   # Softwood strength class C14.
        'C16',   # Softwood strength class C16.
        'C18',   # Softwood strength class C18.
        'C24',   # Softwood strength class C24.
        'C27',   # Softwood strength class C27.
        'C30',   # Softwood strength class C30.
        'GL24h', # Glue laminated timber grade GL24h.
        'GL28h', # Glue laminated timber grade GL28h.
        'GL32h', # Glue laminated timber grade GL32h.
        'CLT',   # Cross-laminated timber panel.
    ],

    # B2: Units & Measurements (per SCI “Blue Book” P363 Appendix A‐2 and RICS NRM2)
    'Units & Measurements': [
        'N',      # Newton (force).
        'N/m²',   # Newtons per square metre.
        'N/mm²',  # Newtons per square millimetre.
        'kN',     # Kilonewton (1000 N).
        'kN/m²',  # Kilonewtons per square metre.
        'kN/m',   # Kilonewtons per  metre.
        'kN/mm²', # Kilonewtons per square millimetre.
        'kNm',    # Kilonewton‐metre.
        'm',      # Metre.
        'm²',     # Square metre.
        'm³',     # Cubic metre.
        'mm',     # Millimetre.
        'mm²',    # Square millimetre.
        'mm³',    # Cubic millimetre.
        'Hz',     # Hertz (frequency).
    ],
}

# ────────────────────────── core helper class ─────────────────────────────
class ExceptionManager(object):
    """Load / merge default, system and project exception sets."""
    _cache = None

    # ------------- public helpers -----------------
    @classmethod
    def clear_cache(cls):
        cls._cache = None

    # ------------- internal IO --------------------
    @classmethod
    def _load_json(cls, path):
        try:
            with open(path, 'r') as fp:
                return json.load(fp)
        except Exception as exc:
            logger.error("JSON load error (%s): %s", path, exc, extra=LOG_KW)
            return {}

    @classmethod
    def _get_project_path(cls):
        """Return <model>_exceptions.json, creating an empty file if missing."""
        doc = __revit__.ActiveUIDocument.Document
        rvt = doc.PathName

        if rvt:
            proj_file = os.path.splitext(rvt)[0] + '_exceptions.json'
        else:
            from pyrevit import forms
            folder = forms.pick_folder(title="Select folder for project exceptions JSON")
            if not folder:
                raise Exception("No folder selected for project exceptions.")
            proj_file = os.path.join(folder, 'project_exceptions.json')

        # Guarantee the file exists and is a flat list
        if not os.path.isfile(proj_file):
            with open(proj_file, 'w') as fp:
                json.dump([], fp, indent=2)
        else:
            try:
                with open(proj_file, 'r') as fp:
                    content = json.load(fp)
                if not isinstance(content, list):
                    MessageBox.Show(
                        "The file\n{}\nexists but is not a flat list. "
                        "Please correct the JSON manually.".format(proj_file),
                        "Invalid project_exceptions.json",
                        MessageBoxButtons.OK)
                    logger.error(
                        "Invalid project exceptions format – expected list, got %s",
                        type(content).__name__, extra=LOG_KW)
            except Exception as exc:
                logger.error("Error reading project exceptions: %s", exc, extra=LOG_KW)

        return proj_file

    # ------------- merged data --------------------
    @classmethod
    def load_all(cls):
        """
        Merge default, system and project tiers into
        ``OrderedDict{category: [items...]}``.

        Result is cached; call :py:meth:`clear_cache` to force reload.
        """
        if cls._cache:
            return cls._cache

        merged = OrderedDict()

        # Defaults (tier-1)
        for cat, items in DEFAULT_EXCEPTIONS.items():
            merged[cat] = items[:]

        # System overrides (tier-2)
        for cat, items in cls._load_json(SYSTEM_FILE).items():
            if isinstance(items, list):
                merged.setdefault(cat, []).extend(items)

        # Project overrides (tier-3)
        project_path = cls._get_project_path()
        raw_project  = cls._load_json(project_path)
        project_items = raw_project if isinstance(raw_project, list) else []

        # Remove any lower-case duplicate across tiers, then append project items
        for phrase in project_items:
            low = phrase.lower()
            for lst in merged.values():
                lst[:] = [x for x in lst if x.lower() != low]

        for cat in merged:                         # de-duplicate per category
            seen = set()
            merged[cat] = [x for x in merged[cat]
                           if not (x.lower() in seen or seen.add(x.lower()))]

        cls._cache = {
            'merged' : merged,
            'project': project_items
        }
        return cls._cache

    # ------------- convenience views ------------
    @classmethod
    def get_single_literals_map(cls):
        """Return *{lowercase: canonical}* for all single-word literals."""
        data = cls.load_all()
        single = OrderedDict()

        # project tier wins ties
        for phrase in data['project']:
            if ' ' not in phrase:
                single[phrase.lower()] = phrase

        for phrases in data['merged'].values():
            for phrase in phrases:
                if ' ' not in phrase and phrase.lower() not in single:
                    single[phrase.lower()] = phrase
        return single

    @classmethod
    def get_multi_word_literals(cls):
        """Return multi-word literals, longest first."""
        data = cls.load_all()
        multi = []

        for phrase in data['project']:
            if ' ' in phrase:
                multi.append(phrase)

        for phrases in data['merged'].values():
            for phrase in phrases:
                if ' ' in phrase and phrase.lower() not in (p.lower() for p in multi):
                    multi.append(phrase)

        return sorted(multi, key=len, reverse=True)

    @classmethod
    def get_single_regexps(cls):
        """
        Build an ``OrderedDict{compiled_regex : canonical}`` for fast restoration.

        *   **Plain alphanum**   –  ``\\b … \\b`` boundaries
        *   **Special chars**    –  no word boundaries (handles ``M&E``, ``c/c`` …)
        *   **Multi-word**       –  exact phrase, case-insensitive
        *   **Patterns**         –  dotted abbr. & “2No.” style counts
        """
        out = OrderedDict()
        plain_map  = {}
        special_map = {}

        for low, canon in cls.get_single_literals_map().items():
            (plain_map if re.match(r'^[A-Za-z0-9]+$', low) else special_map)[low] = canon

        # plain words
        for low, canon in plain_map.items():
            out[re.compile(r'\b' + re.escape(low) + r'\b', re.I)] = canon

        # words containing &, /, dots, etc.
        for low, canon in special_map.items():
            out[re.compile(re.escape(low), re.I)] = canon

        # multi-word literals (longest first for greedy matching)
        for phrase in cls.get_multi_word_literals():
            out[re.compile(re.escape(phrase), re.I)] = phrase

        # special patterns
        out[re.compile(r'\b(\d+)\s*No\.?', re.I)] = \
            (lambda m: u'%sNo%s' % (m.group(1), '.' if m.group(0).endswith('.') else ''))
        out[re.compile(r'\bu\.n\.o\.', re.I)] = 'u.n.o.'
        out[re.compile(r'\bt\.b\.c\.', re.I)] = 'T.B.C.'

        return out

    # Possessive (apostrophe) terms – stored only in system JSON
    @classmethod
    def load_apostrophe_exceptions(cls):
        return cls._load_json(SYSTEM_FILE).get("Possessive Terms", {})

# ───────────────────── small runtime helper (used elsewhere) ──────────────
class ExceptionApplier(object):
    """Utility class used by Change Register to enforce per-token rules."""
    def __init__(self):
        self.apostrophe_map = ExceptionManager.load_apostrophe_exceptions()

    def apply_apostrophe_exceptions(self, text):
        if not self.apostrophe_map:
            return text
        pattern = re.compile(r'\b({})\b'.format('|'.join(map(re.escape,
                                                            self.apostrophe_map.keys()))),
                             re.IGNORECASE)
        return pattern.sub(lambda m: self.apostrophe_map.get(m.group(0).lower(),
                                                             m.group(0)),
                           text)

# ───────────────────── cache cleanup on pyRevit exit ───────────────────────
atexit.register(ExceptionManager.clear_cache)