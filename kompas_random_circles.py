"""
KOMPAS-3D Random Coaxial Circles Drawing Application.

This application connects to KOMPAS-3D via COM automation and draws
randomly placed pairs of coaxial circles (concentric rings with black fill)
on a drawing sheet. Settings are configured via a tkinter GUI.

Requirements:
    - Windows OS with KOMPAS-3D installed (v16+ recommended)
    - Python 3.7+
    - pywin32 package (pip install pywin32)

Usage:
    python kompas_random_circles.py
"""

import logging
import math
import random
import sys
import tkinter as tk
from tkinter import messagebox, ttk

# Logging configuration (off by default, enable with --debug flag)
logger = logging.getLogger(__name__)
logger.setLevel(logging.WARNING)

_handler = logging.StreamHandler()
_handler.setFormatter(logging.Formatter(
    "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
))
logger.addHandler(_handler)


# ---------------------------------------------------------------------------
# KOMPAS-3D COM connection helpers
# ---------------------------------------------------------------------------

# TLB GUIDs for KOMPAS-3D SDK
_KOMPAS_CONSTANTS_TLB = "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}"
_KOMPAS_API5_TLB = "{0422828C-F174-495E-AC5D-D31014DBBE87}"
_KOMPAS_API7_TLB = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"

# Sheet format constants (index -> (width_mm, height_mm) in landscape)
SHEET_FORMATS = {
    0: ("A0", 1189, 841),
    1: ("A1", 841, 594),
    2: ("A2", 594, 420),
    3: ("A3", 420, 297),
    4: ("A4", 297, 210),
}


def connect_to_kompas():
    """Connect to KOMPAS-3D application via COM automation.

    Returns a tuple (kompas_object, kompas6_api5_module, kompas6_constants,
    app7, kompas_api7_module) for use with both API5 and API7.

    Raises RuntimeError if KOMPAS-3D is not available.
    """
    try:
        import pythoncom
        from win32com.client import Dispatch, gencache
    except ImportError:
        raise RuntimeError(
            "pywin32 is not installed. Run: pip install pywin32"
        )

    logger.info("Loading KOMPAS-3D type libraries...")

    try:
        kompas6_constants = gencache.EnsureModule(
            _KOMPAS_CONSTANTS_TLB, 0, 1, 0
        ).constants
    except Exception as exc:
        raise RuntimeError(
            f"Cannot load KOMPAS constants TLB. Is KOMPAS-3D installed? {exc}"
        )

    try:
        kompas6_api5_module = gencache.EnsureModule(
            _KOMPAS_API5_TLB, 0, 1, 0
        )
    except Exception as exc:
        raise RuntimeError(f"Cannot load KOMPAS API5 TLB: {exc}")

    try:
        kompas_api7_module = gencache.EnsureModule(
            _KOMPAS_API7_TLB, 0, 1, 0
        )
    except Exception as exc:
        raise RuntimeError(f"Cannot load KOMPAS API7 TLB: {exc}")

    logger.info("Connecting to KOMPAS-3D application...")

    try:
        kompas_object = Dispatch("KOMPAS.Application.5")
    except Exception as exc:
        raise RuntimeError(
            f"Cannot connect to KOMPAS-3D. Is the application running? {exc}"
        )

    try:
        api7 = kompas_api7_module.IKompasAPIObject(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
                kompas_api7_module.IKompasAPIObject.CLSID,
                pythoncom.IID_IDispatch,
            )
        )
        app7 = api7.Application
        app7.Visible = True
    except Exception as exc:
        raise RuntimeError(f"Cannot connect to KOMPAS API7: {exc}")

    logger.info("Successfully connected to KOMPAS-3D.")
    return (
        kompas_object,
        kompas6_api5_module,
        kompas6_constants,
        app7,
        kompas_api7_module,
    )


# ---------------------------------------------------------------------------
# Drawing logic
# ---------------------------------------------------------------------------

def create_drawing_document(kompas_object, api5_module, constants,
                            sheet_format=4, landscape=False, no_frame=True):
    """Create a new 2D drawing document in KOMPAS-3D.

    Args:
        kompas_object: KOMPAS API5 application object.
        api5_module: API5 module (gencache result).
        constants: KOMPAS constants module.
        sheet_format: Sheet format index (0=A0 .. 4=A4).
        landscape: True for landscape orientation.
        no_frame: If True, create sheet without title block / border.

    Returns:
        iDocument2D interface for the created document.
    """
    logger.info(
        "Creating drawing: format=%s, landscape=%s, no_frame=%s",
        SHEET_FORMATS.get(sheet_format, ("?",))[0], landscape, no_frame,
    )

    doc_param = api5_module.ksDocumentParam(
        kompas_object.GetParamStruct(constants.ko_DocumentParam)
    )
    doc_param.Init()
    doc_param.type = 1  # lt_DocSheetStandart (standard drawing)

    sheet_par = doc_param.GetLayoutParam()
    sheet_par.Init()
    # layoutName: name of the layout/style library.
    # An empty string means "use the default graphic.lyt library".
    # Per the KOMPAS SDK docs, passing the full path causes malfunctions;
    # passing an empty string is the correct way to use the default library.
    sheet_par.layoutName = ""
    # shtType: selects the layout (оформление) from the library by its
    # "Номер" (number) column value in graphic.lyt.
    # Common values (from KOMPAS graphic.lyt):
    #   1  – "Чертеж констр. Первый лист. ГОСТ 2.104-2006" (standard 1st sheet)
    #   2  – "Чертеж констр. Последующие листы. ГОСТ 2.104-2006"
    #   13 – "Без внутренней рамки" (without inner frame / title block)
    # When no_frame=True we want shtType=13 ("без внутренней рамки").
    # Setting shtType=0 produces an undefined/default style — use 13 instead.
    if no_frame:
        sheet_par.shtType = 13  # "Без внутренней рамки" (no inner frame)
    else:
        sheet_par.shtType = 1   # Standard first-sheet layout with title block

    standart_sheet = sheet_par.GetSheetParam()
    standart_sheet.format = sheet_format
    standart_sheet.multiply = 1
    # For A4 portrait: direct=False; landscape=True means direct=True
    standart_sheet.direct = landscape

    iDocument2D = kompas_object.Document2D
    result = iDocument2D.ksCreateDocument(doc_param)
    if not result:
        raise RuntimeError("Failed to create drawing document in KOMPAS-3D.")

    kompas_object.Visible = True
    logger.info("Drawing document created successfully.")
    return iDocument2D


def get_drawing_area(sheet_format, landscape=False, margin=10):
    """Calculate the usable drawing area for random placement.

    Args:
        sheet_format: Sheet format index (0=A0 .. 4=A4).
        landscape: True for landscape orientation.
        margin: Margin in mm from sheet edges.

    Returns:
        Tuple (x_min, y_min, x_max, y_max) in mm.
    """
    _, w, h = SHEET_FORMATS[sheet_format]
    if not landscape:
        w, h = h, w  # portrait: swap width and height

    x_min = margin
    y_min = margin
    x_max = w - margin
    y_max = h - margin

    logger.debug(
        "Drawing area: (%.1f, %.1f) - (%.1f, %.1f)", x_min, y_min, x_max, y_max
    )
    return x_min, y_min, x_max, y_max


def generate_circle_positions(count, area, outer_radius, min_distance,
                              max_distance, max_attempts=1000):
    """Generate random positions for pairs of coaxial circles.

    Each pair consists of two circles: the pair has a center, and the two
    circles in the pair are coaxial (share the same center). The pairs
    themselves are placed randomly with a distance constraint between their
    centers.

    Args:
        count: Number of circle pairs to place.
        area: Tuple (x_min, y_min, x_max, y_max).
        outer_radius: Outer radius of each circle pair in mm.
        min_distance: Minimum distance between pair centers in mm.
        max_distance: Maximum distance between pair centers in mm.
        max_attempts: Maximum placement attempts before giving up.

    Returns:
        List of (x, y) tuples for circle pair centers.
    """
    x_min, y_min, x_max, y_max = area
    positions = []

    # Ensure circles fit within the drawing area
    effective_x_min = x_min + outer_radius
    effective_y_min = y_min + outer_radius
    effective_x_max = x_max - outer_radius
    effective_y_max = y_max - outer_radius

    if effective_x_min >= effective_x_max or effective_y_min >= effective_y_max:
        logger.warning("Drawing area too small for circles with radius %.1f", outer_radius)
        return positions

    for i in range(count):
        placed = False
        for attempt in range(max_attempts):
            x = random.uniform(effective_x_min, effective_x_max)
            y = random.uniform(effective_y_min, effective_y_max)

            # Check distance constraints with already placed circles
            valid = True
            for px, py in positions:
                dist = math.hypot(x - px, y - py)
                if dist < min_distance:
                    valid = False
                    break
                # We also want pairs not to be farther than max_distance
                # from at least one neighbor (soft constraint for clustering)

            if valid:
                positions.append((x, y))
                placed = True
                logger.debug(
                    "Placed circle pair %d at (%.1f, %.1f) on attempt %d",
                    i + 1, x, y, attempt + 1,
                )
                break

        if not placed:
            logger.warning(
                "Could not place circle pair %d after %d attempts. "
                "Consider reducing count or increasing drawing area.",
                i + 1, max_attempts,
            )

    return positions


def draw_coaxial_circles(iDocument2D, positions, outer_radius, inner_radius):
    """Draw coaxial circle pairs with black fill at the given positions.

    Each pair is two concentric circles with solid black fill (ksColouring)
    between them.

    Args:
        iDocument2D: KOMPAS 2D document interface.
        positions: List of (x, y) center coordinates.
        outer_radius: Outer circle radius in mm.
        inner_radius: Inner circle radius in mm.
    """
    logger.info(
        "Drawing %d circle pairs (R_outer=%.1f, R_inner=%.1f)...",
        len(positions), outer_radius, inner_radius,
    )

    for idx, (cx, cy) in enumerate(positions):
        logger.debug("Drawing pair %d at (%.1f, %.1f)", idx + 1, cx, cy)

        # Draw visible circles (outer and inner)
        iDocument2D.ksCircle(cx, cy, outer_radius, 1)
        iDocument2D.ksCircle(cx, cy, inner_radius, 1)

        # Create solid black fill between the two circles.
        #
        # KOMPAS-3D Automation (COM) API5 notes:
        #
        # 1. ksHatch Automation signature:
        #      long ksHatch(long style, double ang, double step,
        #                   double width, double x0, double y0)
        #    It takes INDIVIDUAL scalar arguments, NOT a struct.
        #    The struct-based version (ksHatchParam) is only available in
        #    the native SDK (C++) interface, not in the COM Automation layer.
        #    The Automation equivalent is ksHatchByParam, but that is not
        #    exposed in the gencache binding either.
        #
        # 2. ksHatch boundary convention:
        #    - Call ksHatch() first to start the hatch object
        #    - Then draw boundary primitives (circles, arcs, etc.)
        #    - Call ksEndObj() last to finalise; it returns a reference to
        #      the created hatch object
        #    There is no separate ksNewGroup/ksContour/ksEndGroup needed;
        #    the primitives drawn between ksHatch() and ksEndObj() become
        #    the boundary automatically.
        #
        # 3. System hatch styles (from KOMPAS SDK hstyles table):
        #      0=Metal, 1=Non-metal, 2=Wood, 3=Stone, 4=Ceramics,
        #      5=Concrete, 6=Glass, 7=Liquid, 8=Natural soil, 9=Fill soil,
        #      10=Artificial stone, 11=Reinforced concrete,
        #      12=Stressed RC, 13=Wood (longitudinal), 14=Sand
        #    None of these is "solid fill". For an opaque solid fill use
        #    ksColouring(color) instead (a separate API).
        #
        # 4. ksColouring Automation signature:
        #      long ksColouring(long color)
        #    Then draw boundary primitives, then ksEndObj() finalises.
        #    color = 0 → black (BGR 0x000000).
        #
        # Strategy: use ksColouring(0) for a solid black fill, because
        # ksHatch only supports patterned material styles, not solid fill.

        # Start solid-fill object (colour 0 = black in KOMPAS BGR palette)
        iDocument2D.ksColouring(0)

        # Boundary primitives: outer circle (exterior boundary) and inner
        # circle (hole), drawn in any order between ksColouring and ksEndObj.
        iDocument2D.ksCircle(cx, cy, outer_radius, 1)
        iDocument2D.ksCircle(cx, cy, inner_radius, 1)

        # Finalise the fill object
        iDocument2D.ksEndObj()

    logger.info("All circle pairs drawn successfully.")


def add_new_sheet(iDocument2D, kompas_object, api5_module, constants,
                  sheet_format=4, landscape=False):
    """Add a new sheet to the current document.

    Uses ksNewSheet method from the API5 to add a page.

    Args:
        iDocument2D: KOMPAS 2D document interface.
        kompas_object: KOMPAS API5 application object.
        api5_module: API5 module.
        constants: KOMPAS constants module.
        sheet_format: Sheet format index.
        landscape: Orientation flag.
    """
    logger.info("Adding new sheet...")

    # Build sheet parameters using the same pattern as create_drawing_document.
    # There is no ko_SheetParam constant in the KOMPAS API — the ksSheetPar
    # interface must be obtained via ksDocumentParam.GetLayoutParam().
    doc_param = api5_module.ksDocumentParam(
        kompas_object.GetParamStruct(constants.ko_DocumentParam)
    )
    doc_param.Init()
    doc_param.type = 1  # lt_DocSheetStandart

    sheet_par = doc_param.GetLayoutParam()
    sheet_par.Init()
    # layoutName = "" → use default graphic.lyt library (correct convention)
    sheet_par.layoutName = ""
    # shtType = 13 → "Без внутренней рамки" (without inner frame / title block)
    sheet_par.shtType = 13

    standart_sheet = sheet_par.GetSheetParam()
    standart_sheet.format = sheet_format
    standart_sheet.multiply = 1
    standart_sheet.direct = landscape

    result = iDocument2D.ksNewSheet(sheet_par)
    if not result:
        logger.warning("ksNewSheet returned False, trying alternative method...")
        # Alternative: use ksInsertSheet
        result = iDocument2D.ksInsertSheet()

    logger.info("New sheet added: %s", result)
    return result


# ---------------------------------------------------------------------------
# Main workflow: draw on one or more sheets
# ---------------------------------------------------------------------------

def run_drawing(settings):
    """Main drawing workflow.

    Args:
        settings: Dictionary with user settings from the GUI:
            - count: number of circle pairs per sheet
            - outer_diameter: outer circle diameter in mm
            - inner_diameter: inner circle diameter in mm
            - min_distance: minimum distance between centers in mm
            - max_distance: maximum distance between centers in mm
            - sheet_format: sheet format index (0-4)
            - landscape: bool
            - num_sheets: number of sheets to create
    """
    logger.info("Starting drawing with settings: %s", settings)

    # Connect to KOMPAS-3D
    (kompas_object, api5_module, constants,
     app7, api7_module) = connect_to_kompas()

    outer_radius = settings["outer_diameter"] / 2.0
    inner_radius = settings["inner_diameter"] / 2.0

    if inner_radius >= outer_radius:
        raise ValueError("Inner diameter must be smaller than outer diameter.")

    num_sheets = settings["num_sheets"]
    sheet_format = settings["sheet_format"]
    landscape = settings["landscape"]

    # Create initial document (first sheet)
    iDocument2D = create_drawing_document(
        kompas_object, api5_module, constants,
        sheet_format=sheet_format,
        landscape=landscape,
        no_frame=True,
    )

    for sheet_idx in range(num_sheets):
        if sheet_idx > 0:
            # Add a new sheet for subsequent pages
            add_new_sheet(
                iDocument2D, kompas_object, api5_module, constants,
                sheet_format=sheet_format, landscape=landscape,
            )

        # Calculate drawing area
        area = get_drawing_area(sheet_format, landscape)

        # Generate random positions
        positions = generate_circle_positions(
            count=settings["count"],
            area=area,
            outer_radius=outer_radius,
            min_distance=settings["min_distance"],
            max_distance=settings["max_distance"],
        )

        if not positions:
            logger.warning("No positions generated for sheet %d.", sheet_idx + 1)
            continue

        # Draw the circles
        draw_coaxial_circles(
            iDocument2D, positions, outer_radius, inner_radius,
        )

        logger.info("Sheet %d completed with %d circle pairs.",
                     sheet_idx + 1, len(positions))

    logger.info("Drawing completed: %d sheet(s).", num_sheets)
    return True


# ---------------------------------------------------------------------------
# GUI (tkinter settings window)
# ---------------------------------------------------------------------------

class SettingsWindow:
    """Tkinter-based settings window for configuring the drawing parameters."""

    # Default values
    DEFAULTS = {
        "count": 10,
        "outer_diameter": 40,
        "inner_diameter": 20,
        "min_distance": 50,
        "max_distance": 80,
        "sheet_format": 4,  # A4
        "landscape": False,
        "num_sheets": 1,
    }

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("KOMPAS-3D: Random Coaxial Circles")
        self.root.resizable(False, False)

        self.result = None
        self._build_ui()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Title
        ttk.Label(
            main_frame,
            text="Settings for Random Coaxial Circles",
            font=("Segoe UI", 12, "bold"),
        ).grid(row=0, column=0, columnspan=2, pady=(0, 10))

        row = 1

        # Number of circle pairs
        ttk.Label(main_frame, text="Number of circle pairs:").grid(
            row=row, column=0, sticky="w", pady=3
        )
        self.count_var = tk.IntVar(value=self.DEFAULTS["count"])
        ttk.Spinbox(
            main_frame, from_=1, to=200, textvariable=self.count_var, width=10
        ).grid(row=row, column=1, sticky="e", pady=3)
        row += 1

        # Outer diameter
        ttk.Label(main_frame, text="Outer circle diameter (mm):").grid(
            row=row, column=0, sticky="w", pady=3
        )
        self.outer_diam_var = tk.DoubleVar(value=self.DEFAULTS["outer_diameter"])
        ttk.Spinbox(
            main_frame, from_=5, to=500, textvariable=self.outer_diam_var,
            width=10, increment=5,
        ).grid(row=row, column=1, sticky="e", pady=3)
        row += 1

        # Inner diameter
        ttk.Label(main_frame, text="Inner circle diameter (mm):").grid(
            row=row, column=0, sticky="w", pady=3
        )
        self.inner_diam_var = tk.DoubleVar(value=self.DEFAULTS["inner_diameter"])
        ttk.Spinbox(
            main_frame, from_=1, to=499, textvariable=self.inner_diam_var,
            width=10, increment=5,
        ).grid(row=row, column=1, sticky="e", pady=3)
        row += 1

        # Min distance between centers
        ttk.Label(main_frame, text="Min distance between centers (mm):").grid(
            row=row, column=0, sticky="w", pady=3
        )
        self.min_dist_var = tk.DoubleVar(value=self.DEFAULTS["min_distance"])
        ttk.Spinbox(
            main_frame, from_=10, to=1000, textvariable=self.min_dist_var,
            width=10, increment=5,
        ).grid(row=row, column=1, sticky="e", pady=3)
        row += 1

        # Max distance between centers
        ttk.Label(main_frame, text="Max distance between centers (mm):").grid(
            row=row, column=0, sticky="w", pady=3
        )
        self.max_dist_var = tk.DoubleVar(value=self.DEFAULTS["max_distance"])
        ttk.Spinbox(
            main_frame, from_=10, to=2000, textvariable=self.max_dist_var,
            width=10, increment=5,
        ).grid(row=row, column=1, sticky="e", pady=3)
        row += 1

        # Sheet format
        ttk.Label(main_frame, text="Sheet format:").grid(
            row=row, column=0, sticky="w", pady=3
        )
        self.format_var = tk.StringVar(value="A4")
        format_combo = ttk.Combobox(
            main_frame, textvariable=self.format_var,
            values=["A0", "A1", "A2", "A3", "A4"],
            state="readonly", width=8,
        )
        format_combo.grid(row=row, column=1, sticky="e", pady=3)
        row += 1

        # Orientation
        ttk.Label(main_frame, text="Orientation:").grid(
            row=row, column=0, sticky="w", pady=3
        )
        self.landscape_var = tk.BooleanVar(value=self.DEFAULTS["landscape"])
        orient_frame = ttk.Frame(main_frame)
        orient_frame.grid(row=row, column=1, sticky="e", pady=3)
        ttk.Radiobutton(
            orient_frame, text="Portrait", variable=self.landscape_var,
            value=False,
        ).pack(side="left")
        ttk.Radiobutton(
            orient_frame, text="Landscape", variable=self.landscape_var,
            value=True,
        ).pack(side="left")
        row += 1

        # Number of sheets
        ttk.Label(main_frame, text="Number of sheets:").grid(
            row=row, column=0, sticky="w", pady=3
        )
        self.sheets_var = tk.IntVar(value=self.DEFAULTS["num_sheets"])
        ttk.Spinbox(
            main_frame, from_=1, to=50, textvariable=self.sheets_var, width=10
        ).grid(row=row, column=1, sticky="e", pady=3)
        row += 1

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=(15, 0))

        ttk.Button(btn_frame, text="Draw", command=self._on_draw).pack(
            side="left", padx=5
        )
        ttk.Button(btn_frame, text="Cancel", command=self._on_cancel).pack(
            side="left", padx=5
        )

    def _format_name_to_index(self, name):
        """Convert format name (e.g. 'A3') to KOMPAS index."""
        mapping = {"A0": 0, "A1": 1, "A2": 2, "A3": 3, "A4": 4}
        return mapping.get(name, 4)

    def _validate(self):
        """Validate user inputs."""
        errors = []

        if self.count_var.get() < 1:
            errors.append("Number of circle pairs must be at least 1.")

        if self.outer_diam_var.get() <= self.inner_diam_var.get():
            errors.append("Outer diameter must be greater than inner diameter.")

        if self.min_dist_var.get() > self.max_dist_var.get():
            errors.append(
                "Minimum distance must not exceed maximum distance."
            )

        if self.min_dist_var.get() < self.outer_diam_var.get():
            errors.append(
                "Minimum distance between centers should be at least "
                "equal to the outer diameter to prevent overlapping."
            )

        if errors:
            messagebox.showerror("Validation Error", "\n".join(errors))
            return False
        return True

    def _on_draw(self):
        if not self._validate():
            return

        self.result = {
            "count": self.count_var.get(),
            "outer_diameter": self.outer_diam_var.get(),
            "inner_diameter": self.inner_diam_var.get(),
            "min_distance": self.min_dist_var.get(),
            "max_distance": self.max_dist_var.get(),
            "sheet_format": self._format_name_to_index(self.format_var.get()),
            "landscape": self.landscape_var.get(),
            "num_sheets": self.sheets_var.get(),
        }
        self.root.destroy()

    def _on_cancel(self):
        self.result = None
        self.root.destroy()

    def show(self):
        """Display the settings window and wait for user action.

        Returns:
            Dictionary of settings if user clicked Draw, None if cancelled.
        """
        self.root.mainloop()
        return self.result


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    # Enable debug logging if requested
    if "--debug" in sys.argv:
        logger.setLevel(logging.DEBUG)
        logger.info("Debug logging enabled.")

    # Show settings window
    window = SettingsWindow()
    settings = window.show()

    if settings is None:
        print("Operation cancelled by user.")
        return

    print(f"Starting drawing with {settings['count']} circle pairs "
          f"on {settings['num_sheets']} sheet(s)...")

    try:
        run_drawing(settings)
        print("Drawing completed successfully!")
        messagebox.showinfo("Success", "Drawing completed successfully!")
    except RuntimeError as exc:
        print(f"Error: {exc}")
        messagebox.showerror("Error", str(exc))
    except Exception as exc:
        logger.exception("Unexpected error")
        print(f"Unexpected error: {exc}")
        messagebox.showerror("Unexpected Error", str(exc))


if __name__ == "__main__":
    main()
