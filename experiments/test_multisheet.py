"""
Test multi-sheet drawing: verifies that circles are drawn on separate sheets.

This script creates a 3-sheet drawing in KOMPAS-3D with a single circle pair
on each sheet, at different positions to make it easy to visually verify
which sheet each circle lands on.

Usage:
    1. Start KOMPAS-3D
    2. Run: python experiments/test_multisheet.py
    3. Check each sheet in KOMPAS-3D: each should have exactly one circle pair
       at a different position.

Expected result:
    Sheet 1: circle pair at (105, 150) - center of A4 portrait
    Sheet 2: circle pair at (50, 50) - bottom-left area
    Sheet 3: circle pair at (160, 250) - top-right area
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import logging
from kompas_random_circles import (
    connect_to_kompas,
    create_drawing_document,
    draw_coaxial_circles,
    add_new_sheet,
)

# Enable debug logging to see view switching details
logger = logging.getLogger("kompas_random_circles")
logger.setLevel(logging.DEBUG)


def main():
    print("Connecting to KOMPAS-3D...")
    kompas_object, api5_module, constants, app7, api7_module = connect_to_kompas()

    print("Creating A4 drawing (3 sheets)...")
    iDocument2D = create_drawing_document(
        kompas_object, api5_module, constants,
        sheet_format=4, landscape=False, no_frame=True,
    )

    outer_r = 20.0
    inner_r = 12.0

    # Sheet 1: draw at center
    print("Drawing on sheet 1...")
    draw_coaxial_circles(iDocument2D, [(105, 150)], outer_r, inner_r)

    # Sheet 2: add sheet and draw at bottom-left
    print("Adding sheet 2 and drawing...")
    add_new_sheet(
        iDocument2D, kompas_object, api5_module, constants,
        sheet_format=4, landscape=False,
        app7=app7, api7_module=api7_module,
    )
    draw_coaxial_circles(iDocument2D, [(50, 50)], outer_r, inner_r)

    # Sheet 3: add sheet and draw at top-right
    print("Adding sheet 3 and drawing...")
    add_new_sheet(
        iDocument2D, kompas_object, api5_module, constants,
        sheet_format=4, landscape=False,
        app7=app7, api7_module=api7_module,
    )
    draw_coaxial_circles(iDocument2D, [(160, 250)], outer_r, inner_r)

    print("Done! Check each sheet in KOMPAS-3D:")
    print("  Sheet 1 should have a circle at center (105, 150)")
    print("  Sheet 2 should have a circle at bottom-left (50, 50)")
    print("  Sheet 3 should have a circle at top-right (160, 250)")


if __name__ == "__main__":
    main()
