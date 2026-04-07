"""
Simple test: draw a single coaxial circle pair in KOMPAS-3D.

This script connects to a running KOMPAS-3D instance and draws one pair
of concentric circles with black fill on an A4 sheet without a frame.

Usage:
    1. Start KOMPAS-3D
    2. Run: python examples/simple_circle_test.py
"""

import sys
import os

# Add parent directory to path so we can import the main module
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from kompas_random_circles import (
    connect_to_kompas,
    create_drawing_document,
    draw_coaxial_circles,
)


def main():
    print("Connecting to KOMPAS-3D...")
    kompas_object, api5_module, constants, app7, api7_module = connect_to_kompas()

    print("Creating A4 drawing without frame...")
    iDocument2D = create_drawing_document(
        kompas_object, api5_module, constants,
        sheet_format=4,  # A4
        landscape=False,
        no_frame=True,
    )

    # Draw a single ring at the center of the A4 sheet (210x297mm portrait)
    center_x = 105.0  # half of 210mm
    center_y = 148.5  # half of 297mm
    outer_radius = 25.0
    inner_radius = 15.0

    positions = [(center_x, center_y)]

    print(f"Drawing coaxial circle pair at ({center_x}, {center_y})...")
    print(f"  Outer radius: {outer_radius} mm")
    print(f"  Inner radius: {inner_radius} mm")

    draw_coaxial_circles(
        iDocument2D, kompas_object, api5_module, constants,
        positions, outer_radius, inner_radius,
    )

    print("Done! Check the KOMPAS-3D window.")


if __name__ == "__main__":
    main()
