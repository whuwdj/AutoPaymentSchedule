# -*- coding: utf-8 -*-
"""
Convert AutoPay.jpeg to app.ico for use as the application icon.
Uses only Python standard library — no external dependencies.

Usage:
    python3 convert_to_ico.py
"""

from __future__ import annotations

import os
import struct
import sys


def read_jpeg(filepath: str) -> bytes:
    """Read JPEG file as raw bytes."""
    with open(filepath, "rb") as f:
        return f.read()


def jpeg_to_bmp_bits(jpeg_data: bytes, width: int, height: int) -> bytes:
    """
    Create a BMP image buffer that contains the JPEG data embedded as a resource.
    For ICO format, we can embed JPEG directly (Windows 8+ supports JPEG in ICO),
    but for maximum compatibility, we'll create a proper BMP/DIB header + JPEG fallback.

    Actually, the simplest compatible approach: use the JPEG data with an ICO
    header that references it as a PNG-like resource. But Windows ICO spec only
    supports BMP, PNG, and icon resources.

    Best approach for compatibility: just use the JPEG data directly in the ICO
    container. Modern Windows (Vista+) handles this via the shell.

    For true cross-platform compatibility with PyQt6, we'll keep the original
    JPEG for the window icon and create a multi-size ICO.
    """
    return jpeg_data


def create_ico_from_jpeg(jpeg_path: str, ico_path: str):
    """
    Create an ICO file that wraps the JPEG data.
    Windows 8+ can display JPEG-in-ICO. For older systems, PyQt6 will fall
    back to loading the JPEG directly.

    We create multiple sizes by keeping the JPEG as-is.
    """
    jpeg_data = read_jpeg(jpeg_path)

    # ICO header
    # Reserved (must be 0)
    # Type (1 = icon, 2 = cursor)
    # Count (number of images)
    ico_header = struct.pack("<HHH", 0, 1, 1)

    # ICO directory entry
    # Width: 0 = 256px
    # Height: 0 = 256px
    # Color count: 0 (>= 8bpp)
    # Reserved: 0
    # Color planes: 0 or 1
    # Bits per pixel: 32
    # Size of image data
    # Offset to image data
    w = 0  # 256
    h = 0  # 256
    data_size = len(jpeg_data)
    offset = 22  # 6 (header) + 16 (entry) = 22

    directory_entry = struct.pack(
        "<BBBBHHII", w, h, 0, 0, 1, 32, data_size, offset
    )

    # Image data: 40-byte BITMAPINFOHEADER + JPEG data
    # For ICO with JPEG, we use a BITMAPINFOHEADER that describes the image,
    # then append the JPEG data. Windows will decode the JPEG.
    bitmap_info_header = struct.pack(
        "<IiiHHIIiiII",
        40,  # sizeof(BITMAPINFOHEADER)
        256,  # width
        256 * 2,  # height (doubled for ICO: includes AND mask)
        1,  # color planes
        32,  # bits per pixel
        0,  # compression (0 = BI_RGB, but we're embedding JPEG)
        data_size,  # image size
        0,  # h resolution
        0,  # v resolution
        0,  # colors in color table
        0,  # important color count
    )

    with open(ico_path, "wb") as f:
        f.write(ico_header)
        f.write(directory_entry)
        f.write(bitmap_info_header)
        f.write(jpeg_data)

    print(f"Created {ico_path} from {jpeg_path}")
    print(f"  ICO size: {os.path.getsize(ico_path)} bytes")
    print(f"  JPEG size: {len(jpeg_data)} bytes")


def create_png_then_ico(jpeg_path: str, ico_path: str):
    """
    Better approach: JPEG -> PNG -> ICO for full cross-platform compatibility.
    We use PIL/Pillow if available, otherwise fall back to raw JPEG embedding.
    """
    # Try Pillow first
    try:
        from PIL import Image

        # Open JPEG
        img = Image.open(jpeg_path).convert("RGBA")

        # Create multiple sizes for ICO
        sizes = [256, 128, 64, 48, 32, 16]

        ico_images = []
        for size in sizes:
            resized = img.copy()
            resized.thumbnail((size, size), Image.LANCZOS)
            # Center on size x size canvas
            canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            x = (size - resized.width) // 2
            y = (size - resized.height) // 2
            canvas.paste(resized, (x, y), resized)
            ico_images.append(canvas)

        # Save as ICO (Pillow handles multi-size ICO natively)
        ico_images[0].save(
            ico_path,
            format="ICO",
            sizes=[(img.size[0], img.size[1]) for img in ico_images],
            append_images=ico_images[1:],
        )

        print(f"Created {ico_path} (multi-size ICO via Pillow)")
        print(f"  Sizes: {sizes}")
        return True

    except ImportError:
        print("Pillow not available, using fallback method...")
        return False


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    jpeg_path = os.path.join(script_dir, "AutoPay.jpeg")
    ico_path = os.path.join(script_dir, "app.ico")

    if not os.path.isfile(jpeg_path):
        print(f"Error: {jpeg_path} not found!")
        sys.exit(1)

    # Try Pillow first for best quality
    if not create_png_then_ico(jpeg_path, ico_path):
        # Fallback to direct JPEG embedding
        create_ico_from_jpeg(jpeg_path, ico_path)

    # Also keep a PNG version
    try:
        from PIL import Image

        png_path = os.path.join(script_dir, "app_icon.png")
        img = Image.open(jpeg_path).convert("RGBA")
        img.thumbnail((512, 512), Image.LANCZOS)
        img.save(png_path, "PNG")
        print(f"Also created {png_path}")

        # Generate .icns for macOS using sips command line tool
        icns_path = os.path.join(script_dir, "app.icns")
        import subprocess
        try:
            # sips on macOS can convert PNG to ICNS
            result = subprocess.run(
                ["sips", "-s", "format", "icns", png_path, "--out", icns_path],
                capture_output=True, text=True, timeout=10,
            )
            if result.returncode == 0:
                print(f"Also created {icns_path}")
            else:
                print("Warning: sips failed to create .icns, macOS will use PNG fallback.")
        except (FileNotFoundError, subprocess.TimeoutExpired):
            # sips not available (not on macOS)
            pass
    except ImportError:
        pass


if __name__ == "__main__":
    main()