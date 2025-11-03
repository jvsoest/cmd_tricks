"""
QR Code Generator
Generates a QR code image from a given URL or text string.
"""

import argparse
import qrcode
import sys
from pathlib import Path


def generate_qr_code(data, output_path=None, size=10, border=4):
    """
    Generate a QR code from the given data.
    
    Args:
        data (str): The URL or text to encode in the QR code
        output_path (str): Path where to save the QR code image (optional)
        size (int): Size of each box in the QR code grid (default: 10)
        border (int): Border size in boxes (default: 4)
    
    Returns:
        str: Path to the saved QR code image
    """
    # Create QR code instance
    qr = qrcode.QRCode(
        version=1,  # Controls the size of the QR code
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=size,
        border=border,
    )
    
    # Add data
    qr.add_data(data)
    qr.make(fit=True)
    
    # Create an image
    img = qr.make_image(fill_color="black", back_color="white")
    
    # Determine output path
    if output_path is None:
        output_path = "qr_code.png"
    
    # Save the image
    img.save(output_path)
    print(f"QR code generated successfully: {output_path}")
    
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="Generate a QR code from a URL or text string"
    )
    parser.add_argument(
        "data",
        help="The URL or text to encode in the QR code"
    )
    parser.add_argument(
        "-o", "--output",
        help="Output file path (default: qr_code.png)",
        default="qr_code.png"
    )
    parser.add_argument(
        "-s", "--size",
        type=int,
        help="Size of each box in the QR code grid (default: 10)",
        default=10
    )
    parser.add_argument(
        "-b", "--border",
        type=int,
        help="Border size in boxes (default: 4)",
        default=4
    )
    
    args = parser.parse_args()
    
    try:
        generate_qr_code(args.data, args.output, args.size, args.border)
        return 0
    except Exception as e:
        print(f"Error generating QR code: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
