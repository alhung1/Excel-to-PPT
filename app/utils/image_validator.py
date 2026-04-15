"""
Image validation utilities for verifying captured chart images.
"""
import os
import statistics
from PIL import Image
from app.config import (
    logger,
    IMAGE_MIN_SIZE_BYTES,
    IMAGE_MIN_UNIQUE_COLORS,
    IMAGE_MIN_STDEV,
)


def validate_image(path: str, min_size: int = None) -> bool:
    """Check if an image file exists, has reasonable size, and has actual content.

    Args:
        path: Path to the image file.
        min_size: Minimum file size in bytes (defaults to config value).

    Returns:
        True if the image is valid, False otherwise.
    """
    if min_size is None:
        min_size = IMAGE_MIN_SIZE_BYTES

    if not os.path.exists(path):
        logger.warning("Validate: file does not exist: %s", path)
        return False

    size = os.path.getsize(path)
    if size < min_size:
        logger.warning("Validate: file too small: %d bytes (min: %d)", size, min_size)
        return False

    # Check image content using PIL
    try:
        img = Image.open(path)
        gray = img.convert("L")
        pixels = list(gray.getdata())

        unique_colors = len(set(pixels))
        if unique_colors < IMAGE_MIN_UNIQUE_COLORS:
            logger.warning(
                "Validate: image appears blank — only %d unique colors", unique_colors
            )
            img.close()
            return False

        try:
            stdev = statistics.stdev(pixels)
            if stdev < IMAGE_MIN_STDEV:
                logger.warning(
                    "Validate: image has very low variance (stdev=%.2f), likely blank",
                    stdev,
                )
                img.close()
                return False
        except statistics.StatisticsError:
            pass  # Not enough data points

        logger.debug("Validate: OK — %d bytes, %d colors", size, unique_colors)
        img.close()

    except Exception as e:
        logger.warning("Validate: PIL check failed (%s), relying on file size only", e)

    return True
