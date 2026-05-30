import colorsys

import os
from IPython import get_ipython



def hex_to_ansi(hex_color):
    # Remove '#' and convert hex to RGB integers
    hex_color = hex_color.lstrip('#')
    r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return f"\033[38;2;{r};{g};{b}m"



def color_formatter(text: str, theme: str):
    
    color = hex_to_ansi(theme)
    reset = "\033[0m"  # Crucial: Resets terminal to default style
    
    return (f"{color}{text}{reset}")


def warn_print(text: str):
    
    
    print(f"\033[1;31m{text}\033[0m")


def hex_to_rgb(hex_str):

    """
    Converts #RRGGBB to (R, G, B) normalized to 0-1.
    
    """

    hex_str = hex_str.lstrip('#')

    return tuple(int(hex_str[i:i+2], 16) / 255.0 for i in (0, 2, 4))

def rgb_to_hex(rgb):

    """
    Converts normalized (R, G, B) back to #RRGGBB.
    
    """

    return '#' + ''.join(f'{int(round(c * 255)):02x}' for c in rgb)

def get_luminance(rgb):

    """
    Calculates relative luminance for WCAG contrast standards.
    
    """

    res = []
    for c in rgb:
        # Standard linearization of sRGB
        res.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
    return 0.2126 * res[0] + 0.7152 * res[1] + 0.0722 * res[2]

def get_contrast(rgb1, rgb2):

    """
    Calculates the contrast ratio between two colors.
    
    """

    l1, l2 = get_luminance(rgb1), get_luminance(rgb2)
    return (max(l1, l2) + 0.05) / (min(l1, l2) + 0.05)

def ensure_readable(hex_color, target_ratio=4.5):

    """
    Lightens a color until it reaches the target contrast on black.
    
    """

    rgb = hex_to_rgb(hex_color)
    black = (0, 0, 0)
    
    if get_contrast(rgb, black) >= target_ratio:
        return hex_color
    
    # Convert to HSL to adjust lightness (l) while keeping hue (h) and saturation (s)
    h, l, s = colorsys.rgb_to_hls(*rgb)
    
    # Binary search for the minimum lightness adjustment
    low, high = l, 1.0
    for _ in range(20):
        mid = (low + high) / 2
        if get_contrast(colorsys.hls_to_rgb(h, mid, s), black) >= target_ratio:
            high = mid
        else:
            low = mid
            
    return rgb_to_hex(colorsys.hls_to_rgb(h, high, s))


def use_black_text(color: str) -> bool:

    """
    Converts theme color to RGB and calculates 
    whether black text is required.
    
    """

    # Remove '#' if present and convert hex to RGB
    hex_color = color.lstrip('#')

    r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    # Standard formula for perceived brightness
    brightness = (r * 0.299 + g * 0.587 + b * 0.114) / 255
    
    # Use black text for light backgrounds, white for dark ones

    return brightness > 0.5





def is_vscode_notebook():
    # 1. Check if the environment is IPython-based
    try:
        shell = get_ipython().__class__.__name__
    except NameError:
        return False  # Not running in any interactive environment

    # 2. Check for VS Code specific markers
    # 'ZMQInteractiveShell' is used by Jupyter kernels (Notebooks/Interactive Window)
    # 'VSCODE_PID' or 'TERM_PROGRAM' identifies the VS Code host
    is_zmq = (shell == 'ZMQInteractiveShell')
    is_vscode = (
        "VSCODE_PID" in os.environ or 
        "VSCODE_CWD" in os.environ or 
        os.environ.get("TERM_PROGRAM") == "vscode"
    )

    return is_zmq and is_vscode