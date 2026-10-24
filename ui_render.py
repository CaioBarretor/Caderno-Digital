"""Renderização de superfícies arredondadas em alta resolução para Tkinter."""
from PIL import Image, ImageDraw, ImageTk

_SCALE = 4


def rounded_rect_photo(master, width, height, radius, fill):
    width, height = max(1, int(width)), max(1, int(height))
    scale = _SCALE
    image = Image.new("RGBA", (width * scale, height * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle(
        (0, 0, width * scale - 1, height * scale - 1),
        radius=max(0, int(radius * scale)),
        fill=fill,
    )
    image = image.resize((width, height), Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(image, master=master)


def switch_photo(master, width, height, active, track_color, knob_color="#FFFFFF"):
    width, height = max(1, int(width)), max(1, int(height))
    scale = _SCALE
    image = Image.new("RGBA", (width * scale, height * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle(
        (scale, 2 * scale, (width - 1) * scale, (height - 2) * scale),
        radius=(height // 2) * scale,
        fill=track_color,
    )
    cx = (width - 12 if active else 12) * scale
    cy = (height // 2) * scale
    r = 8 * scale
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=knob_color)
    image = image.resize((width, height), Image.Resampling.LANCZOS)
    return ImageTk.PhotoImage(image, master=master)
