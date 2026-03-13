from PIL import Image, ImageDraw

WIDTH, HEIGHT = 300, 400

img = Image.new("L", (WIDTH, HEIGHT), color=240)
draw = ImageDraw.Draw(img)

# Ground strip at the bottom
draw.rectangle([0, 350, WIDTH, HEIGHT], fill=120)

# Tree trunk
trunk_x0, trunk_y0 = 130, 250
trunk_x1, trunk_y1 = 170, 355
draw.rectangle([trunk_x0, trunk_y0, trunk_x1, trunk_y1], fill=80)

# Tree crown – three overlapping ellipses for depth
draw.ellipse([60, 180, 240, 310], fill=70)
draw.ellipse([80, 120, 220, 260], fill=85)
draw.ellipse([100, 60, 200, 200], fill=60)

# Subtle shadow on crown (right side, slightly darker)
draw.ellipse([150, 100, 220, 220], fill=50)

img.save("223344.png")
print("Saved 223344.png")
