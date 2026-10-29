import os
from PIL import Image, ImageDraw, ImageFont

def render_angel(size, font_size, bg_color=None, y_offset=0):
    if bg_color:
        img = Image.new('RGBA', (size, size), bg_color)
    else:
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0)) # transparent
    
    font = ImageFont.truetype('C:/Windows/Fonts/seguiemj.ttf', font_size)
    draw = ImageDraw.Draw(img)
    center_x = size // 2
    center_y = (size // 2) + y_offset
    draw.text((center_x, center_y), '\U0001F47C', font=font, anchor='mm', embedded_color=True)
    return img

public_dir = 'd:/Tool/MyMemo/public'
os.makedirs(public_dir, exist_ok=True)

# 1. Multi-resolution transparent ICO for Windows Title Bar, Taskbar & Favicon
ico_sizes = [16, 24, 32, 48, 64, 128, 256]
ico_images = []
for s in ico_sizes:
    f_size = int(s * 0.82)
    ico_images.append(render_angel(s, f_size, None, y_offset=1 if s > 16 else 0))

# Save ICO
ico_images[-1].save(
    f'{public_dir}/favicon.ico',
    format='ICO',
    sizes=[(s, s) for s in ico_sizes],
    append_images=ico_images[:-1]
)

# 2. Individual transparent PNGs
render_angel(16, 13).save(f'{public_dir}/favicon-16x16.png')
render_angel(32, 26).save(f'{public_dir}/favicon-32x32.png')
render_angel(48, 39).save(f'{public_dir}/favicon-48x48.png')
render_angel(192, 155).save(f'{public_dir}/pwa-192x192.png')
render_angel(512, 415).save(f'{public_dir}/pwa-512x512.png')

# 3. Maskable & Apple Touch (with dark slate background #0f172a)
render_angel(512, 330, (15, 23, 42, 255), y_offset=5).save(f'{public_dir}/pwa-maskable-512x512.png')
render_angel(180, 116, (15, 23, 42, 255), y_offset=2).save(f'{public_dir}/apple-touch-icon.png')

# 4. Transparent SVG
svg_content = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">
  <text x="50" y="58" font-size="75" text-anchor="middle" dominant-baseline="central" font-family="'Segoe UI Emoji', 'Apple Color Emoji', 'Noto Color Emoji', sans-serif">&#x1F47C;</text>
</svg>'''
with open(f'{public_dir}/favicon.svg', 'w', encoding='utf-8') as f:
    f.write(svg_content)

# 5. DIRECT UPDATE TO CHROME INSTALLED PWA CACHE ON THIS MACHINE!
chrome_pwa_ico = r'C:\Users\JUNG\AppData\Local\Google\Chrome\User Data\Profile 4\Web Applications\_crx_nkliikdoiniclpoijjfiiodbooanpoeh\MyMemo - GitHub Sync & Private Notes.ico'
if os.path.exists(chrome_pwa_ico):
    ico_images[-1].save(
        chrome_pwa_ico,
        format='ICO',
        sizes=[(s, s) for s in ico_sizes],
        append_images=ico_images[:-1]
    )
    print('CHROME_LOCAL_PWA_ICO_UPDATED')

print('ALL_TRANSPARENT_ANGEL_ICONS_GENERATED')
