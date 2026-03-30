"""Generates icon.ico for KnobMixer — ghost icon."""
from PIL import Image, ImageDraw

def make_ghost(size, color="#1DB954"):
    sz  = size
    img = Image.new("RGBA", (sz, sz), (0,0,0,0))
    d   = ImageDraw.Draw(img)
    m   = max(2, sz//16)
    cx  = sz//2
    d.ellipse([m, m, sz-m, sz//2+m], fill=color)
    d.rectangle([m, sz//3, sz-m, sz-m], fill=color)
    bump = (sz-2*m)//3
    for i in range(3):
        x0=m+i*bump; x1=m+(i+1)*bump; ymid=sz-m
        d.ellipse([x0, ymid-bump//2, x1, ymid+bump//2], fill=(0,0,0,0))
    ey = sz//4; er = max(2, sz//10)
    d.ellipse([cx-sz//5-er, ey, cx-sz//5+er, ey+er*2], fill="white")
    d.ellipse([cx+sz//5-er, ey, cx+sz//5+er, ey+er*2], fill="white")
    return img

frames = [make_ghost(sz) for sz in [256,128,64,48,32,16]]
frames[0].save("icon.ico", format="ICO",
               sizes=[(f.width,f.height) for f in frames],
               append_images=frames[1:])
print("[OK] icon.ico created (ghost icon).")
