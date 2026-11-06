# Logo « SM » de Smart Mail : carré arrondi en dégradé bleu (charte des panneaux),
# lettres blanches Avenir Next Heavy, étincelle ambrée (le côté « smart »).
import sys
from PIL import Image, ImageDraw, ImageFont, ImageFilter

FONT = "/System/Library/Fonts/Avenir Next.ttc"

def lerp(a, b, t): return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))

def logo(size, etincelle=True, plein=False):
    N = 1024
    img = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    # dégradé diagonal : bleu nuit en bas à gauche → bleu vif en haut à droite
    c1, c2 = (16, 42, 92), (59, 130, 246)
    grad = Image.new("RGBA", (N, N))
    px = grad.load()
    for y in range(N):
        for x in range(N):
            t = (x + (N - y)) / (2 * N)
            px[x, y] = lerp(c1, c2, t) + (255,)
    marge = 0 if size <= 32 else 40
    masque = Image.new("L", (N, N), 0)
    if plein:
        masque = Image.new("L", (N, N), 255)
    else:
        ImageDraw.Draw(masque).rounded_rectangle([marge, marge, N - marge, N - marge], radius=int((N - 2 * marge) * 0.24), fill=255)
    img.paste(grad, (0, 0), masque)
    # léger reflet en haut
    # lumière douce en haut à droite (pas de reflet « verre »)
    lum = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    ImageDraw.Draw(lum).ellipse([N * 0.45, -N * 0.35, N * 1.35, N * 0.55], fill=(255, 255, 255, 60))
    lum = lum.filter(ImageFilter.GaussianBlur(120))
    img = Image.alpha_composite(img, Image.composite(lum, Image.new("RGBA", (N, N), (0, 0, 0, 0)), masque))
    d = ImageDraw.Draw(img)
    # lettres
    taille = 390 if size > 32 else 520
    f = ImageFont.truetype(FONT, taille, index=8)
    texte = "SM"
    bb = d.textbbox((0, 0), texte, font=f)
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    x = (N - w) / 2 - bb[0]
    y = (N - h) / 2 - bb[1] + (40 if etincelle else 0)
    ombre = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    ImageDraw.Draw(ombre).text((x, y + 12), texte, font=f, fill=(6, 20, 50, 90))
    img = Image.alpha_composite(img, ombre.filter(ImageFilter.GaussianBlur(14)))
    d = ImageDraw.Draw(img)
    d.text((x, y), texte, font=f, fill=(255, 255, 255, 255))
    # étincelle à 4 branches, en haut à droite
    if etincelle:
        cx, cy, R, r = N * 0.755, N * 0.245, 82, 20
        pts = [(cx, cy - R), (cx + r, cy - r), (cx + R, cy), (cx + r, cy + r), (cx, cy + R), (cx - r, cy + r), (cx - R, cy), (cx - r, cy - r)]
        d.polygon(pts, fill=(251, 191, 36, 255))
        
    return img.resize((size, size), Image.LANCZOS)

def sobre(size, plein=False):
    """Version professionnelle : aplat bleu nuit, « SM » blanc, sans étincelle ni reflet."""
    N = 1024
    img = Image.new("RGBA", (N, N), (0, 0, 0, 0))
    masque = Image.new("L", (N, N), 255 if plein else 0)
    if not plein:
        ImageDraw.Draw(masque).rounded_rectangle([0, 0, N, N], radius=int(N * 0.22), fill=255)
    img.paste(Image.new("RGBA", (N, N), (17, 24, 39, 255)), (0, 0), masque)
    d = ImageDraw.Draw(img)
    f = ImageFont.truetype(FONT, 400, index=2)  # Avenir Next Demi Bold
    bb = d.textbbox((0, 0), "SM", font=f)
    d.text(((N - (bb[2] - bb[0])) / 2 - bb[0], (N - (bb[3] - bb[1])) / 2 - bb[1]), "SM", font=f, fill=(255, 255, 255, 255))
    return img.resize((size, size), Image.LANCZOS)

dossier = sys.argv[1]
if len(sys.argv) > 2 and sys.argv[2] == "sobre":
    sobre(192).save(f"{dossier}/icone-192.png")
    sobre(512).save(f"{dossier}/icone-512.png")
    sobre(512, plein=True).save(f"{dossier}/icone-maskable-512.png")
    sobre(180, plein=True).convert("RGB").save(f"{dossier}/apple-touch-icon.png")
    sobre(32).save(f"{dossier}/favicon-32.png")
elif len(sys.argv) > 2 and sys.argv[2] == "web":
    logo(192).save(f"{dossier}/icone-192.png")
    logo(512).save(f"{dossier}/icone-512.png")
    logo(512, plein=True).save(f"{dossier}/icone-maskable-512.png")
    logo(180, plein=True).convert("RGB").save(f"{dossier}/apple-touch-icon.png")
    logo(32).save(f"{dossier}/favicon-32.png")
else:
    for s in (16, 32, 48, 128):
        logo(s, etincelle=s >= 32).save(f"{dossier}/icon-{s}.png")
    logo(512).save(f"{dossier}/logo-512.png")
