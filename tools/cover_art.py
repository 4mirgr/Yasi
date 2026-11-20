"""Original vector scientific motifs for the three book covers (no raster text)."""
import math


def science_art(volume):
    pieces = ['<svg class="science-art" viewBox="0 0 680 440" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">',
              '<defs><radialGradient id="halo"><stop stop-color="#7fffe0" stop-opacity=".2"/><stop offset="1" stop-color="#7fffe0" stop-opacity="0"/></radialGradient>',
              '<linearGradient id="strand" x2="1" y2="1"><stop stop-color="#8bffe1"/><stop offset="1" stop-color="#2fa9d3"/></linearGradient>',
              '<linearGradient id="capsid" x2=".7" y2="1"><stop stop-color="#d3d9ff"/><stop offset="1" stop-color="#7774cb"/></linearGradient></defs>',
              '<ellipse cx="340" cy="240" rx="330" ry="220" fill="url(#halo)"/>']
    # Faint molecular network creates depth without competing with the title.
    for x, y, r in ((58, 110, 8), (105, 170, 5), (570, 100, 7), (610, 170, 5),
                    (560, 350, 8), (110, 360, 7), (48, 275, 4), (630, 295, 6)):
        pieces.append(f'<circle cx="{x}" cy="{y}" r="{r}" fill="none" stroke="#c6f1ff" stroke-opacity=".3" stroke-width="2"/>')
    pieces.append('<path d="M58 110L105 170L48 275L110 360M570 100L610 170L630 295L560 350" fill="none" stroke="#c6f1ff" stroke-opacity=".12" stroke-width="2"/>')
    if volume == 12:
        # A circular plasmid, gene segments and a spatial double helix.
        pieces += ['<ellipse cx="397" cy="239" rx="154" ry="140" fill="#062e3c" fill-opacity=".45" stroke="#63d4c0" stroke-width="3"/>',
                   '<ellipse cx="397" cy="239" rx="131" ry="118" fill="none" stroke="#a5f8da" stroke-opacity=".32" stroke-width="2"/>']
        for start, color in ((-85, '#f5cf83'), (32, '#7ee5e0'), (143, '#7fafff')):
            coords = []
            for angle in range(start, start + 66, 2):
                a = math.radians(angle)
                coords.append(f'{397+154*math.cos(a):.2f},{239+140*math.sin(a):.2f}')
            pieces.append(f'<polyline points="{" ".join(coords)}" fill="none" stroke="{color}" stroke-width="16" stroke-linecap="round"/>')
        for y in range(45, 396, 15):
            angle = (y-45)/37
            x1 = 184 + 56 * math.sin(angle)
            x2 = 184 - 56 * math.sin(angle)
            pieces.append(f'<path d="M{x1:.1f} {y}L{x2:.1f} {y}" stroke="#b7fce1" stroke-opacity=".65" stroke-width="4"/>')
        for sign, color in ((1, '#7df3d6'), (-1, '#f5cf83')):
            coords = [f'{184+sign*56*math.sin((y-45)/37):.2f},{y}' for y in range(45, 396, 2)]
            pieces.append(f'<polyline points="{" ".join(coords)}" fill="none" stroke="{color}" stroke-width="9" stroke-linecap="round"/>')
        pieces += ['<path d="M294 225Q355 155 424 208T494 253Q422 296 352 263T294 225" fill="none" stroke="#91e7e6" stroke-width="4" stroke-opacity=".65"/>',
                   '<circle cx="393" cy="228" r="18" fill="#f5cf83" fill-opacity=".85"/>']
    elif volume == 13:
        # A cutaway enveloped virion with capsid facets and an RNA ribbon.
        for a in range(0, 360, 24):
            rad = math.radians(a)
            x1, y1 = 350+143*math.cos(rad), 230+143*math.sin(rad)
            x2, y2 = 350+174*math.cos(rad), 230+174*math.sin(rad)
            pieces.append(f'<path d="M{x1:.2f} {y1:.2f}L{x2:.2f} {y2:.2f}" stroke="#b9c5ff" stroke-width="6"/><circle cx="{x2:.2f}" cy="{y2:.2f}" r="8" fill="#f4d18d"/>')
        pieces += ['<circle cx="350" cy="230" r="142" fill="#4c497d" stroke="#c1cbff" stroke-width="7"/>',
                   '<circle cx="350" cy="230" r="128" fill="#252949" stroke="#b4c7ed" stroke-width="2" stroke-dasharray="3 8"/>',
                   '<path d="M350 126L444 179L444 282L350 336L256 282L256 179Z" fill="url(#capsid)" fill-opacity=".5" stroke="#b6c7ff" stroke-width="3"/>',
                   '<path d="M350 126L350 230L444 179M350 230L444 282M350 230L350 336M350 230L256 282M350 230L256 179" fill="none" stroke="#aebcf4" stroke-width="2"/>',
                   '<path d="M304 184C401 169 299 236 389 224S391 286 317 277S380 252 350 200" fill="none" stroke="#ffe0a0" stroke-width="7" stroke-linecap="round"/>',
                   '<ellipse cx="350" cy="230" rx="242" ry="72" fill="none" stroke="#b0c6ff" stroke-opacity=".28" transform="rotate(-25 350 230)" stroke-width="2"/>']
    else:
        # Three year booklets, answer choices, and a magnifying glass.
        for dx, dy, angle, color in ((-50, -15, -13, '#467799'), (35, -5, 11, '#77acbd'), (0, 20, -3, '#e4f1ef')):
            pieces.append(f'<g transform="translate({dx} {dy}) rotate({angle} 340 220)"><rect x="223" y="66" width="234" height="292" rx="12" fill="{color}" stroke="#b3e4ed" stroke-width="2"/>')
            pieces += ['<rect x="245" y="90" width="190" height="30" rx="5" fill="#1c526b"/>',
                       '<path d="M253 145H425M253 157H390" stroke="#285e75" stroke-width="5" stroke-opacity=".65"/>']
            for row in range(5):
                y = 190 + row*30
                for col in range(4):
                    x = 267+col*45
                    chosen = col == (row+1)%4
                    pieces.append(f'<circle cx="{x}" cy="{y}" r="7" fill="{"#277c8e" if chosen else "none"}" stroke="#285e75" stroke-width="2"/>')
            pieces.append('</g>')
        pieces += ['<path d="M432 291L510 370" stroke="#f5cf83" stroke-width="23" stroke-linecap="round"/>',
                   '<circle cx="410" cy="263" r="64" fill="#143e5c" fill-opacity=".9" stroke="#f5cf83" stroke-width="10"/>',
                   '<path d="M380 263L401 284L440 240" fill="none" stroke="#8aefcd" stroke-width="10" stroke-linecap="round" stroke-linejoin="round"/>']
    pieces.append('</svg>')
    return ''.join(pieces)
