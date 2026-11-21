"""Small subject ornaments for the original illustrated Yasi cover."""


def science_art(volume):
    icons = {
        12: '<circle cx="50" cy="11" r="8"/><path d="M43 9a7 7 0 0 1 7-5m4 2a6 6 0 0 1 2 8" stroke-width="2.4"/>',
        13: '<path d="m50 2 8 4v9l-8 5-8-5V6Zm0 0-4 11 12-7m-12 7 4 7 5-8-13-6"/>',
        14: '<path d="M43 3h12v14H43zm2 0V1h12v14m-12-7 2 2 4-5m-6 8h7"/>',
    }
    return ('<svg class="science-art" viewBox="0 0 100 22" aria-hidden="true" '
            'fill="none" stroke="#74561b" stroke-width="1.3" stroke-linejoin="round">'
            '<path d="M2 11h35m26 0h35" stroke="#b28b38"/>'
            + icons[volume] + '</svg>')
