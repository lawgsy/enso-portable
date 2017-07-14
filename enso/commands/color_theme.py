"""Colour themes."""

COLOR_THEMES = {
    'default': ("#ffffff", "#9fbe57", "#7f9845", "#000000"),
    'green':   ("#ffffff", "#9fbe57", "#7f9845", "#000000"),
    'orange':  ("#ffffff", "#be9f57", "#987f45", "#000000"),
    'magenta': ("#ffffff", "#be579f", "#98457f", "#000000"),
    'cyan':    ("#ffffff", "#99cccc", "#99aaaa", "#000000"),
    'red':     ("#ffffff", "#cc0033", "#ff0066", "#000000")
    }


def cmd_enso_theme(ensoapi, color=None):
    """Change Enso theme color."""
    if not color:
        return

    from enso.quasimode import layout
    layout.WHITE = COLOR_THEMES[color][0]
    layout.DESIGNER_GREEN = COLOR_THEMES[color][1]
    layout.DARK_GREEN = COLOR_THEMES[color][2]
    layout.BLACK = COLOR_THEMES[color][3]
    layout.DESCRIPTION_BACKGROUND_COLOR = layout.DESIGNER_GREEN + "cc"
    layout.MAIN_BACKGROUND_COLOR = layout.BLACK + "d8"
    ensoapi.display_message(u"Enso theme changed to “%s”" % color, "enso")

cmd_enso_theme.valid_args = COLOR_THEMES.keys()
