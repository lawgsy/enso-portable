"""Text tools."""

import re


def get_text(ensoapi):
    """Return the selected text as supplied by ensoapi."""
    return ensoapi.get_selection().get("text", "")


def display_count(ensoapi, msg_type, count):
    """Display count of 'msg_type'."""
    msg = "%s %s." % (count, msg_type) if count > 0 else "No selection."
    ensoapi.display_message(msg)


def display_replace(ensoapi, newtext):
    """Replace selection with modified text."""
    if newtext:
        ensoapi.set_selection({"text": newtext})
    else:
        ensoapi.display_message("No selection.")


def cmd_count_lines(ensoapi):
    """Count lines in selected text."""
    count = len(get_text(ensoapi).splitlines())
    display_count(ensoapi, "lines", count)


def cmd_count_characters(ensoapi):
    """Count characters in selected text."""
    count = len(re.sub(r'[\n\r]', '', get_text(ensoapi)))
    display_count(ensoapi, "characters", count)


def cmd_count_words(ensoapi):
    """Count words in selected text."""
    count = len(get_text(ensoapi).strip().split())
    display_count(ensoapi, "words", count)


def cmd_uppercase(ensoapi):
    """Uppercase selected text."""
    display_replace(ensoapi, get_text(ensoapi).upper())


def cmd_lowercase(ensoapi):
    """Lowercase selected text."""
    display_replace(ensoapi, get_text(ensoapi).lower())


def cmd_titlecase(ensoapi):
    """Titlecase selected text."""
    display_replace(ensoapi, get_text(ensoapi).title())


def cmd_scramble(ensoapi):
    """Titlecase selected text."""
    from random import shuffle
    char_list = list(get_text(ensoapi))
    shuffle(char_list)
    display_replace(ensoapi, ''.join(char_list))


def cmd_repeat(ensoapi, repitition="1"):
    """Titlecase selected text."""
    repitition = unicode(repitition, 'utf-8')
    if repitition.isnumeric():
        display_replace(ensoapi, get_text(ensoapi)*int(repitition))
    else:
        ensoapi.display_message("Repitition is not a number.")


def cmd_unaccent(ensoapi):
    """Unaccent (normalize) selected text."""
    import unicodedata
    text = get_text(ensoapi)
    newtext = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore')
    display_replace(ensoapi, newtext)


def cmd_datetime(ensoapi):
    """Insert short datetime string in the form yymmdd-HHMMss."""
    from time import localtime, strftime
    date_time = strftime("%y%m%d-%H%M%S", localtime())
    ensoapi.set_selection(date_time)


def cmd_unquote_url(ensoapi):
    """Decode encoded characters in a URL."""
    import urllib
    display_replace(ensoapi, urllib.unquote(get_text(ensoapi)))


def cmd_sort(ensoapi):
    """Sort the words in the current selection. (non alpha-num delimiters)."""
    if get_text(ensoapi):
        # wordlist = re.findall("[\w+('?\w*)?]", get_text(ensoapi))
        # wordlist = re.split(r'\W+', get_text(ensoapi))
        wordlist = filter(None, re.split(r"\W+", get_text(ensoapi)))
        wordlist = sorted(wordlist, key=lambda x: x.lower())
        ensoapi.set_selection(', '.join(wordlist))
    else:
        ensoapi.display_message("No selection")
