"""Conversions and calculations."""
import math
import re
import logging
from random import randint
from random import random

from enso.messages import displayMessage
from enso import selection


def cmd_calculate(ensoapi, expression=None):
    """
    Calculate the given mathematical expression (or selection).

    <i>Calculates the given mathematical expression.</i><br /><br />
    <b>Supported operators:</b> -, +, /, *, ^, **, (, ), %, mod<br />
    <b>Supported functions:</b> acos, asin, atan, atan2, ceil, cos, cosh,
                                degrees, exp, fabs, floor, fmod, frexp, hypot,
                                ldexp, log, log10, mod, modf, pow, radians, sin,
                                sinh, sqrt, tan, tanh<br />
    <b>Supported constants:</b> e, pi
    """
    seldict = ensoapi.get_selection()
    selected_text = None
    if seldict.get("text"):
        selected_text = seldict.get("text", "").strip().strip("\0")

    got_selection = False
    if selected_text and expression is None:
        expression = selected_text
        got_selection = expression is not None

    if expression is None:
        ensoapi.display_message("No expression given.")
        return

    math_funcs = [f for f in dir(math) if f[:2] != '__']
    ops1 = [' ', r'\.', ',', r'\-', r'\+', '/', r'\\', r'\*', r'\^', r'\*\*',
            r'\(', r'\)', '%', r'\d+']
    ops2 = ['abs', r'chr\([0-9]+\)', r'hex\([0-9]+\)', 'mod']

    whitelist = '|'.join(ops1 + ops2 + math_funcs)

    math_funcs_dict = {mf: eval('math.%s' % mf) for mf in math_funcs}
    math_funcs_dict['abs'] = abs
    math_funcs_dict['chr'] = chr
    math_funcs_dict['hex'] = hex

    expression = expression.replace(' mod ', ' % ')
    if expression.endswith("="):
        expression = expression[:-1]
        append_result = True
    else:
        append_result = False
    if re.match(whitelist, expression):
        try:
            result = eval(expression, {"__builtins__": None}, math_funcs_dict)

            pasted = False
            if got_selection:
                if append_result:
                    txt = expression.strip() + " = " + unicode(result)
                    pasted = selection.set({"text": txt})
                else:
                    pasted = selection.set({"text": unicode(result)})

            if not pasted:
                msg = "<p>%s</p><caption>%s</caption>" % (result, expression)
                displayMessage(msg)
        except SyntaxError as error:
            logging.info(error)
            ensoapi.display_message("Invalid syntax", "Error")
    else:
        ensoapi.display_message("Invalid expression", "Error")


def cmd_random(ensoapi, X_Y=""):
    """Display a random integer between X and Y (inclusive)."""
    try:
        X, Y = [int(i) for i in X_Y.split()[:2]]
    except ValueError:
        err = "Insufficient boundaries (X and Y must both be given)"
        ensoapi.display_message(err)
        return

    try:
        ensoapi.display_message(str(randint(X, Y)))
    except ValueError:
        err = "Incorrect boundaries (X must be smaller than or equal to Y)."
        ensoapi.display_message(err)


def cmd_random_paste(ensoapi, X_Y=""):
    """Replace selected text with random integer between X and Y (inclusive)."""
    if X_Y == "":
        """Replace selected text with random float within the range [0, 1)."""
        ensoapi.set_selection({"text": str(random())})
    else:
        try:
            X, Y = [int(i) for i in X_Y.split()[:2]]
        except ValueError:
            err = "Insufficient boundaries (X and Y must both be given)"
            ensoapi.display_message(err)
            return

        try:
            ensoapi.set_selection({"text": str(randint(X, Y))})
        except ValueError:
            err = "Incorrect boundaries (X must be smaller than or equal to Y)."
            ensoapi.display_message(err)
