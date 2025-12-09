from talon import Module

mod = Module()

@mod.capture(rule="<user.letter>+ <number>")
def excel_reference(m) -> str:
    """Captures Excel cell references like B2, AA100"""
    letters = ""
    number = ""
    
    for item in m:
        if hasattr(item, 'isdigit'):
            if str(item).isdigit():
                number += str(item)
            else:
                letters += str(item).upper()
        else:
            letters += str(item).upper()
    
    return letters + number
