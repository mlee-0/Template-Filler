import glob
import re

from settings import PLACEHOLDER_PREFIX, PLACEHOLDER_SUFFIX


def replace_text(placeholder: str, replacement: str, string: str) -> str:
    """Return a string with instances of placeholder replaced by replacement."""
    # Find instances in which the replacement text should be inserted as is.
    string = re.sub(placeholder, replacement, string)
    # Find instances in which the replacement text should be capitalized.
    string = re.sub(placeholder, replacement.capitalize(), string, flags=re.IGNORECASE)
    return string

def is_placeholder_in_text(string: str) -> bool:
    """Return True if at least one placeholder is found in the string."""
    return re.search(f"{PLACEHOLDER_PREFIX}.*{PLACEHOLDER_SUFFIX}", string) is not None

def format_as_placeholder(string: str) -> str:
    """Return the string with placeholder characters surrounding it."""
    return f"{PLACEHOLDER_PREFIX}{string}{PLACEHOLDER_SUFFIX}"

def consolidate_runs(paragraph):
    """Return a paragraph that combines runs that share certain formatting."""
    # Dictionary used to assign runs' attributes to runs' style attributes.
    # font_to_style = {
    #     run.style.font.highlight_color: run.font.highlight_color,
    #     run.style.font.italic: run.italic,
    #     run.style.font.color.rgb: run.font.color.rgb,
    #     run.style.font.bold: run.bold,
    #     run.style.font.underline: run.underline,
    # }
    # To improve performance when comparing the formatting of two runs, order these attributes from most used to least used.
    formatting = lambda run: (run.font.highlight_color, run.italic, run.font.color.rgb, run.bold, run.underline)

    runs_consolidated = []
    for run in paragraph.runs:
        run.style.font.highlight_color = run.font.highlight_color
        if len(runs_consolidated) == 0 or formatting(runs_consolidated[-1]) != formatting(run):
            runs_consolidated.append(run)
        else:
            runs_consolidated[-1].text = f"{runs_consolidated[-1].text}{run.text}"
    
    paragraph.clear()
    for run in runs_consolidated:
        run.style.font.highlight_color = run.font.highlight_color
        run.style.font.italic = run.italic
        run.style.font.color.rgb = run.font.color.rgb
        run.style.font.bold = run.bold
        run.style.font.underline = run.underline

        paragraph.add_run(run.text, run.style)
    return paragraph

def ordinal_number(integer: int) -> str:
    """Return a string containing an ordinal number as a digit and corresponding suffix (1st, 2nd, 3rd)."""
    suffix = "th"
    if not (11 <= integer <= 19):
        last_digit = integer % 10
        if 1 <= last_digit <= 3:
            suffix = ["st", "nd", "rd"][last_digit - 1]
    return f"{str(integer)}{suffix}"

def ordinal_word(integer: int) -> str:
    """Return a string containing an ordinal number as a word (first, second, third)."""
    try:
        return {1: "first", 2: "second", 3: "third", 4: "fourth", 5: "fifth", 6: "sixth", 7: "seventh", 8: "eighth", 9: "ninth", 10: "tenth", 11: "eleventh", 12: "twelfth"}[integer]
    except KeyError:
        return str(integer)

def are_keys_lowercase(dictionary: dict) -> bool:
    """Return True if no keys in the dictionary contain any capital letters."""
    return all(key.islower() for key in dictionary)

def detect_file(extension: str, string: str = None) -> str or None:
    """Return the first document of a specified type in the current folder, if found. Optionally require that a string is contained in its filename."""
    filenames = [
        filename for filename in glob.glob(extension)
        if (string in filename.lower() if string is not None else True)
    ]
    if filenames:
        return filenames[0]