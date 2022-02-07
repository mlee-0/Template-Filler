import glob
import re

from settings import PLACEHOLDER_PREFIX, PLACEHOLDER_SUFFIX


# Return a string with instances of placeholder replaced by replacement.
def replace_text(placeholder: str, replacement: str, string: str) -> str:
    # Find instances in which the replacement text should be inserted as is.
    string = re.sub(placeholder, replacement, string)
    # Find instances in which the replacement text should be capitalized.
    string = re.sub(placeholder, replacement.capitalize(), string, flags=re.IGNORECASE)
    return string

# Return True if at least one placeholder is found in the string.
def is_placeholder_in_text(string: str) -> bool:
    return re.search(f"{PLACEHOLDER_PREFIX}.*{PLACEHOLDER_SUFFIX}", string) is not None

# Return the string with placeholder characters surrounding it.
def format_as_placeholder(string: str) -> str:
    return f"{PLACEHOLDER_PREFIX}{string}{PLACEHOLDER_SUFFIX}"

# Return a string containing an ordinal number as a digit and corresponding suffix (1st, 2nd, 3rd).
def ordinal_number(integer: int) -> str:
    suffix = "th"
    if not (11 <= integer <= 19):
        last_digit = integer % 10
        if 1 <= last_digit <= 3:
            suffix = ["st", "nd", "rd"][last_digit - 1]
    return f"{str(integer)}{suffix}"

# Return a string containing an ordinal number as a word (first, second, third).
def ordinal_word(integer: int) -> str:
    try:
        return {1: "first", 2: "second", 3: "third", 4: "fourth", 5: "fifth", 6: "sixth", 7: "seventh", 8: "eighth", 9: "ninth", 10: "tenth", 11: "eleventh", 12: "twelfth"}[integer]
    except KeyError:
        return str(integer)

# Return True if no keys in the dictionary contain any capital letters.
def are_keys_lowercase(dictionary: dict) -> bool:
    return all(key.islower() for key in dictionary)

# Return the first document of a specified type in the current folder, if found. Optionally require that a string is contained in its filename.
def detect_file(extension: str, string: str = None) -> str or None:
    filenames = [
        filename for filename in glob.glob(extension)
        if (string in filename.lower() if string is not None else True)
    ]
    if filenames:
        return filenames[0]