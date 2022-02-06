import re


# Return a string with instances of placeholder replaced by replacement.
def replace_text(placeholder: str, replacement: str, string: str):
    # Find instances in which the replacement text should be inserted as is.
    string = re.sub(placeholder, replacement, string)
    # Find instances in which the replacement text should be capitalized.
    string = re.sub(placeholder, replacement.capitalize(), string, flags=re.IGNORECASE)
    return string

# Return True if at least one placeholder is found in the string.
def is_placeholder_in_text(string, placeholder_prefix, placeholder_suffix):
    return re.search(f"{placeholder_prefix}.*{placeholder_suffix}", string) is not None

# Return a string containing an ordinal number as a digit and corresponding suffix (1st, 2nd, 3rd).
def ordinal_number(integer):
    suffix = "th"
    if not (11 <= integer <= 19):
        last_digit = integer % 10
        if 1 <= last_digit <= 3:
            suffix = ["st", "nd", "rd"][last_digit - 1]
    return f"{str(integer)}{suffix}"

# Return a string containing an ordinal number as a word (first, second, third).
def ordinal_word(integer):
    if 1 <= integer <= 12:
        return {1: "first", 2: "second", 3: "third", 4: "fourth", 5: "fifth", 6: "sixth", 7: "seventh", 8: "eighth", 9: "ninth", 10: "tenth", 11: "eleventh", 12: "twelfth"}[integer]
    else:
        return None

# Return True if no keys in the dictionary contain any capital letters.
def are_keys_lowercase(dictionary):
    return all(key.islower() for key in dictionary)