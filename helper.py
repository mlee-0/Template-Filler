def ordinal_number(integer):
    suffix = "th"
    if not (11 <= integer <= 19):
        last_digit = integer % 10
        if 1 <= last_digit <= 3:
            suffix = ["st", "nd", "rd"][last_digit - 1]
    return f"{str(integer)}{suffix}"

def ordinal_word(integer):
    pass