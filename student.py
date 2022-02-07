from datetime import date
from enum import Enum

import pandas as pd

from helper import ordinal_number

class Gender(Enum):
    MALE = "male"
    FEMALE = "female"

class Student:
    def __init__(self, name_first, name_middle, name_last, gender, grade, age, birthday):
        self.name_first = name_first
        self.name_middle = name_middle
        self.name_last = name_last
        self.name_full = " ".join([string for string in [self.name_first, self.name_middle, self.name_last] if string is not None])

        self.initial_first = self.name_first[0].upper() if self.name_first else ''
        self.initial_middle = self.name_middle[0].upper() if self.name_middle else ''
        self.initial_last = self.name_last[0].upper() if self.name_last else ''

        self.gender = Gender(gender.lower())
        if self.gender == Gender.MALE:
            self.pronoun_subject, self.pronoun_object = "he", "him"
            self.pronoun_possessive_dependent, self.pronoun_possessive_independent = "his", "his"
            self.pronoun_reflexive = "himself"
        elif self.gender == Gender.FEMALE:
            self.pronoun_subject, self.pronoun_object = "she", "her"
            self.pronoun_possessive_dependent, self.pronoun_possessive_independent = "her", "hers"
            self.pronoun_reflexive = "herself"
        else:
            self.pronoun_subject, self.pronoun_object = "they", "them"
            self.pronoun_possessive_dependent, self.pronoun_possessive_independent = "their", "theirs"
            self.pronoun_reflexive = "themself"
        
        self.grade = grade
        self.grade_ordinal = ordinal_number(self.grade)
        self.age = age
        self.birthday = birthday