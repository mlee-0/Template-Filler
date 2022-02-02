from datetime import date
from enum import Enum

class Gender(Enum):
    MALE = "male"
    FEMALE = "female"
    OTHER = "other"

class Student:
    def __init__(self, name_first, name_middle, name_last, gender, age, birthday):
        self.name_first = name_first
        self.name_middle = name_middle
        self.name_last = name_last

        self.initial_first = self.name_first[0].upper()
        self.initial_middle = self.name_middle[0].upper()
        self.initial_last = self.name_last[0].upper()

        self.gender = Gender(gender.strip().lower())
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
        
        self.age = age
        self.birthday = birthday