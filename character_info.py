# -*- coding: utf-8 -*-

import logging

####################
# Single Character Data
SCD_SERVER = "server"
SCD_NAME   = "name"
SCD_CLASS  = "class"
SCD_LVL    = "level"
SCD_ILVL   = "ilvl"


class CharacterInfo(dict):
    """Enchanced dictionary containing a character data. Only the keys/values
    in the dictionary will be saved.
    The class members only have a utility purpose."""

    common_headers = (SCD_SERVER, SCD_NAME, SCD_CLASS, SCD_ILVL, SCD_LVL)

    def __init__(self, server, name):
        """Construct a new CharacterInfo object

        Args:
            server (str): server of the character
            name (str): name of the character
        """
        super().__init__()
        self[SCD_SERVER] = server
        self[SCD_NAME] = name
        self.nb_missing_gems = None
        self.missing_enchants = []

    @property
    def server(self):
        """Character's server

        Returns:
            (str) name of the server
        """
        return self[SCD_SERVER]

    @property
    def name(self):
        """Character's name

        Returns:
            (str) name of the character
        """
        return self[SCD_NAME]

    @property
    def classname(self):
        """Character's class

        Returns:
            (str) classname of the character
        """
        return self[SCD_CLASS]

    @classname.setter
    def classname(self, classname):
        self.set_data(SCD_CLASS, classname)

    @property
    def level(self):
        """Character's level

        Returns:
            (int) level of the character
        """
        return self[SCD_LVL]

    @level.setter
    def level(self, level):
        self.set_data(SCD_LVL, int(level))

    @property
    def ilevel(self):
        """Character's item level

        Returns:
            (int) item level of the character
        """
        return self[SCD_ILVL]

    @ilevel.setter
    def ilevel(self, ilevel):
        self.set_data(SCD_ILVL, ilevel)

    def set_data(self, key, value):
        """Set character info and log it (INFO)

        Args:
            key (str): key of the info to set
            value (str): value of the info
        """
        self[key] = value
        logging.info("%s: %s", key, value)

    def get_hex_color(self):
        """Get the color according to the character's class.

        Returns:
            (string) the hex code
        """
        COLOR_DICT = {
            "Druid": "#FF7D0A",
            "Warlock": "#9482C9",
            "Shaman": "#0070DE",
            "Paladin": "#F58CBA",
            "Warrior": "#C79C6E",
            "Priest": "#FFFFFF",
            "Death Knight": "#C41F3B",
            "Demon Hunter": "#A330C9",
            "Monk": "#00FF96",
            "Mage": "#69CCF0",
            "Hunter": "#ABD473",
            "Rogue": "#FFF569",
        }
        if self.classname in COLOR_DICT:
            return COLOR_DICT[self.classname]
        elif self.classname:
            logging.error("Color not found for class '%s'" % self.classname)
        else:
            logging.error("Classname not set when requiring color")
        return "#FFFFFF"

    def has_missing_gem_or_enchant(self):
        return self.nb_missing_gems or self.missing_enchants
