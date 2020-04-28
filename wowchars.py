# -*- coding: utf-8 -*-

__version__ = "4.0.0"

"""
TODO: Add missing docstrings
TODO: fix support of achievements
TODO: fix support of professions
TODO: remove hardcoded enchants / gems suggestions
      ==> replace by CSV file
      and/or
      ==> retrieve advices from websites like Noxxic or Icy Veins

TODO: add suboptions to --guild: min level, min ilvl
TODO: [Google Sheets] colorize names with class colors ?
TODO: [Google Sheets] header freeze + color?
TODO: [Google Sheets] automatically add graph(s) ?
TODO: [Google Sheets] batch color update instead of many single updates (might be faster)
"""

import argparse
import csv
import httplib2
import logging
import os
import re
import requests
import string
from time import strftime

# import googleapiclient
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

####################
# URLs
TOKEN_URL       = "https://{zone}.battle.net/oauth/token"
CHAR_URL        = "https://{zone}.api.blizzard.com/profile/wow/character/{server}/{name}?namespace=profile-{zone}&locale=en_GB&access_token={access_token}"
EQUIPMENT_URL   = "https://{zone}.api.blizzard.com/profile/wow/character/{server}/{name}/equipment?namespace=profile-{zone}&locale=en_GB&access_token={access_token}"
PROFESSIONS_URL = "https://{zone}.api.blizzard.com/profile/wow/character/{server}/{name}/professions?namespace=profile-{zone}&locale=en_GB&access_token={access_token}"
BASE_ACHIEV_URL = "https://{zone}.api.blizzard.com/wow/achievement/{id}?access_token={access_token}"
BASE_ITEM_URL   = "https://{zone}.api.blizzard.com/wow/item/{id}{slash_context}?bl={bonus_list}&access_token={access_token}"
CLASSES_URL     = "https://{zone}.api.blizzard.com/data/wow/playable-class/index?namespace=static-{zone}&locale=en_GB&access_token={access_token}"
GUILD_URL       = "https://{zone}.api.blizzard.com/data/wow/guild/{server}/{name}/roster?namespace=profile-{zone}&locale=en_GB&access_token={access_token}"

####################
# Headers
H_DATE         = "date"
H_SERVER       = "server"
H_NAME         = "name"
H_CLASS        = "class"
H_LVL          = "level"
H_ILVL         = "ilvl"
H_AZERITE_LVL  = "Azerite lvl"
H_ASHJRA_KAMAS = "Ashjra’kamas"

####################
# Achievements: {ID: stepped}
ACHIEVEMENTS = {
    # 11609: False,  # POWER_UNBOUND
}


def main():
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument("--blizzard-client-id", help="Client ID of Blizzard's Battle.net API", required=True)
    parser.add_argument("--blizzard-client-secret", help="Token to Blizzard's Battle.net API", required=True)
    parser.add_argument("-o", "--output", help="Output CSV file", required=False)
    parser.add_argument("-c", "--char", help="Check character (server:charname)", action="append", default=[], required=False)
    parser.add_argument("--guild", help="Check characters from given GUILD with minimum level of 111")
    parser.add_argument("-r", "--raid", action="store_true", help="Only keeps info that are usefull for raids (class, lvl, ilvl)")
    parser.add_argument("-s", "--summary", help="Display summary", action="store_true")
    parser.add_argument("-v", "--verbosity", action="count", default=0, help="increase output verbosity")
    parser.add_argument("-g", "--google-sheet", help="ID of the Google sheet where the results will be saved")
    parser.add_argument("-d", "--dry-run", action="store_true", help="does not update target output")
    parser.add_argument("--check-gear", action="store_true", help="inspect gear for legendaries or any missing gem/enchantment")
    parser.add_argument("--default-server", help="Default server when not given with the '-c' option", default=None)
    parser.add_argument("--zone", choices=["eu", "us", "kr", "tw"], help="Select server's zone.", default="eu")
    parser.add_argument('--version', action='version', version=__version__)
    args = parser.parse_args()

    set_logger(args.verbosity)

    ce = CharactersExtractor(args.blizzard_client_id,
                             args.blizzard_client_secret,
                             args.zone)
    ce.run(args.guild,
           args.char,
           args.raid,
           args.output,
           args.summary,
           args.check_gear,
           args.google_sheet,
           args.dry_run,
           args.default_server)


class CharInfo(dict):
    """Enchanced dictionary containing a character data. Only the keys/values
    in the dictionary will be saved.
    The class members only have a utility purpose."""

    def __init__(self, server, name):
        """Construct a new CharInfo object

        Args:
            server (str): server of the character
            name (str): name of the character
        """
        super().__init__()
        self[H_SERVER] = server
        self[H_NAME] = name
        self.nb_missing_gems = None
        self.missing_enchants = []

    def server(self):
        """Get character's server

        Returns:
            (str) name of the server
        """
        return self[H_SERVER]

    def name(self):
        """Get character's name

        Returns:
            (str) name of the character
        """
        return self[H_NAME]

    def fullname(self):
        """Get character's full name

        Returns:
            (str) name & server of the character
        """
        return "%s:%s" % (self.name(), self.server())

    def classname(self):
        """Get character's class

        Returns:
            (str) classname of the character
        """
        return self[H_CLASS]

    def level(self):
        """Get character's level

        Returns:
            (int) level of the character
        """
        return int(self[H_LVL])

    def ilevel(self):
        """Get character's item level

        Returns:
            (int) item level of the character
        """
        return int(self[H_ILVL])

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
        if self.classname() in COLOR_DICT:
            return COLOR_DICT[self.classname()]
        elif self.classname():
            logging.error("Color not found for class '%s'" % self.classname())
        else:
            logging.error("Classname not set when requiring color")
        return "#FFFFFF"

    def set_nb_missing_gems(self, nb_missing_gems):
        self.nb_missing_gems = nb_missing_gems

    def set_missing_enchants(self, missing_enchants):
        self.missing_enchants = missing_enchants

    def has_missing_gem_or_enchant(self):
        return self.nb_missing_gems or self.missing_enchants


class CharactersExtractor:
    """Class processing World Of Warcraft characters:
    - extract and process data from Blizzard API
    - save in CSV and/or export to a Google Sheets document"""

    def __init__(self, client_id, client_secret, zone):
        """Contructor

        Args:
            client_id (str): Blizzard client ID
            client_secret (str): Blizzard client secret
            zone (str): Zone of the target guild and/or characters
        """

        # getting auth token
        url = TOKEN_URL.format(zone=zone)
        logging.debug(url)
        r = requests.post(url, data={"grant_type": "client_credentials"},
                          auth=(client_id, client_secret))
        r.raise_for_status()

        self.access_token = r.json()["access_token"]
        logging.debug("Got access token: %s", self.access_token)
        self.zone = zone
        self.achievements = []  # achievement details
        self.characters = []    # fetched characters

    def run(self, guild, chars, raid, csv_output, summary,
            check_gear, google_sheet_id, dry_run,
            default_server):
        """main function

        Args:
            guild (str or None): guild to process
            chars (str array): list of characters to process
            raid (bool): only keeps info usefull for raids
            csv_output (str): if not None, save results in the CSV file
            summary (bool): print a results' sumamry
            check_gear (bool): check gear for any missing gem or enchantment
            google_sheet_id (str): if not None, save results in Google Sheets
            dry_run (boolean): does not modify the Google Sheets document
            default_server (string): default server if not given in 'chars'
        """
        self.fetch_achievements_details()

        guild_chars = self.find_guild_characters(guild, default_server) if guild else []

        for c in guild_chars + chars:
            self.fetch_char(c, default_server, raid, check_gear)

        if csv_output:
            self.save_csv(csv_output)

        if summary:
            self.display_summary()

        if google_sheet_id:
            self.save_summary_in_google_sheets(google_sheet_id, dry_run)
            if not raid:
                self.save_extra_google_sheets(google_sheet_id, dry_run)

        if check_gear:
            self.display_gear_to_fix()

    def get_known_char(self, server, name):
        """Search an already known/processed character

        Args:
            server (str): server ot the searched character
            name (str): name ot the searched character

        Returns
            a CharInfo object or None if not found
        """
        for r in self.characters:
            if (r.server() == server) and (r.name() == name):
                return r
        return None

    def find_guild_characters(self, serv_and_guildname, default_server=None):
        """Find characters from given list

        Args:
            serv_and_guildname (str): server and name of the guild to prcess.
                                      Expected format is "server:guildname".
                                      Using only the guildname is supported but
                                      will produce a warning.
        """
        print("======================================================")
        print("Processing guild: '%s'" % serv_and_guildname)
        server = None
        name = None
        try:
            server, name = serv_and_guildname.split(":")
        except ValueError:
            if default_server:
                logging.info("no server name, using default server '%s'", default_server)
                server = default_server
            else:
                logging.warning("no server name, using default server 'voljin'")
                server = "voljin"
            name = serv_and_guildname

        # Blizzard API does not recognize guild names with spaces anymore
        name = name.replace(" ", "-")

        url = GUILD_URL.format(zone=self.zone, access_token=self.access_token, server=server, name=name.replace(" ", "-"))
        logging.debug(url)
        r = requests.get(url)
        r.raise_for_status()
        body = r.json()
        logging.trace(body)
        members = body["members"]
        guild_chars = []
        for m in members:
            charname = m["character"]["name"]
            level = m["character"]["level"]
            realm = m["character"]["realm"]["slug"]
            logging.debug("%3d %s" % (level, charname))
            if level >= 111:
                logging.info("Found valid character: %3d %s" % (level, charname))
                guild_chars.append("%s:%s" % (realm, charname))
        return guild_chars

    def fetch_char(self, serv_and_name, default_server=None, raid=False,
                   check_gear=False):
        """Fetch and register a character from Blizzard's API.

        Args:
            serv_and_name (str): server and name of the character to fetch.
                                 Expected format is "server:name". Using only
                                 the name is supported but will produce a warning.
            default_server (string): default server if not given in 'serv_and_name'
            raid (bool): Only keeps info that are usefull for raids (class, lvl, ilvl)
            check_gear (bool): check gear for any missing gem or enchantment
        """
        print("======================================================")
        print("Processing: %s" % (serv_and_name))
        server = None
        name = None
        try:
            server, name = serv_and_name.split(":")
        except ValueError:
            if default_server:
                logging.info("no server name, using default server '%s'", default_server)
                server = default_server
            else:
                logging.warning("no server name, using default server 'voljin'")
                server = "voljin"
            name = serv_and_name

        if self.get_known_char(server, name):
            logging.warning("character '%s' already processed" % (serv_and_name))
            return

        char = CharInfo(server, name)
        try:
            self.fetch_char_base(char)
            self.fetch_char_items(char, check_gear)
            if not raid:
                self.fetch_char_achievements(char)
                self.fetch_char_professions(char)
        except (ValueError, KeyError, requests.exceptions.HTTPError):
            logging.error("cannot fetch %s/%s", server, name)
            return
        self.characters.append(char)

    def fetch_char_base(self, char):
        """Fetch and fill info for the given character: level + items related info

        Args:
            char (CharInfo): the character to fetch
        """
        url = CHAR_URL.format(zone=self.zone, access_token=self.access_token, server=char.server(), name=char.name().lower())
        logging.debug(url)
        r = requests.get(url)
        r.raise_for_status()
        char_json = r.json()
        logging.trace(char_json)
        char.set_data(H_CLASS, char_json["character_class"]["name"])
        char.set_data(H_LVL, str(char_json[H_LVL]))
        char.set_data(H_ILVL, str(char_json["average_item_level"]))

    def fetch_char_items(self, char, check_gear):
        """Fetch and fill info for the items of the given character

        Args:
            char (CharInfo): the character to fetch
            check_gear (bool): check gear for any missing gem or enchantment
        """
        url = EQUIPMENT_URL.format(zone=self.zone, access_token=self.access_token, server=char.server(), name=char.name().lower())
        logging.debug(url)
        r = requests.get(url)
        r.raise_for_status()
        items = r.json()["equipped_items"]

        # Searching the 'azeriteLevel' of the item in the neck slot
        try:
            neck = self.get_item(items, "NECK")
            char.set_data(H_AZERITE_LVL, str(neck["azerite_details"]["level"]["value"]))
        except (ValueError, KeyError):
            logging.warning("Cannot find azerite level.")
            char.set_data(H_AZERITE_LVL, "NA")

        # patch 8.3 introduced a new legendary cloak called 'Ashjra’kamas'
        # Searching the 'rank' of this item.
        try:
            back_item = self.get_item(items, "BACK")
            if back_item["item"]["id"] != 169223:
                raise ValueError("Item is not Ashjra’kamas: %s", back_item["name"])
            # TODO: find how to get rank, using itemLevel until then
            ilvl = back_item["level"]["value"]
            rank = back_item["name_description"]["display_string"]
            char.set_data(H_ASHJRA_KAMAS, "%d (%s)" % (ilvl, rank))
        except (ValueError, KeyError):
            logging.warning("Cannot find 'Ashjra’kamas'.")
            char.set_data(H_ASHJRA_KAMAS, "NA")

        # Checking gear
        if check_gear:
            total_empty_sockets = 0
            missing_enchants = []
            for item in items:
                nb_empty_sockets, missing_enchant = self.check_item_enchants_and_gems(item)
                total_empty_sockets += nb_empty_sockets
                if missing_enchant:
                    missing_enchants.append(item["inventory_type"]["name"])

            char.set_nb_missing_gems(total_empty_sockets)
            char.set_missing_enchants(missing_enchants)

    def check_item_enchants_and_gems(self, item):
        """Check any missing enchant or gem in the given item

        Args:
            item (dict): description of the item

        Returns:
            (int, boolean): (number of empty gem slot, true is not enchanted)

        """
        item_type = item["inventory_type"]["type"]
        slot = item["slot"]["name"]
        logging.debug("Checking enchants and gems of slot: %s", slot)

        nb_empty_sockets = 0
        missing_enchant = False

        # checking gem slots & checking with current item state
        if "sockets" in item:
            for s in item["sockets"]:
                if "item" not in s:
                    nb_empty_sockets += 1

        if nb_empty_sockets:
            logging.debug("Found %d empty socket(s)", nb_empty_sockets)

        # checking enchants
        if item_type in ["FINGER", "MAIN_HAND"]:
            if ("enchantments" not in item) or (not item["enchantments"]):
                logging.debug("Detected missing enchantment for this slot !")
                missing_enchant = True

        return nb_empty_sockets, missing_enchant

    def fetch_char_achievements(self, char):
        """Fetch and fill achievements

        Args:
            char (CharInfo): the character to fetch
        """
        try:
            url = CHAR_URL.format(zone=self.zone, access_token=self.access_token, server=char.server(), name=char.name().lower(), fields="achievements,quests")
            logging.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            obj = r.json()
            logging.trace(obj)
            achievements = obj["achievements"]
        except ValueError:
            logging.warning("cannot retrieve achievements for %s/%s", char.server(), char.name())

        for ach_desc in self.achievements:
            ach_id = ach_desc["id"]
            title = ach_desc["title"]
            if ACHIEVEMENTS[ach_id]:
                char.set_data(title, self.check_stepped_achievement(ach_desc, ach_id, achievements))
            else:
                char.set_data(title, self.check_achievement(ach_desc, ach_id, achievements))

    def check_achievement(self, ach_desc, ach_id, char_achievements):
        """Check if achievement is completed (ie. one of its criterias is met)

        Args:
            ach_desc (dict): description of the tested achievement
            ach_id (int): id of the tested achievement
            char_achievements (dict): achivements info of the character

        Returns:
            "Ok" if the achievement is validated. Else: ""
        """
        ach_ok = False
        if ach_id in char_achievements["achievementsCompleted"]:
            logging.debug("char has ach: %s", ach_desc)
            for criteria in ach_desc["criteria"]:
                crit_ok = False
                try:
                    crit_index = char_achievements["criteria"].index(criteria["id"])
                    if char_achievements["criteriaQuantity"][crit_index] >= criteria["max"]:
                        ach_ok = True
                        crit_ok = True
                except ValueError:
                    pass
                logging.debug("Criteria %s ==> %s", criteria, crit_ok)
        return "OK" if ach_ok else ""

    def check_stepped_achievement(self, ach_desc, ach_id, char_achievements):
        """Check if achievement is completed (all its criterias are met)

        Args:
            ach_desc (dict): description of the tested achievement
            ach_id (int): id of the tested achievement
            char_achievements (dict): achivements info of the character

        Returns:
            "X/X" if the achievement is validated. Else: "Y/X: {next incomplete step}"
        """
        count = 0
        next_step = ""
        if ach_id in char_achievements["achievementsCompleted"]:
            logging.debug("char has ach: %s", ach_desc)
            for criteria in ach_desc["criteria"]:
                if criteria["id"] in char_achievements["criteria"]:
                    logging.debug("Criteria '%s' ==> OK", criteria["description"])
                    count += 1
                else:
                    logging.debug("Criteria '%s' ==> NOK", criteria["description"])
                    next_step = (": " + criteria["description"]) if not next_step else next_step
        return "%d/%d%s" % (count, len(ach_desc["criteria"]), next_step)

    def fetch_char_professions(self, char):
        """Fetch and fill BfA professions

        Args:
            char (CharInfo): the character to fetch
        """
        try:
            url = PROFESSIONS_URL.format(zone=self.zone, access_token=self.access_token, server=char.server(), name=char.name().lower())
            logging.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            obj = r.json()
            logging.trace(obj)
            professions = obj["primaries"]
        except ValueError:
            logging.warning("cannot retrieve professions for %s/%s", char.server(), char.name())

        count = 0
        for profession in professions:
            for tier in profession["tiers"]:
                tier_name = tier["tier"]["name"]
                if tier_name.startswith("Kul Tiran") or tier_name.startswith("Zandalari"):
                    count += 1
                    tier_name = tier_name.replace("Kul Tiran ", "").replace("Zandalari ", "")
                    char.set_data("BfA profession %d" % count, "%s: %d" % (tier_name, tier["skill_points"]))

    def fetch_achievements_details(self):
        """Fetch details for the achievements to check"""
        print("======================================================")
        print("Fetching achievements details")
        for a_id in ACHIEVEMENTS:
            url = BASE_ACHIEV_URL.format(zone=self.zone, access_token=self.access_token, id=a_id)
            logging.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            ach = r.json()
            logging.trace(ach)
            logging.info("%6d: %s", ach["id"], ach["title"])
            self.achievements.append(ach)

    def get_achievement_title(self, ach_id):
        """Get name of a known achievement

        Args:
            ach_id (int): ID of the achievement

        Returns:
            The title of the achievement
        """
        for a in self.achievements:
            if ach_id == a["id"]:
                return a["title"]
        raise ValueError("Cannot retrieve achievement %d" % ach_id)

    def get_item(self, items, slot_type):
        """retrieve item of the selected slot.

        Args:
            items (array): list of items of a character
            slot_type (str): searched slot

        Returns:
            the item of the given slot or None
        """ 
        for item in items:
            if item["slot"]["type"] == slot_type:
                return item
        return None

    def save_csv(self, output_file):
        """Save the results in a CSV file

        Args:
            output_file (str): path to the CSV file
        """
        with open(output_file, 'w') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=self.get_ordered_fieldnames(), delimiter=';')
            writer.writeheader()
            for r in self.characters:
                writer.writerow(r)

    def get_ordered_fieldnames(self):
        """Sort the fieldnames (headers) in a specific order

        Returns;
            (str array) the ordered headers
        """
        fieldnames = [H_SERVER, H_NAME, H_CLASS, H_ILVL, H_LVL]
        for c in self.characters:
            for k in c.keys():
                if k not in fieldnames:
                    fieldnames.append(k)
        return fieldnames

    def display_summary(self):
        """Print a summary of the results"""
        print("======================================================")
        print("Fetched %d character(s)" % len(self.characters))
        fieldnames = self.get_ordered_fieldnames()
        # computing width of each column
        widths = {f: len(f) for f in fieldnames}
        for f in fieldnames:
            for char in self.characters:
                if f in char:
                    widths[f] = max(widths[f], len(str(char[f])))

        # printing fieldnames
        line = []
        for f in fieldnames:
            v = "%-" + str(widths[f]) + "s"
            line.append(v % f)
        print((", ").join(line))

        # printing rows
        for char in sorted(self.characters, key=lambda x: x[H_ILVL], reverse=True):
            line = []
            for f in fieldnames:
                v = "%-" + str(widths[f]) + "s"
                line.append(v % (char[f] if (f in char) else ""))
            print((", ").join(line))

    def display_gear_to_fix(self):
        """Display gear to fix: missing gems or enchantments"""
        print("======================================================")

        chars_to_fix = list(filter(lambda x: x.has_missing_gem_or_enchant(), self.characters))
        print("/!\\ %d character(s) to fix!" % len(chars_to_fix))
        for c in chars_to_fix:
            missing = ""
            if c.nb_missing_gems:
                missing += "%d gem%s" % (c.nb_missing_gems, "s" if (c.nb_missing_gems > 2) else "")
            if c.missing_enchants:
                missing += ", " + ", ".join(c.missing_enchants)
            print("%s: %s" % (c.name(), missing))

        print("---------------------------")
        to_create = {}  # {enchant : [nb, set of chars]}
        for c in chars_to_fix:
            nb_gems = c.nb_missing_gems
            if nb_gems:
                if "Gem" not in to_create:
                    to_create["Gem"] = [1, set()]
                else:
                    to_create["Gem"][0] += 1
                to_create["Gem"][1].add("%s(%d)" %(c.name(), nb_gems))

            for enchant in c.missing_enchants:
                if enchant not in to_create:
                    to_create[enchant] = [1, set()]
                else:
                    to_create[enchant][0] += 1
                to_create[enchant][1].add(c.name())

        for tc in sorted(to_create):
            print("%s: %s %s" % (tc, to_create[tc][0], to_create[tc][1]))

    def save_summary_in_google_sheets(self, google_sheet_id, dry_run):
        """Save summary in Google Sheets

        Args:
            google_sheet_id (str): the ID of the document
            dry_run (bool): if True, do not modify the document
        """
        print("======================================================")
        print("Synching summary in Google Sheets")

        SUMMARY = "summary"

        sc = SheetConnector(google_sheet_id, dry_run)
        sc.check_or_create_sheet(SUMMARY)
        fieldnames = self.get_ordered_fieldnames()
        sc.ensure_headers(SUMMARY, fieldnames)

        sheet_values, _range = sc.get_values(SUMMARY + "!A:Z")
        headers = sheet_values[0]
        h_indexes = {h: i for i, h in enumerate(headers)}
        values = sheet_values[1:]

        update_data = SheetBatchUpdateData()  # cell values to update
        to_colorize = []  # cells to colorize (when adding new character(s))

        # updating / adding characters info
        for r in sorted(self.characters, key=lambda x: x[H_ILVL], reverse=True):
            char_index = None
            # checking if character is already known
            for i, g_line in enumerate(values):
                if g_line[h_indexes[H_SERVER]] == r[H_SERVER]:
                    g_name = g_line[h_indexes[H_NAME]]
                    if (g_name == r.name()) or (g_name == r.name().lower()):
                        char_index = i
                        break

            if char_index is not None:
                g_row = values[char_index]
                for i, h in enumerate(headers):
                    if (h in fieldnames) and (h in r) and r[h] and ((i >= len(g_row)) or (r[h] != g_row[i])):
                        cell_col = column_letter(i)
                        cell_row = char_index + 2
                        update_data.add_data(SUMMARY, cell_col, cell_row, cell_col, cell_row, [[r[h]]])

            else:
                line = [(r[h] if h in r else None) for h in headers]
                values.append(line)
                row_index = len(values) + 1
                update_data.add_data(SUMMARY, "A", row_index, column_letter(len(line) - 1), row_index, [line])
                to_colorize.append((h_indexes[H_NAME], row_index, r.get_hex_color()))

        if update_data:
            sc.update_values(update_data)
            # TODO: batch update instead of single one
            for tc in to_colorize:
                color = RGBColor.from_hex(tc[2])
                column = column_letter(tc[0])
                sc.set_background_color(SUMMARY, column, tc[1], color)
        else:
            print("Nothing to update")

    def save_extra_google_sheets(self, google_sheet_id, dry_run):
        """Save level and ilvl in Google Sheets in separated Sheets

        Args:
            google_sheet_id (str): the ID of the document
            dry_run (bool): if True, do not modify the document
        """
        print("======================================================")
        print("Synching ilvl/level in Google Sheets")

        sc = SheetConnector(google_sheet_id, dry_run)
        names = [c.name() for c in sorted(self.characters, key=lambda x:x.name())]

        update_data = SheetBatchUpdateData()
        today = strftime("%Y-%m-%d")

        for s in [H_LVL, H_ILVL]:
            v_dict = {r[H_NAME]: r[s] for r in sorted(self.characters, key=lambda x: x[H_NAME])}
            sc.ensure_headers(s, [H_DATE] + names)
            sheet_values, _range = sc.get_values("%s!A:Z" % s)
            headers = sheet_values[0]
            h_indexes = {h: i for i, h in enumerate(headers)}
            update_needed = False
            last_update_today = (len(sheet_values) > 1) and (sheet_values[-1][h_indexes[H_DATE]] == today)
            # Checking if an update is needed on the sheet
            if len(sheet_values) <= 1:
                update_needed = True
            else:
                values = sheet_values[1:]
                last_line = values[-1]
                for name in sorted(v_dict):
                    if name not in h_indexes:
                        update_needed = True
                        break
                    h_i = h_indexes[name]
                    if len(last_line) <= h_i:
                        update_needed = True
                        break
                    new_v = int(v_dict[name])
                    cur_v = last_line[h_i]
                    if (not cur_v) or (int(cur_v) < new_v):
                        update_needed = True
                        break

            if not update_needed:
                continue

            line = values[-1] if last_update_today else ([None] * len(headers))
            for i, h in enumerate(headers):
                if h in v_dict:
                    if len(line) <= i:
                        line += [None] * (i + 1 - len(line))
                    line[i] = v_dict[h]
            line[h_indexes[H_DATE]] = today

            # Adding a new line if last date is not today, else updating the last one
            line_nb = len(sheet_values)
            if (len(sheet_values) <= 1) or (sheet_values[-1][h_indexes[H_DATE]] < today):
                line_nb += 1
            update_data.add_data(s, "A", line_nb, "Z", line_nb, [line])

        if update_data:
            sc.update_values(update_data)
        else:
            print("Nothing to update")


def install_trace_logger():
    if install_trace_logger.trace_installed:
        return
    level = logging.TRACE = logging.DEBUG - 5

    def log_logger(self, message, *args, **kwargs):
        if self.isEnabledFor(level):
            self._log(level, message, args, **kwargs)
    logging.getLoggerClass().trace = log_logger

    def log_root(msg, *args, **kwargs):
        logging.log(level, msg, *args, **kwargs)
    logging.addLevelName(level, "TRACE")
    logging.trace = log_root
    install_trace_logger.trace_installed = True


install_trace_logger.trace_installed = False


def set_logger(verbosity):
    """Initialize and set the logger

    Args:
        verbosity (int): 0 -> default, 1 -> info, 2 -> debug, 3 -> trace
    """
    log_level = logging.WARNING - (verbosity * 10)

    logging.basicConfig(
        level=log_level,
        format="[%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler()
        ]
    )
    install_trace_logger()


def column_letter(index):
    """In Sheets the columns are identified by letters, not integers.
    This function translates the column index into letter(s).

    Args:
        index (int): column index

    Returns:
        (str) the translated index as letter(s)
    """
    res = ""
    i = index
    while True:
        j = i % len(string.ascii_uppercase)
        res = string.ascii_uppercase[j] + res
        i = (i - j) // len(string.ascii_uppercase) - 1
        if i < 0:
            break
    return res


def column_index(column_str):
    """In Sheets the columns are identified by letters, not integers.
    This function translates the column string into an integer index.

    Args:
        index (str): column identifier

    Returns:
        (int) the translated index
    """
    res = string.ascii_uppercase.index(column_str[-1])
    mult = 1
    for l in column_str[-2::-1]:
        mult *= len(string.ascii_uppercase)
        res += (1 + string.ascii_uppercase.index(l)) * mult
    return res


class RGBColor:
    """Helper class to easily deal with colors"""
    def __init__(self, red, green, blue):
        """Constructor

        Args:
            red (int): red integer value (from 0 to 255)
            green (int): green integer value (from 0 to 255)
            blue (int): blue integer value (from 0 to 255)
        """
        self.red = red
        self.green = green
        self.blue = blue

    @classmethod
    def from_float_rgb(fgbc, fred, fgreen, fblue):
        """Factory to create RGBColor from float values

        Args:
            fgbc (int): red float value (from 0 to 1)
            fgreen (int): green float value (from 0 to 1)
            fblue (int): blue float value (from 0 to 1)

        Returns:
            a RGBColor object
        """
        r = int(int(round(fred * 255)))
        g = int(int(round(fgreen * 255)))
        b = int(int(round(fblue * 255)))
        return fgbc(r, g, b)

    @classmethod
    def from_float_rgb_dict(fgbc, d):
        """Factory to create RGBColor from float values

        Args:
            d (dict): dictionnary containing the red/green/blue float values

        Returns:
            a RGBColor object
        """
        r = d["red"] if "red" in d else 0.0
        g = d["green"] if "green" in d else 0.0
        b = d["blue"] if "blue" in d else 0.0
        return fgbc.from_float_rgb(r, g, b)

    @classmethod
    def from_hex(fgbc, hex_code):
        """Factory to create RGBColor from the hexadecimal value

        Args:
            hex_code (str): hexadecimal string of the color, with or without hash. (ex: "#FF0000")

        Returns:
            a RGBColor object
        """
        h = hex_code.lstrip('#')
        rgb = tuple(int(h[i: i + 2], 16) for i in (0, 2, 4))
        return fgbc(rgb[0], rgb[1], rgb[2])

    def to_hex(self):
        """Returns: color as a hexadecimal string (ex: "#FF0000")"""
        return '#%02X%02X%02X' % (self.red, self.green, self.blue)

    def to_float_rgb_dict(self):
        """Returns: color as a dictionnary with float values"""
        return {"red": self.red / 255.0,
                "green": self.green / 255.0,
                "blue": self.blue / 255.0}

    def to_rgb_dict(self):
        """Returns: color as a dictionnary with int values"""
        return {"red": self.red,
                "green": self.green,
                "blue": self.blue}

    def __eq__(self, other):
        """Equality operator"""
        return (self.red == other.red) and (self.green == other.green) and (self.blue == other.blue)


# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'wowchars'


class SheetConnector:
    """Helper class to use Google Sheets"""
    def __init__(self, sheet_id, dry_run):
        """Constructor

        Args:
            sheet_id (str): ID of the document.
            dry_run (bool): if True, do not modify the document
        """
        self.dry_run = dry_run
        self.credentials = self.get_credentials()

        http = self.credentials.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                        'version=v4')
        self.service = discovery.build('sheets', 'v4', http=http,
                                       discoveryServiceUrl=discoveryUrl)

        self.spreadsheetId = sheet_id

    def check_or_create_sheet(self, sheet_name):
        """Check if the sheet exists in the document, create it otherwise

        Args:
            sheet_name (str): name of the sheet to check
        """
        if self.sheet_exists(sheet_name):
            return

        body = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": sheet_name,
                            # "gridProperties": {
                            #   "rowCount": 20,
                            #   "columnCount": 12
                            # },
                            # "tabColor": {
                            #   "red": 1.0,
                            #   "green": 0.3,
                            #   "blue": 0.4
                            # }
                        }
                    }
                }
            ]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def sheet_exists(self, sheet_name):
        """Check if the sheet exists in the document

        Args:
            sheet_name (str): name of the sheet to check
        """
        return sheet_name in self.get_sheets()

    def delete_sheet(self, sheet_name):
        """Delete the sheet in the spreadsheet

        Args:
            sheet_name (str): name of the sheet to delete
        """
        sheets = self.get_sheets()
        if sheet_name not in sheets:
            return

        body = {
            "requests": [
                {
                    "deleteSheet": {
                        "sheetId": sheets[sheet_name]
                    }
                }
            ]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def get_sheets(self):
        """Get the sheets in the doc

        Returns:
            (dict) keys are names, values are IDs
        """
        sheet_metadata = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        sheets = sheet_metadata.get('sheets', '')
        return {s["properties"]["title"]: s["properties"]["sheetId"] for s in sheets}

    def ensure_headers(self, sheet_name, fieldnames):
        """Ensure the sheet exists and contains the specified headers

        Args:
            sheet_name (str): name of the sheet to check
            sheet_name (str array): headers to check
        """
        self.check_or_create_sheet(sheet_name)
        values, _range = self.get_values(sheet_name + "!1:1")
        g_headers = values[0] if values else []

        appended_headers = []

        # building indexes map and adding extra headers
        g_headers_indexes = {g_h: i for i, g_h in enumerate(g_headers)}
        for field in fieldnames:
            try:
                g_headers_indexes[field] = g_headers.index(field)
            except ValueError:
                g_headers.append(field)
                appended_headers.append(field)
                g_headers_indexes[field] = len(g_headers)

        if(appended_headers):
            logging.info("Adding headers %s", appended_headers)
            update_data = SheetBatchUpdateData()
            update_data.add_data(sheet_name,
                                 column_letter(len(g_headers_indexes) - len(appended_headers)), 1,
                                 column_letter(len(g_headers_indexes) - 1), 1,
                                 [appended_headers])
            self.update_values(update_data)

    def update_values(self, update_data):
        """Update values in the document

        Args:
            update_data (array): data to update, ex: [{"values": [["val1", "val2"]], "range": "sheet2!A2:B2"}]
        """
        self.ensure_no_missing_row_or_column(update_data)

        data = update_data.to_query_data()
        body = {"data": data, "value_input_option": "USER_ENTERED"}
        logging.info("%sUpdating data in Google sheets: %s", ("DRYRUN: " if self.dry_run else ""), data)
        if not self.dry_run:
            self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def ensure_no_missing_row_or_column(self, update_data):
        max_cols = update_data.get_max_columns()

        for mc in max_cols:
            last_col = self.get_last_column(mc)
            last_col_index = column_index(last_col)
            max_col_index = column_index(max_cols[mc])
            if max_col_index > last_col_index:
                sheet_id = self.get_sheet_id(mc)
                self.append_columns(sheet_id, max_col_index - last_col_index)

        max_rows = update_data.get_max_rows()
        for mr in max_rows:
            last_row_index = self.get_last_row(mr)
            max_row_index = max_rows[mr]
            if max_row_index > last_row_index:
                sheet_id = self.get_sheet_id(mr)
                self.append_rows(sheet_id, max_row_index - last_row_index)

    def get_values(self, rangeName):
        """Get values

        Args:
            rangeName (str): range of the values to get. Ex: "sheet2!A2:B2"
        """
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId, range=rangeName).execute()
        return result.get('values', []), result.get('range', None)

    def get_credentials(self, flags=None):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'sheets.googleapis.com-python-wowchars.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else:  # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def get_background_color(self, sheet, column, row):
        """Get background color of a cell

        Args:
            sheet (string): name of the sheet
            column (string): column id of the cell
            row (int): row of the cell

        Returns:
            The background color as a RGBColor object
        """
        ranges = "%s!%s%d" % (sheet, column, row)
        data = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId, ranges=ranges, includeGridData=True).execute()
        v = data["sheets"][0]["data"][0]["rowData"][0]["values"][0]
        if "effectiveFormat" in v:
            return RGBColor.from_float_rgb_dict(v["effectiveFormat"]["backgroundColor"])
        else:
            return RGBColor.from_float_rgb_dict(v["userEnteredFormat"]["backgroundColor"])

    def set_background_color(self, sheet_name, column, row, rgb_color):
        col_i = column_index(column) if type(column) is str else column

        sheets = self.get_sheets()

        body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheets[sheet_name],
                            "startRowIndex": row - 1,
                            "endRowIndex": row,
                            "startColumnIndex": col_i,
                            "endColumnIndex": col_i + 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": rgb_color.to_float_rgb_dict()
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)"
                    }
                },
            ]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def get_sheet_id(self, sheet_name):
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        for _sheet in spreadsheet['sheets']:
            if _sheet['properties']['title'] == sheet_name:
                return _sheet['properties']['sheetId']
        return None

    def append_columns(self, sheet_id, nb_cols):
        logging.debug("Adding %d column(s)" % nb_cols)
        request_body = {
            "requests": [{
                "appendDimension": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "length": nb_cols
                }
            }]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=request_body).execute()

    def append_rows(self, sheet_id, nb_rows):
        logging.debug("Adding %d row(s)" % nb_rows)
        request_body = {
            "requests": [{
                "appendDimension": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "length": nb_rows
                }
            }]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=request_body).execute()

    def get_last_column(self, sheet_name):
        first_line_range = "{sheet}!1:1".format(sheet=sheet_name)
        _, res_range = self.get_values(first_line_range)
        match = re.match(r"(?P<sheet>.*)!(?P<start_col>\w+)(\d+)(:(?P<end_col>\w+)(\d+))?", res_range)
        return match.group("end_col") or match.group("start_col")

    def get_last_row(self, sheet_name):
        first_row_range = "{sheet}!A:A".format(sheet=sheet_name)
        _, res_range = self.get_values(first_row_range)
        match = re.match(r"(?P<sheet>.*)!(\w+)(?P<start_row>\d+)(:(\w+)(?P<end_row>\d+))?", res_range)
        return int(match.group("end_row") or match.group("start_row"))


class SheetSingleUpdateData:
    def __init__(self, sheet_name, start_col, start_row, end_col, end_row, values):
        self.sheet_name = sheet_name
        self.start_col = start_col
        self.end_col = end_col
        self.start_row = start_row
        self.end_row = end_row
        self.values = values

    def to_query_data(self):
        return {
            "values": self.values,
            "range": "%s!%s%d:%s%d" % (self.sheet_name, self.start_col, self.start_row, self.end_col, self.end_row)
        }


class SheetBatchUpdateData(list):
    def add_data(self, sheet_name, start_col, start_row, end_col, end_row, values):
        self.append(SheetSingleUpdateData(sheet_name, start_col, start_row, end_col, end_row, values))

    def get_max_columns(self):
        # grabbing max ranges for each sheet
        max_columns = {}     # {sheet: max column}
        for data in self:
            if data.sheet_name not in max_columns:
                max_columns[data.sheet_name] = data.end_col
            elif data.end_col > max_columns[data.sheet_name]:
                max_columns[data.sheet_name] = data.end_col
        return max_columns

    def get_max_rows(self):
        max_rows = {}  # {sheet: max row}
        for data in self:
            if data.sheet_name not in max_rows:
                max_rows[data.sheet_name] = data.end_row
            elif data.end_row > max_rows[data.sheet_name]:
                max_rows[data.sheet_name] = data.end_row
        return max_rows

    def to_query_data(self):
        return [d.to_query_data() for d in self]


if __name__ == "__main__":
    main()
