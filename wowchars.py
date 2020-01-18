# -*- coding: utf-8 -*-

__version__ = "3.1.2"

"""
TODO: remove hardcoded enchants / gems suggestions
      ==> replace by CSV file
      and/or
      ==> retrieve advices from websites like Noxxic or Icy Veins

TODO: add suboptions to --guild: min level, min ilvl
TODO: [Google Sheets] colorize names with class colors ?
TODO: [Google Sheets] header freeze + color?
TODO: [Google Sheets] automatically add graph(s) ?
"""

import requests
import csv
import argparse
import logging
import httplib2
import os
import string
from time import strftime

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

####################
# URLs
TOKEN_URL       = "https://{zone}.battle.net/oauth/token"
BASE_CHAR_URL   = "https://{zone}.api.blizzard.com/wow/character/{server}/{name}?fields={fields}&access_token={access_token}"
BASE_ACHIEV_URL = "https://{zone}.api.blizzard.com/wow/achievement/{id}?access_token={access_token}"
BASE_ITEM_URL   = "https://{zone}.api.blizzard.com/wow/item/{id}{slash_context}?bl={bonus_list}&access_token={access_token}"
CLASSES_URL     = "https://{zone}.api.blizzard.com/wow/data/character/classes?locale=en_GB&access_token={access_token}"
GUILD_URL       = "https://{zone}.api.blizzard.com/wow/guild/{server}/{name}?fields={fields}&access_token={access_token}"

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
        self.to_fix = {}        # {char, [to fix]}
        self.classnames = {}    # {id, classname}

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
        self.fetch_classes()

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
                logging.warn("no server name, using default server 'voljin'")
                server = "voljin"
            name = serv_and_guildname

        url = GUILD_URL.format(zone=self.zone, access_token=self.access_token, server=server, name=name, fields="members")
        logging.debug(url)
        r = requests.get(url)
        r.raise_for_status()
        body = r.json()
        logging.trace(body)
        members = ["members"]
        guild_chars = []
        for m in members:
            charname = m["character"]["name"]
            level = m["character"]["level"]
            realm = m["character"]["realm"]
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
                logging.warn("no server name, using default server 'voljin'")
                server = "voljin"
            name = serv_and_name

        if self.get_known_char(server, name):
            logging.warn("character '%s' already processed" % (serv_and_name))
            return

        char = CharInfo(server, name)
        try:
            self.fetch_char_base(char, check_gear)
            if not raid:
                self.fetch_char_achievements(char)
                self.fetch_char_professions(char)
        except (ValueError, KeyError, requests.exceptions.HTTPError):
            logging.error("cannot fetch %s/%s", server, name)
            return
        self.characters.append(char)

    def fetch_char_base(self, char, check_gear):
        """Fetch and fill info for the given character: level + items related info

        Args:
            char (CharInfo): the character to fetch
            check_gear (bool): check gear for any missing gem or enchantment
        """
        url = BASE_CHAR_URL.format(zone=self.zone, access_token=self.access_token, server=char.server(), name=char.name(), fields="items")
        logging.debug(url)
        r = requests.get(url)
        r.raise_for_status()
        char_json = r.json()
        logging.trace(char_json)
        char.set_data(H_CLASS, self.classnames[char_json[H_CLASS]])
        char.set_data(H_LVL, str(char_json[H_LVL]))
        items = char_json["items"]
        char.set_data(H_ILVL, str(items["averageItemLevelEquipped"]))

        # Searching the 'azeriteLevel' of the item in the neck slot
        try:
            char.set_data(H_AZERITE_LVL, str(items["neck"]["azeriteItem"]["azeriteLevel"]))
        except (ValueError, KeyError):
            logging.warn("Cannot find azerite level.")
            char.set_data(H_AZERITE_LVL, "0")

        # patch 8.3 introduced a new legendary cloak called 'Ashjra’kamas'
        # Searching the 'rank' of this item.
        try:
            back_item = items["back"]
            if back_item["id"] != 169223:
                raise ValueError("Item is not Ashjra’kamas: %s", back_item["name"])
            char.set_data(H_ASHJRA_KAMAS, back_item["quality"])
        except (ValueError, KeyError):
            logging.warn("Cannot find 'Ashjra’kamas'.")
            char.set_data(H_ASHJRA_KAMAS, "None")

        # Checking gear
        if check_gear:
            total_empty_sockets = 0
            missing_enchants = []
            for slot in sorted(items):
                item = items[slot]
                nb_empty_sockets, missing_enchant = self.check_item_enchants_and_gems(slot, item)
                total_empty_sockets += nb_empty_sockets
                if missing_enchant:
                    missing_enchants.append(slot)

            # specific to my characters
            STAT_ENCHANTS = {
                # "oxyde"   : ("versatility", "agi", "heavy hide"),
                # "ayonis"  : ("haste",       "int", "satyr"),
                # "oxyr"    : ("haste",       "agi", "satyr"),
                # "agoniss" : ("haste",       "int", "satyr"),
                # "palaniss": ("haste",       "str", "satyr"),
                # "odyxe"   : ("mastery",     "agi", "satyr"),
                # "kodyx"   : ("mastery",     "str", "satyr"),
                # "oxymus"  : ("haste",       "int", "satyr"),
                # "monxy"   : ("mastery",     "agi", "satyr"),
                # "oxgrom"  : ("haste",       "str", "satyr"),
                # "oxydhe"  : ("crit",        "agi", "satyr"),
                # "voxy"    : ("mastery",     "agi", "satyr"),
            }

            to_fix = []
            if total_empty_sockets:
                if char.name() in STAT_ENCHANTS:
                    to_fix.extend(["gem " + STAT_ENCHANTS[char.name()][0] for i in range(total_empty_sockets)])
                else:
                    to_fix.append("%d gem(s)" % total_empty_sockets)
            for m in missing_enchants:
                if char.name() in STAT_ENCHANTS:
                    if "finger" in m:
                        to_fix.append("enchant ring " + STAT_ENCHANTS[char.name()][0])
                    elif m == "back":
                        to_fix.append("enchant back " + STAT_ENCHANTS[char.name()][1])
                    elif m == "neck" and STAT_ENCHANTS[char.name()][2]:
                        to_fix.append("enchant neck " + STAT_ENCHANTS[char.name()][2])
                    else:
                        to_fix.append("enchant " + m)
                else:
                    to_fix.append("enchant " + m)

            if to_fix:
                self.to_fix["%s-%s" % (char.name(), char.server())] = to_fix

    def check_item_enchants_and_gems(self, slot, item_dict):
        """Check any missing enchant or gem in the given item

        Args:
            slot (str): slot of the item (ex: head, back, shoulders, neck...)
            item_dict (dict): incomplete description of the item received from
                              the API

        Returns:
            (int, boolean): (number of empty gem slot, true is not enchanted)

        """
        if not isinstance(item_dict, dict) or "id" not in item_dict:
            return 0, False
        logging.debug("Checking enchants and gems of slot: %s", slot)
        item_id = item_dict["id"]
        context = item_dict["context"]
        # Some contexts seem to be invalid in the API, so we do not use them
        if context in ["vendor", "scenario-normal", "quest-reward"]:
            context = ""
        elif context:
            context = "/" + context
        bonus_list = ",".join([str(b) for b in item_dict["bonusLists"]])
        nb_empty_sockets = 0
        missing_enchant = False
        # getting full item description
        try:
            # logging.debug(item_dict)
            url = BASE_ITEM_URL.format(zone=self.zone, access_token=self.access_token, id=item_id,
                                       slash_context=context,
                                       bonus_list=bonus_list)
            logging.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            item = r.json()
            logging.trace(item)
            # checking gem slots & checking with current item state
            nb_sockets = len(item["socketInfo"]) if "socketInfo" in item else 0
            for i in range(nb_sockets):
                if ("gem%d" % i) not in item_dict["tooltipParams"]:
                    nb_empty_sockets += 1
            if nb_empty_sockets:
                logging.debug("Found %d empty socket(s)", nb_empty_sockets)

            # checking enchants
            if slot in ["finger1", "finger2", "mainHand"]:
                if ("enchant") not in item_dict["tooltipParams"]:
                    logging.debug("Detected missing enchantment for this slot !")
                    missing_enchant = True
        except ValueError:
            logging.error("cannot get full item description for %s", item_id)

        return nb_empty_sockets, missing_enchant

    def fetch_char_achievements(self, char):
        """Fetch and fill achievements

        Args:
            char (CharInfo): the character to fetch
        """
        try:
            url = BASE_CHAR_URL.format(zone=self.zone, access_token=self.access_token, server=char.server(), name=char.name(), fields="achievements,quests")
            logging.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            obj = r.json()
            logging.trace(obj)
            achievements = obj["achievements"]
        except ValueError:
            logging.warn("cannot retrieve achievements for %s/%s", char.server(), char.name())

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
            url = BASE_CHAR_URL.format(zone=self.zone, access_token=self.access_token, server=char.server(), name=char.name(), fields="professions")
            logging.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            obj = r.json()
            logging.trace(obj)
            professions = obj["professions"]["primary"]
        except ValueError:
            logging.warn("cannot retrieve professions for %s/%s", char.server(), char.name())

        count = 0
        for profession in professions:
            if profession["name"].startswith("Kul Tiran"):
                count += 1
                char.set_data("BfA profession %d" % count, "%s: %d" % (profession["name"].replace("Kul Tiran", "BfA"), profession["rank"]))

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

    def fetch_classes(self):
        """Fetch the 'class id' to 'name' mapping"""
        print("======================================================")
        print("Fetching classes")
        url = CLASSES_URL.format(zone=self.zone, access_token=self.access_token)
        logging.debug(url)
        r = requests.get(url)
        r.raise_for_status()
        body = r.json()
        logging.trace(body)
        classes = body["classes"]
        for c in classes:
            cid = int(c["id"])
            name = c["name"]
            self.classnames[cid] = name
            logging.info("%2d: %s", cid, name)

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
        print("/!\\ %d character(s) to fix!" % len(self.to_fix))
        for c in self.to_fix:
            print("%s: %s" % (c, ", ".join(self.to_fix[c])))
        print("---------------------------")
        to_create = {}
        for c in self.to_fix:
            for tc in self.to_fix[c]:
                if tc not in to_create:
                    to_create[tc] = [1, set()]
                else:
                    to_create[tc][0] += 1
                to_create[tc][1].add(c)

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

        sheet_values = sc.get_values(SUMMARY + "!A:Z")
        headers = sheet_values[0]
        h_indexes = {h: i for i, h in enumerate(headers)}
        values = sheet_values[1:]

        update_data = []  # cell values to update
        to_colorize = []  # cells to colorize (when adding new character(s))

        # updating / adding characters info
        for r in sorted(self.characters, key=lambda x: x[H_ILVL], reverse=True):
            char_index = None
            # checking if character is already known
            for i, g_line in enumerate(values):
                if g_line[h_indexes[H_SERVER]] == r[H_SERVER] and g_line[h_indexes[H_NAME]] == r[H_NAME]:
                    char_index = i
                    break

            if char_index is not None:
                g_row = values[char_index]
                for i, h in enumerate(headers):
                    if (h in fieldnames) and (h in r) and r[h] and ((i >= len(g_row)) or (r[h] != g_row[i])):
                        cell_id = "%s%d" % (column_letter(i), char_index + 2)
                        update_data.append({
                            "values": [[r[h]]],
                            "range": SUMMARY + "!%s:%s" % (cell_id, cell_id),
                        })
            else:
                line = [(r[h] if h in r else None) for h in headers]
                values.append(line)
                row_index = len(values) + 1
                update_data.append({
                    "values": [line],
                    "range": SUMMARY + "!A{row}:Z{row}".format(row=row_index),
                })
                to_colorize.append((h_indexes[H_NAME], row_index, r.get_hex_color()))

        if update_data:
            sc.update_values(update_data)
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
        names = [r[H_NAME] for r in sorted(self.characters, key=lambda x:x[H_NAME])]

        update_data = []
        today = strftime("%Y-%m-%d")

        for s in [H_LVL, H_ILVL]:
            v_dict = {r[H_NAME]: r[s] for r in sorted(self.characters, key=lambda x: x[H_NAME])}
            sc.ensure_headers(s, [H_DATE] + names)
            sheet_values = sc.get_values("%s!A:Z" % s)
            headers = sheet_values[0]
            h_indexes = {h: i for i, h in enumerate(headers)}
            update_needed = False
            last_update_today = (len(sheet_values) > 1) and (sheet_values[-1][h_indexes[H_DATE]] == today)
            if len(sheet_values) <= 1:
                update_needed = True
            else:
                values = sheet_values[1:]
                last_line = values[-1]
                for name in sorted(v_dict):
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
            update_data.append({
                "values": [line],
                "range": "%s!A%d:Z%d" % (s, line_nb, line_nb),
            })

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

    def check_or_create_sheet(self, sheetName):
        """Check if the sheet exists in the document, create it otherwise

        Args:
            sheetName (str): name of the sheet to check
        """
        if self.sheet_exists(sheetName):
            return

        body = {
            "requests": [
                {
                    "addSheet": {
                        "properties": {
                            "title": sheetName,
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

    def sheet_exists(self, sheetName):
        """Check if the sheet exists in the document

        Args:
            sheetName (str): name of the sheet to check
        """
        return sheetName in self.get_sheets()

    def delete_sheet(self, sheetName):
        """Delete the sheet in the spreadsheet

        Args:
            sheetName (str): name of the sheet to delete
        """
        sheets = self.get_sheets()
        if sheetName not in sheets:
            return

        body = {
            "requests": [
                {
                    "deleteSheet": {
                        "sheetId": sheets[sheetName]
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

    def ensure_headers(self, sheetName, fieldnames):
        """Ensure the sheet exists and contains the specified headers

        Args:
            sheetName (str): name of the sheet to check
            sheetName (str array): headers to check
        """
        self.check_or_create_sheet(sheetName)
        values = self.get_values(sheetName + "!1:1")
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
            range_name = "%s!%s1:AZ1" % (sheetName, column_letter(len(g_headers_indexes) - len(appended_headers)))
            logging.info("Adding headers in %s => %s", range_name, appended_headers)

            update_data = [{
                "values": [appended_headers],
                "range": range_name,
            }]
            self.update_values(update_data)

    def update_values(self, update_data):
        """Update values in the document

        Args:
            update_data (array): data to update, ex: [{"values": [["val1", "val2"]], "range": "sheet2!A2:B2"}]
        """
        body = {"data": update_data, "value_input_option": "USER_ENTERED"}
        logging.info("%sUpdating data in Google sheets: %s", ("DRYRUN: " if self.dry_run else ""), update_data)
        if not self.dry_run:
            self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def get_values(self, rangeName):
        """Get values

        Args:
            rangeName (str): range of the values to get. Ex: "sheet2!A2:B2"
        """
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId, range=rangeName).execute()
        return result.get('values', [])

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


if __name__ == "__main__":
    main()
