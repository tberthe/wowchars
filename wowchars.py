# -*- coding: utf-8 -*-

__version__="1.0"

"""
TODO: option to set the default server
TODO: remove hardcoded enchants / gems suggestions
      ==> replace by CSV file
      and/or
      ==> retrieve advises from websites like Noxxic or Icy Veins
TODO: [Google Sheets] colorize names with class colors ?
TODO: [Google Sheets] header freeze + color?
TODO: [Google Sheets] automatically add graph(s) ?
"""

import requests
import csv
import argparse
import re
import logging
import httplib2
import os
import string
from time import strftime

import googleapiclient
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

logger = logging.getLogger('wowchars')

##########
# URLs
BASE_CHAR_URL   = "https://eu.api.battle.net/wow/character/{server}/{name}?fields={fields}&apikey={apikey}"
BASE_ACHIEV_URL = "https://eu.api.battle.net/wow/achievement/{id}?apikey={apikey}"
BASE_ITEM_URL   = "https://eu.api.battle.net/wow/item/{id}{slash_context}?bl={bonus_list}&apikey={apikey}"

##########
# Headers
H_DATE        = "date"
H_SERVER      = "server"
H_NAME        = "name"
H_LVL         = "level"
H_ILVL        = "ilvl"
H_3RD_RELIC   = "3rd relic"
H_CLASS_CHAMPIONS = "class champions"
H_CLASS_MOUNT = "class mount"
H_LEG_ITEMS   = "leg. items"

# Achievements
A_POWER_UNBOUND           = 11609
A_CHAMPIONS_OF_LEGIONFALL = 11846
A_BREACHING_THE_TOMB      = 11546
#A_ROSTER_OF_CHAMPIONS     = 11220 #checked directly via criteria to check the 9 champions
A_LEGENDARY_RESEARCH      = 11223
#A_POWER_ASCENDED          = 11772
ACHIEVEMENTS         = (A_POWER_UNBOUND, A_CHAMPIONS_OF_LEGIONFALL, A_LEGENDARY_RESEARCH,)
STEPPED_ACHIEVEMENTS = (A_BREACHING_THE_TOMB,)


def main():
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument("-b", "--bnet-api-key", help="Key to Blizzard's Battle.net API", required=True)
    parser.add_argument("-o", "--output", help="Output CSV file", required=False)
    parser.add_argument("-c", "--char", help="Check character (server:charname)", action="append", default=[], required=True)
    parser.add_argument("-s", "--summary", help="Display summary", action="store_true")
    parser.add_argument("-v", "--verbosity", action="count", default=0, help="increase output verbosity")
    parser.add_argument("-g", "--google-sheet", help="ID of the Google sheet where the results will be saved")
    parser.add_argument("-d", "--dry-run", action="store_true", help="does not update target output")
    parser.add_argument("--check-gear", action="store_true", help="inspect gear for legendaries or any missing gem/enchantment")
    parser.add_argument("--default-server", help="Default server when not given with the '-c' option", default=None)
    parser.add_argument('--version', action='version', version=__version__)
    args = parser.parse_args()

    set_logger(args.verbosity)

    ce = CharactersExtractor(args.bnet_api_key)
    ce.run(args.char,
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
        super(dict, self).__init__()
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
        logger.info("%s: %s", key, value)


class CharactersExtractor:
    """Class processing World Of Warcraft characters:
    - extract and process data from Blizzard API
    - save in CSV and/or export to a Google Sheets document"""

    def __init__(self, bnet_api_key):
        """Contructor

        Args:
            bnet_api_key (str): Battle.net API key
        """
        self.bnet_key = bnet_api_key
        self.achievements = [] # achievement details
        self.characters = []   # fetched characters
        self.to_fix = {}       # {char, [to fix]}

    def run(self, chars, csv_output, summary,
            check_gear, google_sheet_id, dry_run,
            default_server):
        """main function

        Args:
            chars (str array): list of characters to process
            csv_output (str): if not None, save results in the CSV file
            summary (bool): print a results' sumamry
            check_gear (bool): check gear for any missing gem or enchantment
            google_sheet_id (str): if not None, save results in Google Sheets
            dry_run (boolean): does not modify the Google Sheets document
            default_server (string): default server if not given in 'chars'
        """
        self.fetch_achievements_details()

        for c in chars:
            self.fetch_char(c, default_server, check_gear)

        if csv_output:
            self.save_csv(csv_output)

        if summary:
            self.display_summary()

        if check_gear:
            self.display_gear_to_fix()

        if google_sheet_id:
            self.save_summary_in_google_sheets(google_sheet_id, dry_run)
            self.save_in_google_sheets_v2(google_sheet_id, dry_run)

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

    def fetch_char(self, serv_and_name, default_server=None, check_gear=False):
        """Fetch and register a character from Blizzard's API.

        Args:
            serv_and_name (str): server and name of the character to fetch.
                                 Expected format is "server:name". Using only
                                 the name is supported but will produce a warning.
            check_gear (bool): check gear for any missing gem or enchantment
            default_server (string): default server if not given in 'serv_and_name'
        """
        print("======================================================")
        print("Processing: %s" % (serv_and_name))
        server = None
        name = None
        try:
            server, name = serv_and_name.split(":")
        except ValueError:
            if default_server:
                logger.info("no server name, using default server '%s'", default_server)
                server = default_server
            else:
                logger.warn("no server name, using default server 'voljin'")
                server = "voljin"
            name = serv_and_name

        if self.get_known_char(server, name):
            logger.warn("character '%s' already processed" % (serv_and_name))
            return

        char = CharInfo(server, name)
        try:
            self.fetch_char_base(char, check_gear)
            self.fetch_char_achievements(char)
            self.fetch_char_professions(char)
        except (ValueError, KeyError):
            logger.error("cannot fetch %s/%s", server, name)
            return
        self.characters.append(char)

    def fetch_char_base(self, char, check_gear):
        """Fetch and fill info for the given character: level + items related info

        Args:
            char (CharInfo): the character to fetch
            check_gear (bool): check gear for any missing gem or enchantment
        """
        url = BASE_CHAR_URL.format(apikey=self.bnet_key, server=char.server(), name=char.name(), fields="items")
        logger.debug(url)
        r = requests.get(url)
        r.raise_for_status()
        char_json = r.json()
        char.set_data(H_LVL, str(char_json[H_LVL]))
        items = char_json["items"]
        char.set_data(H_ILVL, str(items["averageItemLevelEquipped"]))

        leg_info_array = []
        for slot in sorted(items):
            item = items[slot]
            if isinstance(item, dict) and "quality" in item and item["quality"] == 5:
                leg_info_array.append(str(item["itemLevel"]))
                logger.debug("Found legendary item: [%d] %s" % (item["itemLevel"], items[slot]["name"]))

        if not leg_info_array:
            leg_info_array = ["NO"]

        char.set_data(H_LEG_ITEMS, ("+").join(leg_info_array))

        if check_gear:
            total_empty_sockets = 0
            missing_enchants = []
            leg_info = {}
            for slot in sorted(items):
                item = items[slot]
                nb_empty_sockets, missing_enchant = self.check_item_enchants_and_gems(slot, item)
                total_empty_sockets += nb_empty_sockets
                if missing_enchant:
                    missing_enchants.append(slot)

            # specific to my characters
            STAT_ENCHANTS = {
                "oxyde"   : ("versatility", "agi", "heavy hide"),
                "ayonis"  : ("haste",       "int", "satyr"),
                "oxyr"    : ("haste",       "agi", "satyr"),
                "agoniss" : ("haste",       "int", "satyr"),
                "palaniss": ("haste",       "str", "satyr"),
                "odyxe"   : ("mastery",     "agi", "satyr"),
                "kodyx"   : ("mastery",     "str", "satyr"),
                "oxymus"  : ("haste",       "int", "satyr"),
                "monxy"   : ("mastery",     "agi", "satyr"),
                "oxgrom"  : ("haste",       "str", "satyr"),
                "oxydhe"  : ("crit",        "agi", "satyr"),
                "voxy"    : ("mastery",     "agi", "satyr"),
            }

            to_fix = []
            if total_empty_sockets:
                if char.name() in STAT_ENCHANTS:
                    to_fix.extend(["gem " + STAT_ENCHANTS[char.name()][0] for i in range(total_empty_sockets)])
                else:
                    to_fix.append("%d gem(s)"%total_empty_sockets)
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
                self.to_fix["%s-%s"%(char.name(), char.server())] = to_fix

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
        logger.debug("Checking enchants and gems of slot: %s", slot)
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
        #getting full item description
        try:
            #logger.debug(item_dict)
            url = BASE_ITEM_URL.format(apikey=self.bnet_key, id=item_id,
                                       slash_context=context,
                                       bonus_list=bonus_list)
            logger.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            item = r.json()

            #checking gem slots & checking with current item state
            nb_sockets = len(item["socketInfo"]) if "socketInfo" in item else 0
            for i in range(nb_sockets):
                if ("gem%d"%i) not in item_dict["tooltipParams"]:
                   nb_empty_sockets += 1
            if nb_empty_sockets:
                logger.debug("Found %d empty socket(s)", nb_empty_sockets)

            #checking enchants
            if slot in ["finger1", "finger2", "back", "neck"]:
                if ("enchant") not in item_dict["tooltipParams"]:
                    logger.debug("Detected missing enchantment for this slot !")
                    missing_enchant = True
        except ValueError:
            logger.error("cannot get full item description for %s", item_id)

        return nb_empty_sockets, missing_enchant

    def fetch_char_achievements(self, char):
        """Fetch and fill achievements

        Args:
            char (CharInfo): the character to fetch
        """
        try:
            url = BASE_CHAR_URL.format(apikey=self.bnet_key, server=char.server(), name=char.name(), fields="achievements,quests")
            logger.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            obj = r.json()
            achievements = obj["achievements"]
            quests = obj["quests"]
        except ValueError:
            logger.warn("cannot retrieve achievements for %s/%s", server, name)

        for ach_desc in self.achievements:
            ach_id = ach_desc["id"]
            title = ach_desc["title"]
            if ach_id in STEPPED_ACHIEVEMENTS:
                char.set_data(title, self.check_stepped_achievement(ach_desc, ach_id, achievements))
            else:
                char.set_data(title, self.check_achievement(ach_desc, ach_id, achievements))

        # 3rd relic
        for relic_quest_id in (43412, 43425, 43409, 43415,
                               43407, 43422, 43423, 43414,
                               43420, 43418, 43359, 43424):
            if relic_quest_id in quests:
                char.set_data(H_3RD_RELIC, "OK")
                break
        else:
            char.set_data(H_3RD_RELIC, "")

        # all class order champions
        count = 0
        try:
            crit_index = achievements["criteria"].index(33142)
            count = achievements["criteriaQuantity"][crit_index]
        except ValueError:
            pass

        char.set_data(H_CLASS_CHAMPIONS, "%d/9" % count)

        # class mount
        # http://fr.wowhead.com/item=142231/renes-putrefiees-de-vainqueur-couvepeste#english-comments
        for mount_quest_id in (46207, 45770, 46337, 46178, 46089,
                               45789, 46813, 46792, 45354,
                               46243, 46350, 46319, 46334):
            if mount_quest_id in quests:
                char.set_data(H_CLASS_MOUNT, "OK")
                break
        else:
            char.set_data(H_CLASS_MOUNT, "")

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
            logger.debug("char has ach: %s", ach_desc)
            for criteria in ach_desc["criteria"]:
                crit_ok = False
                try:
                    crit_index = char_achievements["criteria"].index(criteria["id"])
                    if char_achievements["criteriaQuantity"][crit_index] >= criteria["max"]:
                        ach_ok = True
                        crit_ok = True
                except ValueError:
                    pass
                logger.debug("Criteria %s ==> %s", criteria, crit_ok)
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
            logger.debug("char has ach: %s", ach_desc)
            for criteria in ach_desc["criteria"]:
                if criteria["id"] in char_achievements["criteria"]:
                    logger.debug("Criteria '%s' ==> OK", criteria["description"])
                    count += 1
                else:
                    logger.debug("Criteria '%s' ==> NOK", criteria["description"])
                    next_step = (": " + criteria["description"]) if not next_step else next_step
        return "%d/%d%s" % (count, len(ach_desc["criteria"]), next_step)

    def fetch_char_professions(self, char):
        """Fetch and fill professions

        Args:
            char (CharInfo): the character to fetch
        """
        try:
            url = BASE_CHAR_URL.format(apikey=self.bnet_key, server=char.server(), name=char.name(), fields="professions")
            logger.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            obj = r.json()
            professions = obj["professions"]["primary"]
        except ValueError:
            logger.warn("cannot retrieve professions for %s/%s", server, name)

        char["profession 1"] = ""
        char["profession 2"] = ""
        count = 0
        for profession in professions:
            count += 1
            char.set_data("profession %d"%count, "%3d: %s" % (profession["rank"], profession["name"]))

    def fetch_achievements_details(self):
        print("======================================================")
        print("Fetching achievements details")
        for a_id in ACHIEVEMENTS + STEPPED_ACHIEVEMENTS:
            url = BASE_ACHIEV_URL.format(apikey=self.bnet_key, id=a_id)
            logger.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            ach = r.json()
            logger.info("%6d: %s", ach["id"], ach["title"])
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
        raise ValueError("Cannot retrieve achievement %d"%ach_id)

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
        fieldnames=[H_SERVER, H_NAME, H_ILVL, H_3RD_RELIC, self.get_achievement_title(A_POWER_UNBOUND)]
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
        #computing width of each column
        widths = {f:len(f) for f in fieldnames}
        for f in fieldnames:
            for char in self.characters:
                if f in char:
                    widths[f] = max(widths[f], len(str(char[f])))

        #printing fieldnames
        line = []
        for f in fieldnames:
            v = "%-" + str(widths[f]) + "s"
            line.append(v % f)
        print((", ").join(line))

        #printing rows
        for char in sorted(self.characters, key=lambda x:x[H_ILVL], reverse=True):
            line = []
            for f in fieldnames:
                v = "%-" + str(widths[f]) + "s"
                line.append(v % char[f])
            print((", ").join(line))

    def display_gear_to_fix(self):
        """Display gear to fix: missing gems or enchantments"""
        print("======================================================")
        print("/!\ %d character(s) to fix!" % len(self.to_fix))
        for c in self.to_fix:
            print ("%s: %s"%(c, ", ".join(self.to_fix[c])))
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

        sc = SheetConnector(google_sheet_id, dry_run)
        sc.check_or_create_sheet("summary")
        fieldnames = self.get_ordered_fieldnames()
        sc.ensure_headers("summary", fieldnames)

        sheet_values = sc.get_values("summary!A:Z")
        headers = sheet_values[0]
        h_indexes = {h:i for i, h in enumerate(headers)}
        values = sheet_values[1:]

        update_data = []

        # updating / adding characters info
        for r in sorted(self.characters, key=lambda x:x[H_ILVL], reverse=True):
            char_index = None
            # checking if character is already known
            for i, g_line in enumerate(values):
                if g_line[h_indexes[H_SERVER]] == r[H_SERVER] and g_line[h_indexes[H_NAME]] == r[H_NAME]:
                    char_index = i
                    break

            if char_index is not None:
                g_row = values[char_index]
                for i, h in enumerate(headers):
                    if (h in fieldnames) and r[h] and ((i >= len(g_row)) or (r[h] != g_row[i])):
                        cell_id = "%s%d" %(column_letter(i), char_index+2)
                        update_data.append({
                            "values": [[r[h]]],
                            "range": "%s:%s"%(cell_id, cell_id),
                        })
            else:
                line = [(r[h] if h in fieldnames else None) for h in headers]
                values.append(line)
                update_data.append({
                    "values": [line],
                    "range": "A{row}:Z{row}".format(row=len(values)+1),
                })

        if update_data:
            sc.update_values(update_data)
        else:
            print("Nothing to update")

    def save_in_google_sheets_v2(self, google_sheet_id, dry_run):
        """Save level and ilvl in Google Sheets in separated Sheets

        Args:
            google_sheet_id (str): the ID of the document
            dry_run (bool): if True, do not modify the document
        """
        print("======================================================")
        print("Synching ilvl/level in Google Sheets")

        sc = SheetConnector(google_sheet_id, dry_run)
        names = [ r[H_NAME] for r in sorted(self.characters, key=lambda x:x[H_NAME])]

        update_data = []
        today = strftime("%Y-%m-%d")

        for s in [H_LVL, H_ILVL]:
            v_dict = {r[H_NAME]:r[s] for r in sorted(self.characters, key=lambda x:x[H_NAME])}
            sc.ensure_headers(s, [H_DATE]+names)
            sheet_values = sc.get_values("%s!A:Z"%s)
            headers = sheet_values[0]
            h_indexes = {h:i for i, h in enumerate(headers)}
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


def set_logger(verbosity):
    """Initialize and set the logger

    Args:
        verbosity (int): 0 -> default, 1 -> info, 2 -> debug
    """
    log_level = logging.WARNING - (verbosity * 10)
    logger.setLevel(log_level)
    # create console handler and set level to debug
    ch = logging.StreamHandler()
    ch.setLevel(log_level)
    # create formatter
    formatter = logging.Formatter('[%(levelname)s] %(message)s')

    # add formatter to ch
    ch.setFormatter(formatter)

    # add ch to logger
    logger.addHandler(ch)
    # logging.basicConfig(format='[%(levelname)s] %(message)s', level=logging.WARNING - (self.args.verbosity * 10))


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
        i = (i-j) // len(string.ascii_uppercase) - 1
        if i < 0:
            break
    return res


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
        return {s["properties"]["title"]:s["properties"]["sheetId"] for s in sheets}

    def ensure_headers(self, sheetName, fieldnames):
        """Ensure the sheet exists and contains the specified headers

        Args:
            sheetName (str): name of the sheet to check
            sheetName (str array): headers to check
        """
        self.check_or_create_sheet(sheetName)
        values = self.get_values(sheetName+"!1:1")
        g_headers = values[0] if values else []

        appended_headers = []

        # building indexes map and adding extra headers
        g_headers_indexes = {g_h:i for i, g_h in enumerate(g_headers)}
        for field in fieldnames:
            try:
                g_headers_indexes[field] = g_headers.index(field)
            except ValueError:
                g_headers.append(field)
                appended_headers.append(field)
                g_headers_indexes[field] = len(g_headers)


        if(appended_headers):
            range_name = "%s!%s1:Z1" % (sheetName, column_letter(len(g_headers_indexes) - len(appended_headers)))
            logger.info("Adding headers in %s => %s", range_name, appended_headers)

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
        body = { "data": update_data, "value_input_option": "USER_ENTERED" }
        logger.info("%sUpdating data in Google sheets: %s", ("DRYRUN: " if self.dry_run else ""), update_data)
        if not self.dry_run:
            response = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def get_values(self, rangeName):
        """Get values

        Args:
            rangeName (str): range of the values to get. Ex: "sheet2!A2:B2"
        """
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId, range=rangeName).execute()
        return result.get('values', [])

    def get_credentials(flags):
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
            else: # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials


if __name__ == "__main__":
    main()
