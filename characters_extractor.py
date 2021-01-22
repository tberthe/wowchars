# -*- coding: utf-8 -*-

from character_info import CharacterInfo
from character_info import SCD_LVL as H_LVL
from character_info import SCD_ILVL as H_ILVL
from character_info import SCD_NAME as H_NAME
from character_info import SCD_SERVER as H_SERVER
from column_utils import column_letter
import csv
from google_sheets_connector import GoogleSheetsConnector, SheetBatchUpdateData
import logging
import requests
from rgb_color import RGBColor
from time import strftime

####################
# URLs
TOKEN_URL       = "https://{zone}.battle.net/oauth/token"
CHAR_URL        = "https://{zone}.api.blizzard.com/profile/wow/character/{server}/{name}?namespace=profile-{zone}&locale=en_GB&access_token={access_token}"
EQUIPMENT_URL   = "https://{zone}.api.blizzard.com/profile/wow/character/{server}/{name}/equipment?namespace=profile-{zone}&locale=en_GB&access_token={access_token}"
PROFESSIONS_URL = "https://{zone}.api.blizzard.com/profile/wow/character/{server}/{name}/professions?namespace=profile-{zone}&locale=en_GB&access_token={access_token}"
GUILD_URL       = "https://{zone}.api.blizzard.com/data/wow/guild/{server}/{name}/roster?namespace=profile-{zone}&locale=en_GB&access_token={access_token}"

# Headers
H_DATE         = "date"
H_SL_LEG       = "SL legendary"


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
            a CharacterInfo object or None if not found
        """
        for r in self.characters:
            if (r.server == server) and (r.name == name):
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
            if level >= 45:
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

        char = CharacterInfo(server, name)
        try:
            self.fetch_char_base(char)
            self.fetch_char_items(char, check_gear)
            if not raid:
                self.fetch_char_professions(char)
        except (ValueError, KeyError, requests.exceptions.HTTPError) as e:
            logging.debug(e)
            logging.error("cannot fetch %s/%s", server, name)
            return
        self.characters.append(char)

    def fetch_char_base(self, char):
        """Fetch and fill info for the given character: level + items related info

        Args:
            char (CharacterInfo): the character to fetch
        """
        url = CHAR_URL.format(zone=self.zone, access_token=self.access_token, server=char.server, name=char.name.lower())
        logging.debug(url)
        r = requests.get(url)
        r.raise_for_status()
        char_json = r.json()
        logging.trace(char_json)
        char.classname = char_json["character_class"]["name"]
        char.level = char_json["level"]
        ilvl = char_json["average_item_level"]
        # If the character has not been connected since patch 9.0, the ilvl does not take into account the ilevel squish applied in this patch.
        # So in a first solution we set it to 0
        if (char.level <= 50) and (ilvl > 140):
            ilvl = 0
        char.ilevel = ilvl

    def fetch_char_items(self, char, check_gear):
        """Fetch and fill info for the items of the given character

        Args:
            char (CharacterInfo): the character to fetch
            check_gear (bool): check gear for any missing gem or enchantment
        """
        url = EQUIPMENT_URL.format(zone=self.zone, access_token=self.access_token, server=char.server, name=char.name.lower())
        logging.debug(url)
        r = requests.get(url)
        r.raise_for_status()
        items = r.json()["equipped_items"]
        logging.trace(items)

        # Search ShadowLands legendary item
        sl_leg = self.get_shadowlands_legendary_item(items)
        if sl_leg:
            ilvl = sl_leg["level"]["value"]
            name = sl_leg["name"]
            char.set_data(H_SL_LEG, "%s (%d)" % (name, ilvl))
        else:
            logging.warning("Cannot find any SL legendary.")
            char.set_data(H_SL_LEG, "NA")

        # Checking gear
        if check_gear:
            total_empty_sockets = 0
            missing_enchants = []
            for item in items:
                nb_empty_sockets, missing_enchant = self.check_item_enchants_and_gems(item)
                total_empty_sockets += nb_empty_sockets
                if missing_enchant:
                    missing_enchants.append(item["inventory_type"]["name"])

            char.missing_gems = total_empty_sockets
            char.missing_enchants = missing_enchants

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

    def fetch_char_professions(self, char):
        """Fetch and fill ShadowLands professions

        Args:
            char (CharacterInfo): the character to fetch
        """
        try:
            url = PROFESSIONS_URL.format(zone=self.zone, access_token=self.access_token, server=char.server, name=char.name.lower())
            logging.debug(url)
            r = requests.get(url)
            r.raise_for_status()
            obj = r.json()
            logging.trace(obj)
            professions = obj["primaries"]
        except ValueError:
            logging.warning("cannot retrieve professions for %s/%s", char.server, char.name)

        count = 0
        for profession in professions:
            for tier in profession["tiers"]:
                tier_name = tier["tier"]["name"]
                if tier_name.startswith("Shadowlands "):
                    tier_name = tier_name.replace("Shadowlands ", "")
                    count += 1
                    char.set_data("SL profession %d" % count, "%s: %d" % (tier_name, tier["skill_points"]))

    def get_item(self, items, slot_type):
        """Retrieve item of the selected slot.

        Args:
            items (array): list of items of a character
            slot_type (str): searched slot

        Returns:
            the item of the given slot
        """
        for item in items:
            if item["slot"]["type"] == slot_type:
                return item
        raise ValueError("No item of type %s" % slot_type)

    def get_shadowlands_legendary_item(self, items):
        """Retrieve the equipped shadowlands legendary item.

        Args:
            items (array): list of items of a character

        Returns:
            The found legendary item
        """
        for item in items:
            if ((item["quality"]["type"] == "LEGENDARY") and (item["requirements"]["level"]["value"] == 60)):
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
        fieldnames = list(CharacterInfo.common_headers)
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
        for char in sorted(self.characters, key=lambda x: x.ilevel, reverse=True):
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
            print("%s: %s" % (c.name, missing))

        print("---------------------------")
        to_create = {}  # {enchant : [nb, set of chars]}
        for c in chars_to_fix:
            nb_gems = c.nb_missing_gems
            if nb_gems:
                if "Gem" not in to_create:
                    to_create["Gem"] = [1, set()]
                else:
                    to_create["Gem"][0] += 1
                to_create["Gem"][1].add("%s(%d)" % (c.name, nb_gems))

            for enchant in c.missing_enchants:
                if enchant not in to_create:
                    to_create[enchant] = [1, set()]
                else:
                    to_create[enchant][0] += 1
                to_create[enchant][1].add(c.name)

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

        sc = GoogleSheetsConnector(google_sheet_id, dry_run)
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
        for r in sorted(self.characters, key=lambda x: x.ilevel, reverse=True):
            char_index = None
            # checking if character is already known
            for i, g_line in enumerate(values):
                if g_line[h_indexes[H_SERVER]] == r.server:
                    g_name = g_line[h_indexes[H_NAME]]
                    if (g_name == r.name) or (g_name == r.name.lower()):
                        char_index = i
                        break

            if char_index is not None:
                g_row = values[char_index]
                for i, h in enumerate(headers):
                    if (h in fieldnames) and (h in r) and r[h] and ((i >= len(g_row)) or (str(r[h]) != g_row[i])):
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

        sc = GoogleSheetsConnector(google_sheet_id, dry_run)
        names = [c.name for c in sorted(self.characters, key=lambda x:x.name)]

        update_data = SheetBatchUpdateData()
        today = strftime("%Y-%m-%d")

        for s in [H_LVL, H_ILVL]:
            v_dict = {r.name: r[s] for r in sorted(self.characters, key=lambda x: x.name)}
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
