import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../../")))

import wowchars

class BlizzardTestHelper(object):
    def __init__(self):
        self.extractor = None

    def init_test(self, api_key):
        if self.extractor:
            print("*WARN* Already initialized")
            return
        self.extractor = wowchars.CharactersExtractor(api_key)

    def get_level(self, server, name):
        char = wowchars.CharInfo(server, name)
        self.extractor.fetch_char_base(char, False)
        return char.level()

    def get_ilevel(self, server, name):
        char = wowchars.CharInfo(server, name)
        self.extractor.fetch_char_base(char, False)
        return char.ilevel()

    def get_legendaries_info(self, server, name):
        char = wowchars.CharInfo(server, name)
        self.extractor.fetch_char_base(char, False)
        return char[wowchars.H_LEG_ITEMS]

    def load_achievements(self):
        self.extractor.fetch_achievements_details()

    def should_know_achievement(self, achievement_id):
        self.extractor.get_achievement_title(achievement_id)