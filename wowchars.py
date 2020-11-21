# -*- coding: utf-8 -*-

__version__ = "4.0.0"

"""
TODO: Add missing docstrings
TODO: fix support of achievements
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
from characters_extractor import CharactersExtractor
import logging
from oauth2client import tools


def main():
    parser = argparse.ArgumentParser(parents=[tools.argparser])
    parser.add_argument("--blizzard-client-id", help="Client ID of Blizzard's Battle.net API", required=True)
    parser.add_argument("--blizzard-client-secret", help="Token to Blizzard's Battle.net API", required=True)
    parser.add_argument("-o", "--output", help="Output CSV file", required=False)
    parser.add_argument("-c", "--char", help="Check character (server:charname)", action="append", default=[], required=False)
    parser.add_argument("--guild", help="Check characters from given GUILD with minimum level of 45")
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


if __name__ == "__main__":
    main()
