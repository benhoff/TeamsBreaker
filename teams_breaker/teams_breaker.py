import argparse
import csv
import sys
import time
from pathlib import Path

import chevron
import teams_api
import yaml
from loguru import logger
from prettytable import PrettyTable
from teams_user import TeamsUser


def build_argparser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-u", "--username", dest="username", type=str, required=True, help="Username for authentication"
    )
    parser.add_argument(
        "-p", "--password", dest="password", type=str, required=True, help="Password for authentication"
    )
    parser.add_argument(
        "-l",
        "--list",
        dest="list",
        type=str,
        default=None,
        required=True,
        help="Full path to a file containing target emails. One per line.",
    )
    return parser


def create_thread_by_method(method):
    sw = {"closed_chat": teams_api.chat_create_closed_chat, "meeting": teams_api.chat_create_meeting}
    return sw.get(method)


def print_users_status(bearer_token, users):
    statuses = []

    for email in users:
        user = TeamsUser(bearer_token, email)
        ustatus = user.get_status()
        if ustatus:
            time.sleep(0.33)
            mri = ustatus["mri"]
            user_info = user.check_teams_presence(mri)[0]
            presence = user_info["presence"]
            availability = presence["availability"]
            # color = availability_map[availability]
            # status = [email, parse(f"<{color}>{availability}</{color}>"), presence["deviceType"], mri]
            status = [email, availability, presence.get("deviceType")]
            statuses.append(status)
        else:
            statuses.append([email, "Could not read", "Could not read"])

    table = PrettyTable()
    table.field_names = ["Email", "Availability", "Device Type"]
    table.add_rows(statuses)
    print(table)

    with open("user_status.csv", "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(statuses)


def convert_str_to_html(s):
    lines = s.splitlines()
    return "\n".join([f"<p>{line if len(line) != 0 else '&nbsp;'}</p>" for line in lines])


if __name__ == "__main__":
    args = build_argparser().parse_args()

    configuration = None

    logger.info("Running in enumeration mode.")


    to_upload = None

    users = None
    list_path = None
    list_path = Path(args.list)

    if list_path.is_file():
        users = list_path.read_text().strip().splitlines()
    else:
        logger.error("User list file does not exist or it is not a file!")
        sys.exit(1)

    bearer_token, skype_token, sharepoint_token, sender_info = teams_api.authenticate(
        args.username, args.password, False
    )

    sender_drive = args.username.replace("@", "_").replace(".", "_").lower()

    method = "closed_chat"

    print_users_status(bearer_token, users)
    sys.exit(0)
