import argparse
import csv
import sys
import time
from pathlib import Path
import curses

import chevron
import teams_api
import yaml
from loguru import logger
from teams_user import TeamsUser

REFRESH_INTERVAL = 10  # seconds between full status refreshes
STAGGER_DELAY = 0.33   # delay between individual presence requests


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
        required=True,
        help="Full path to a file containing target emails. One per line.",
    )
    return parser


def create_thread_by_method(method):
    sw = {"closed_chat": teams_api.chat_create_closed_chat, "meeting": teams_api.chat_create_meeting}
    return sw.get(method)


def get_users_status(bearer_token, users):
    """Returns a list of [email, availability, device type] for each user."""
    statuses = []
    for email in users:
        user = TeamsUser(bearer_token, email)
        ustatus = user.get_status()
        if ustatus:
            # Stagger presence requests
            time.sleep(STAGGER_DELAY)
            user_info = user.check_teams_presence()[0]
            presence = user_info["presence"]
            availability = presence.get("availability", "Unknown")
            statuses.append([email, availability, presence.get("deviceType", "Unknown")])
        else:
            statuses.append([email, "Could not read", "Could not read"])
    return statuses


def write_status_csv(statuses, filename="user_status.csv"):
    with open(filename, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Email", "Availability", "Device Type"])
        writer.writerows(statuses)


def prompt_input(win, prompt):
    """Prompt user input in the provided window (curses)."""
    win.clear()
    win.addstr(0, 0, prompt)
    win.refresh()
    curses.echo()
    input_str = win.getstr(1, 0).decode("utf-8")
    curses.noecho()
    return input_str.strip()


def curses_main(stdscr, bearer_token, users):
    curses.curs_set(0)  # hide cursor
    stdscr.nodelay(True)  # non-blocking input
    height, width = stdscr.getmaxyx()

    # Create subwindows: one for status and one for commands/input.
    status_win = curses.newwin(height - 4, width, 0, 0)
    input_win = curses.newwin(4, width, height - 4, 0)

    last_update = 0
    statuses = []

    while True:
        now = time.time()
        if now - last_update >= REFRESH_INTERVAL:
            statuses = get_users_status(bearer_token, users)
            write_status_csv(statuses)  # optional: update CSV file
            last_update = now

        # Clear windows and draw borders if desired.
        status_win.clear()
        input_win.clear()

        # Draw the status table header.
        header = " Email ".ljust(30) + "| Availability ".ljust(20) + "| Device Type "
        status_win.addstr(0, 0, header)
        status_win.addstr(1, 0, "-" * (len(header) + 10))

        # Display each user status.
        for idx, row in enumerate(statuses):
            email, availability, device = row
            line = f" {email:<28} | {availability:<18} | {device}"
            # Prevent writing outside the window height
            if idx + 2 < height - 4:
                status_win.addstr(idx + 2, 0, line)

        status_win.refresh()

        # Display the command options.
        input_win.addstr(0, 0, "Commands: (a)dd user, (r)emove user, (q)uit")
        input_win.addstr(1, 0, "Press the corresponding key for action.")
        input_win.refresh()

        # Check for user input
        try:
            key = stdscr.getch()
        except Exception:
            key = -1

        if key != -1:
            if key in [ord("q"), ord("Q")]:
                break  # exit the loop
            elif key in [ord("a"), ord("A")]:
                # Temporarily switch to blocking mode for input
                stdscr.nodelay(False)
                new_email = prompt_input(input_win, "Enter email to add:")
                if new_email and new_email not in users:
                    users.append(new_email)
                stdscr.nodelay(True)
            elif key in [ord("r"), ord("R")]:
                stdscr.nodelay(False)
                rem_email = prompt_input(input_win, "Enter email to remove:")
                if rem_email in users:
                    users.remove(rem_email)
                stdscr.nodelay(True)
        time.sleep(0.1)


if __name__ == "__main__":
    args = build_argparser().parse_args()

    logger.info("Running in TUI mode.")

    list_path = Path(args.list)
    if list_path.is_file():
        users = list_path.read_text().strip().splitlines()
    else:
        logger.error("User list file does not exist or it is not a file!")
        sys.exit(1)

    # Authenticate and retrieve tokens.
    bearer_token, skype_token, sharepoint_token, sender_info = teams_api.authenticate(
        args.username, args.password, False
    )

    # Launch the curses TUI.
    curses.wrapper(curses_main, bearer_token, users)

