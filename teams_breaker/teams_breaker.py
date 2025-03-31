import argparse
import csv
import sys
import time
import threading
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
            time.sleep(STAGGER_DELAY)
            presence_data = user.check_teams_presence()
            if presence_data is None:
                statuses.append([email, "No presence", "No presence"])
            else:
                # Assume presence_data is a list and we use the first element.
                user_info = presence_data[0]
                presence = user_info.get("presence", {})
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

def update_user_list_file(list_path: Path, users):
    """Writes the current user list back to the file."""
    with open(list_path, "w") as f:
        for user in users:
            f.write(user + "\n")

def prompt_input(win, prompt):
    """Prompt user input in the provided window (curses)."""
    win.clear()
    win.addstr(0, 0, prompt)
    win.refresh()
    curses.echo()
    input_str = win.getstr(1, 0).decode("utf-8")
    curses.noecho()
    return input_str.strip()

def update_status_thread(bearer_token, users, statuses, users_lock, statuses_lock, termination_event):
    """Background thread that updates statuses periodically."""
    while not termination_event.is_set():
        # Copy the current user list safely.
        with users_lock:
            current_users = list(users)
        new_statuses = get_users_status(bearer_token, current_users)
        write_status_csv(new_statuses)  # Update CSV file.
        with statuses_lock:
            statuses.clear()
            statuses.extend(new_statuses)
        # Wait for the next refresh cycle.
        for _ in range(REFRESH_INTERVAL):
            if termination_event.is_set():
                break
            time.sleep(1)

def curses_main(stdscr, bearer_token, users, list_path, statuses, users_lock, statuses_lock, termination_event):
    curses.curs_set(0)  # Hide cursor.
    stdscr.nodelay(True)  # Non-blocking input.

    # Initialize colors if supported.
    if curses.has_colors():
        curses.start_color()
        curses.init_pair(1, curses.COLOR_RED, curses.COLOR_BLACK)    # Busy -> red
        curses.init_pair(2, curses.COLOR_GREEN, curses.COLOR_BLACK)  # Available -> green
        curses.init_pair(3, curses.COLOR_BLUE, curses.COLOR_BLACK)   # Offline -> blue

    height, width = stdscr.getmaxyx()

    # Create subwindows: one for status and one for command/input.
    status_win = curses.newwin(height - 4, width, 0, 0)
    input_win = curses.newwin(4, width, height - 4, 0)

    while True:
        # Read the current statuses.
        with statuses_lock:
            current_statuses = list(statuses)

        # Clear windows.
        status_win.clear()
        input_win.clear()

        # Draw table header.
        header = " Email ".ljust(30) + "| Availability ".ljust(20) + "| Device Type "
        status_win.addstr(0, 0, header)
        status_win.addstr(1, 0, "-" * (len(header) + 10))

        # Display each user status.
        for idx, row in enumerate(current_statuses):
            email, availability, device = row
            row_y = idx + 2
            if row_y >= height - 4:
                break  # Prevent writing outside the window.

            # Print email.
            status_win.addstr(row_y, 0, f" {email:<28} | ")

            # Determine color for availability.
            color_attr = curses.A_NORMAL
            if curses.has_colors():
                if availability.lower() == "busy":
                    color_attr = curses.color_pair(1)
                elif availability.lower() == "available":
                    color_attr = curses.color_pair(2)
                elif availability.lower() == "offline":
                    color_attr = curses.color_pair(3)

            # Print availability column in color.
            status_win.addstr(f"{availability:<18}", color_attr)
            # Continue with device type.
            status_win.addstr(" | " + f"{device}")
        status_win.refresh()

        # Display command options.
        input_win.addstr(0, 0, "Commands: (a)dd user, (r)emove user, (q)uit")
        input_win.addstr(1, 0, "Press the corresponding key for action.")
        input_win.refresh()

        # Check for user input.
        try:
            key = stdscr.getch()
        except Exception:
            key = -1

        if key != -1:
            if key in [ord("q"), ord("Q")]:
                termination_event.set()
                break  # Exit the loop.
            elif key in [ord("a"), ord("A")]:
                # Switch temporarily to blocking mode for input.
                stdscr.nodelay(False)
                new_email = prompt_input(input_win, "Enter email to add:")
                if new_email:
                    with users_lock:
                        if new_email not in users:
                            users.append(new_email)
                            update_user_list_file(list_path, users)
                stdscr.nodelay(True)
            elif key in [ord("r"), ord("R")]:
                stdscr.nodelay(False)
                rem_email = prompt_input(input_win, "Enter email to remove:")
                if rem_email:
                    with users_lock:
                        if rem_email in users:
                            users.remove(rem_email)
                            update_user_list_file(list_path, users)
                stdscr.nodelay(True)
        time.sleep(0.1)

if __name__ == "__main__":
    args = build_argparser().parse_args()

    logger.info("Running in TUI mode.")

    list_path = Path(args.list)
    if list_path.is_file():
        users = list(set(list_path.read_text().strip().splitlines()))
    else:
        logger.error("User list file does not exist or it is not a file!")
        sys.exit(1)

    # Authenticate and retrieve tokens.
    bearer_token, skype_token, sharepoint_token, sender_info = teams_api.authenticate(
        args.username, args.password, False
    )

    # Shared data structures for thread communication.
    statuses = []           # Shared statuses list.
    users_lock = threading.Lock()    # Protects the users list.
    statuses_lock = threading.Lock() # Protects the statuses list.
    termination_event = threading.Event()

    # Start the background thread to update statuses.
    status_thread = threading.Thread(
        target=update_status_thread,
        args=(bearer_token, users, statuses, users_lock, statuses_lock, termination_event),
        daemon=True,
    )
    status_thread.start()

    # Launch the curses TUI.
    curses.wrapper(curses_main, bearer_token, users, list_path, statuses, users_lock, statuses_lock, termination_event)

    # Wait for the background thread to exit.
    status_thread.join()

