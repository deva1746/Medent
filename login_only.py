"""
Log in (or confirm already logged in) and stop at the main menu.

Run this on its own right before inspect_reports_menu.py, so there's no
manual typing/clicking delay between reaching the main menu and clicking
"Reports" -- the Reports flyout appears to auto-dismiss quickly, so minimizing
the gap matters.
"""

from app import (
    audit,
    connect_to_app,
    get_credentials,
    is_logged_in,
    login,
    wait_for_main_menu,
)


def main():
    user_id, password = get_credentials()

    app = connect_to_app()
    if is_logged_in(app):
        audit("already_logged_in")
    else:
        login(app, user_id, password)
    wait_for_main_menu(app)
    print("At main menu. Run inspect_reports_menu.py now.")


if __name__ == "__main__":
    main()
