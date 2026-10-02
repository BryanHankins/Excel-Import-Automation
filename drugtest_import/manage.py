"""Server admin commands.

  python -m drugtest_import.manage gen-key
  python -m drugtest_import.manage create-user NAME --role admin|reviewer|viewer
  python -m drugtest_import.manage set-password NAME
  python -m drugtest_import.manage import-csv PATH
"""
import argparse
import getpass
import os
import sys

from .crypto import Cipher, ConfigError, generate_key
from .storage import ROLES, Database

ACTOR = f"cli:{getpass.getuser()}"


def open_db() -> Database:
    return Database(os.environ.get("DRUGTEST_DB", "drugtest.db"), Cipher.from_env(), os.environ.get("DRUGTEST_UPLOADS"))


def ask_password() -> str:
    password = getpass.getpass("Password (12+ characters): ")
    if password != getpass.getpass("Repeat password: "):
        raise ValueError("Passwords don't match")
    return password


def main(argv=None):
    parser = argparse.ArgumentParser(prog="python -m drugtest_import.manage")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("gen-key", help="print a new encryption key for DRUGTEST_KEY")
    create = commands.add_parser("create-user", help="add a user")
    create.add_argument("username")
    create.add_argument("--role", choices=ROLES, required=True)
    reset = commands.add_parser("set-password", help="reset a user's password")
    reset.add_argument("username")
    imp = commands.add_parser("import-csv", help="import records from the old CSV format")
    imp.add_argument("path")
    args = parser.parse_args(argv)

    if args.command == "gen-key":
        print(generate_key())
        print("Store this key somewhere safe (e.g. a password manager). Without it the data can't be read.",
              file=sys.stderr)
        return 0
    try:
        db = open_db()
        if args.command == "create-user":
            db.create_user(args.username, ask_password(), args.role, ACTOR)
            print(f"Created {args.role} {args.username}")
        elif args.command == "set-password":
            user = db.get_user_by_name(args.username)
            if not user:
                raise ValueError(f"No user {args.username!r}")
            db.update_user(user["id"], ACTOR, password=ask_password())
            print(f"Password updated for {args.username}")
        elif args.command == "import-csv":
            imported, skipped = db.import_csv(args.path, ACTOR)
            print(f"Imported {imported} records, skipped {skipped} incomplete rows")
    except (ConfigError, ValueError, OSError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
