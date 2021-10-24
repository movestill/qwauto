from argparse import ArgumentParser, Namespace
from configparser import ConfigParser
from msvcrt import getch
from pywinauto.controls.hwndwrapper import HwndWrapper
from pywinauto.application import Application, ProcessNotFoundError, WindowSpecification
import sys
from win32gui import GetForegroundWindow

QUICKEN_EXE = r"c:\Program Files (x86)\Quicken\qw.exe"
MAX_TRIES = 5


class QwAutoConfig:
    def __init__(self, quicken_exe, boughtx_acct):
        self.quicken_exe = quicken_exe
        self.boughtx_acct = boughtx_acct


def connect(quicken_path: str = QUICKEN_EXE) -> Application:
    """
    Connect to a running Quicken.  Exits if unable to connect.

    Args:
        quicken_path: Path to the Quicken executable.

    Returns:
        Quicken application connection.
    """
    print("Connecting to Quicken", end="")
    tries = 0
    while tries < MAX_TRIES:
        try:
            print(".", end="", flush=True)
            quicken = Application().connect(path=quicken_path)
        except ProcessNotFoundError:
            tries += 1
            continue
        break

    if tries == MAX_TRIES:
        print("Could not connect to Quicken.  Is it running?")
        sys.exit(1)

    print("Connected to Quicken")
    return quicken


def switch_to_boughtx(window: WindowSpecification, account: str):
    """
    Convert a downloaded investment transacton from Bought to BoughtX.

    Args:
        window: Main Quicken window.
        account: Source of funds for purchase.
    """
    PAUSE_TIME = .1
    window.type_keys("{TAB}")
    window.type_keys("boughtx")

    # Tab to transfer field.
    import time
    for i in range(7):
        time.sleep(PAUSE_TIME)
        window.type_keys("{TAB}")

    time.sleep(PAUSE_TIME)
    window.type_keys(account)
    time.sleep(PAUSE_TIME)
    # Accept account.
    window.type_keys("{ENTER}")
    time.sleep(PAUSE_TIME)
    # Save transaction.
    window.type_keys("{ENTER}")


def print_interact_help():
    """
    Print help for interactive mode.
    """
    print("Commands:")
    print("\tj - up")
    print("\tk - down")
    print("\t<Space> - switch Bought to BoughtX")
    print("\t<Enter> - send Enter to Quicken")
    print("\t<Escape> - quit")


def interact(qwauto_window, window: WindowSpecification, account: str):
    """
    Pass select inputs to Quicken.

    Args:
        window: Main Quicken window.
        account: Source of funds for purchase.
    """
    print_interact_help()

    while True:
        key = getch()

        if key == bytes("\x1b", "utf-8") or key == bytes("\x03", "utf-8"):
            return
        elif key == bytes("j", "utf-8"):
            window.type_keys("{DOWN}")
        elif key == bytes("k", "utf-8"):
            window.type_keys("{UP}")
        elif key == bytes(" ", "utf-8"):
            switch_to_boughtx(window, account)
        elif key == bytes("\r", "utf-8"):
            window.type_keys("{ENTER}")

        # Switch input focus back to qwauto.
        qwauto_window.set_focus()
        qwauto_window.set_keyboard_focus()


def parse() -> Namespace:
    """
    Parse command line arguments.
    """
    parser = ArgumentParser()
    parser.add_argument(
        "--config",
        "-c",
        default="qwauto.cfg",
        help="Config file.  Defaults to qwauto.cfg.",
    )
    return parser.parse_args()


def load_config(config_file: str) -> QwAutoConfig:
    """
    Load config file.

    Args:
        config_file: Name of config file.

    Returns:
        Configuration.
    """
    config = ConfigParser()
    config.read(config_file)
    default = config["DEFAULT"]
    return QwAutoConfig(default.get("quicken", QUICKEN_EXE), default["boughtx_account"])


def main() -> int:
    qwauto_window = HwndWrapper(GetForegroundWindow())
    cmd_args = parse()
    config = load_config(cmd_args.config)
    quicken = connect(config.quicken_exe)
    window = quicken.top_window()
    interact(qwauto_window, window, config.boughtx_acct)

    # Restore focus to Quicken window.
    window.set_focus()

    return 0


if __name__ == "__main__":
    sys.exit(main())
