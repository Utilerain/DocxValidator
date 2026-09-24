from colorama import Fore

FORE_VALUES = {
    getattr(Fore, attr)
    for attr in dir(Fore)
    if not attr.startswith('_')
}

def print_color(text: str, fore_color):
    if not fore_color in FORE_VALUES:
        raise TypeError("fore color is not AnsiFore enum")
    print(str(fore_color) + text + Fore.RESET)
