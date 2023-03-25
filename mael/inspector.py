import sys

LOAD_COMMANDS = ["load", "l"]
EXIT_COMMANDS = ["quit", "exit", "q", "bye"]
HELP_COMMANDS = ["help", "h", "?"]

def process_command(command):
    command = command.strip()
    if command == "":
        return
    if command in HELP_COMMANDS:
        print("Commands:")
        print("  hello: Print 'Hello, World!'")
        print("  load: Load a file")
        print("  exit: Exit the program")
        return
    if command in LOAD_COMMANDS:
        print("Loading...")
        return
    if command in EXIT_COMMANDS:
        print("Exiting...")
        sys.exit()

    print("Unknown command: " + command)

def repl(directory, env = None):
    while True:
        command = input("> ")

        process_command(command)
