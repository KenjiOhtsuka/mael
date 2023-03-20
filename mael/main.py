import argparse
import os
from .excel_builder import build_excel
from .initializer import Initializer


def main() -> None:
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawTextHelpFormatter,
        description="""
This is a tool to convert markdown file to excel.
""",
        epilog="""
== Example Use Case ==

# Initialize the directory
% mael init .

# Create some markdown files in the directory 

# Build the Excel file
% mael build .
"""
    )
    subparsers = parser.add_subparsers(dest='command')
    # parser for init command
    parser_init = subparsers.add_parser('init', help='Generate config files')
    parser_init.add_argument('directory', default=os.getcwd(), help='Directory to be initialized.')
    # parser for build command
    parser_build = subparsers.add_parser('build', help='Build Excel from markdown files')
    parser_build.add_argument('directory', default=os.getcwd(), help='Directory which holds markdown files.')
    parser_build.add_argument('-e', '--environment', help='Environment signature such as "dev" or "prod"')
    # parser for inspect command
    parser_build = subparsers.add_parser('inspect', help='Under development')
    parser_build.add_argument('directory', default=os.getcwd(), help='Directory which holds markdown files.')
    parser_build.add_argument('-e', '--environment', help='Environment signature such as "dev" or "prod"')

    args = parser.parse_args()

    if args.directory:
        if os.path.isabs(args.directory):
            target_dir = args.directory
        else:
            target_dir = os.path.join(os.getcwd(), args.directory)
    else:
        target_dir = os.getcwd()

    if args.command == 'init':
        # create init file
        i = Initializer(target_dir)
        i.initialize()
    elif args.command == 'build':
        # read the directory and save the Excel file
        build_excel(target_dir, args.environment)
    elif args.command == 'inspect':
        # read the directory and get into REPL
        #d = load_data()
        pass


def repl(directory, env = None):
    pass

# TODO:
#   * font configuration
