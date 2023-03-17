import argparse
import os
from initializer import Initializer
from excel_builder import build_excel

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
        # read the directory
        build_excel(target_dir)

# TODO:
#   * font configuration
