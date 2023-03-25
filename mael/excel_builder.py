import functools
import glob
import os
import re

from .column_config import ColumnConfig, ValueType, Alignment, Document
from .composer import OutputFormat

COLUMN_CONFIG_PATHS = [
    'columns.yml',
    'columns.yaml',
]

IGNORE_FILE_PATH = 'ignore.txt'


def variable_config_names(environment: str = None) -> list[str]:
    """Return a list of paths to variable config files.

    :param environment: environment signature such as "dev" or "test"
    :return: list of paths
    """
    if environment is None or environment == '':
        return ['variables.ini']

    return [
        'variables.ini',
        f'variables.{environment}.ini',
    ]


def read_variables(directory_path, environment: str = None) -> dict[str, str]:
    """Read variables from config files.

    :param directory_path: path to the directory which holds config files
    :param environment: environment signature such as "dev" or "test"
    :return: dictionary of variables
    """
    variables = {}
    for file_name in variable_config_names(environment):
        file_path = os.path.join(directory_path, 'config', file_name)
        if os.path.exists(file_path):
            with open(file_path, 'r') as f:
                result = re.findall(r'^(?P<key>[^#].*)=(?P<value>.*)', f.read(), flags=re.MULTILINE)
                for key, value in result:
                    variables[key.strip()] = value.strip()
    return variables


def read_column_config(directory_path) -> ColumnConfig:
    """Read column config from config files.

    :param directory_path: path to the directory which holds config files
    :return: ColumnConfig object
    """
    column_config = ColumnConfig()
    for path in COLUMN_CONFIG_PATHS:
        file_path = os.path.join(directory_path, 'config', path)
        if os.path.exists(file_path):
            column_config.parse(file_path)
            break
    return column_config


def read_ignore_file(directory_path) -> list[str]:
    """Read ignore file.

    :param directory_path: path to the directory which holds config files
    :return: list of file names
    """
    ignore_file_path = os.path.join(directory_path, 'config', IGNORE_FILE_PATH)
    if os.path.exists(ignore_file_path):
        with open(ignore_file_path, 'r') as f:
            return [line.strip() for line in f.readlines()]
    return []


def trim_blank_lines(lines: list[str]) -> list[str]:
    """Remove blank lines from front and back of lines.

    :param lines: list of lines
    :return: list of lines

    >>> trim_blank_lines([' ', '', 'a', '', 'b', 'c', '', ' ', ''])
    ['a', '', 'b', 'c']
    """
    # remove front blank lines
    blank_count = 0
    for i in range(len(lines)):
        if re.match(r"^\s*$", lines[i]):
            blank_count += 1
        else:
            break
    lines = lines[blank_count:]
    # remove back blank lines
    blank_count = 0
    for i in range(1, len(lines) + 1):
        if re.match(r"^\s*$", lines[-i]):
            blank_count += 1
        else:
            break
    if blank_count > 0:
        lines = lines[:-blank_count]
    return lines


class StepItem:
    def __init__(self, title: str = None, step_type: ValueType = ValueType.STRING):
        self.title = title
        self.type = step_type
        self.content_lines = []
        self.content_items = []

    def add_content_line(self, content) -> 'StepItem':
        if self.type == ValueType.STRING:
            if len(self.content_lines) == 0 and re.match(r'^\s*$', content):
                return self
            self.content_lines.append(content)
        elif self.type == ValueType.LIST:
            if len(self.content_items) == 0 and re.match(r'^\s*$', content):
                return self
            self.content_items.append(re.sub(r'^\s*\*\s*', '', content))
        return self

    def get_content(self) -> str | list:
        if self.type == ValueType.STRING:
            self.content_lines = trim_blank_lines(self.content_lines)
            return "\n".join(self.content_lines)
        if self.type == ValueType.LIST:
            self.content_items = trim_blank_lines(self.content_items)
            return self.content_items
        raise ValueError(f'Type {self.type} does not provide content.')


def convert(directory_path, environment: str = None, format: OutputFormat = OutputFormat.EXCEL):
    target_files = sorted(glob.glob(os.path.join(directory_path, '*.md')))
    if len(target_files) == 0:
        print(f'No markdown files found in {directory_path}')
        return

    # load column config
    column_config = read_column_config(directory_path)
    list_columns = column_config.list_columns()

    # load variables from ini
    variables = read_variables(directory_path, environment)

    ignore_file_path = os.path.join(directory_path, 'config', IGNORE_FILE_PATH)
    ignore_file_names = []
    if os.path.exists(ignore_file_path):
        with open(ignore_file_path, 'r') as f:
            ignore_file_names =\
                list(filter(lambda x: x != '', map(lambda x: x.strip(), f.readlines())))

    composer = OutputFormat.build_composer(format)

    # compose output
    for scenario_file in target_files:
        if os.path.basename(scenario_file) in ignore_file_names:
            continue

        document = Document(os.path.abspath(scenario_file), variables)
        # add Excel sheet
        with open(scenario_file) as f:
            # set name
            document.title = os.path.basename(scenario_file)
            while True:
                line = f.readline()

                if not line:
                    break

                result = re.match(r'^#[^#]\s*(\S.*)\s*$', line.rstrip())
                if result:
                    document.title = result.group(1)
                    break

            # set summary
            has_summary = False
            while True:
                line = f.readline()

                if not line:
                    break

                result = re.match(r'^##\s*Summary\s*$', line)
                if not result:
                    continue
                has_summary = True
                break

            if not has_summary:
                break

            # read summary lines
            summary_lines = []
            while True:
                line = f.readline()
                if re.match(r'^##\s*(List|Steps|Rows)\s*$', line):
                    break
                summary_lines.append(line.rstrip())
            document.summary_lines = trim_blank_lines(summary_lines)

            # read steps
            steps = []
            step_dict = {}
            item = None
            while True:
                line = f.readline()

                if not line:
                    if item:
                        step_dict[item.title] = item.get_content()
                    if len(step_dict) > 0:
                        steps.append(step_dict)
                    break

                if re.match(r'^\s*---\s*$', line):
                    if item:
                        step_dict[item.title] = item.get_content()
                        item = None
                    if len(step_dict) > 0:
                        steps.append(step_dict)
                        step_dict = {}
                    continue

                result = re.match(r'^#{3,}\s*(\S.*\S|\S)\s*$', line)
                if result:
                    if item:
                        step_dict[item.title] = item.get_content()
                    title = result.group(1)
                    if not column_config.overwrite_for_repeat:
                        if title in step_dict:
                            steps.append(step_dict)
                            step_dict = {}
                    item = StepItem(
                        title,
                        column_config.type_of(title)
                    )
                    continue

                if item:
                    item.add_content_line(line.rstrip())

        all_conditions = column_config.all_conditions()

        # update steps
        columns = functools.reduce(lambda x, y: x + [z for z in y if z not in x], map(lambda x: x.keys(), steps), [])
        # Copy the previous column value if the step doesn't have the column
        for index, step in enumerate(steps):
            step.update({
                k: v for k, v in steps[index - 1].items() \
                    if k not in step and (k not in all_conditions or all_conditions[k].duplicate_previous_for_blank)
            })
        for column in list_columns:
            if column in columns:
                index = columns.index(column)
                count = functools.reduce(max, map(lambda x: len(x[column]) if column in x else 0, steps), 0)
                # add numbered column
                for i in range(count - 1, -1, -1):
                    columns.insert(index + 1, f'{column} ({i + 1})')

                # split list column
                for index, step in enumerate(steps):
                    if column in step:
                        for i in range(len(step[column])):
                            step[f'{column} ({i + 1})'] = step[column][i]
                        del step[column]
                # remove original column
                columns.remove(column)

        for column_index, column in enumerate(column_config.prepend_columns.items()):
            columns.insert(column_index, column[0])

        for column in column_config.append_columns:
            columns.append(column)

        # add sheet
        composer.add_sheet(document, column_config, variables, all_conditions, columns, steps)

    basename = os.path.basename(os.path.abspath(directory_path))

    return composer.compose(directory_path, environment, basename)


def is_none_or_blank_string(value):
    return value is None or value.strip() == ''
