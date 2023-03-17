from .column_config import ColumnConfig, ValueType, Alignment
from openpyxl.utils.cell import get_column_letter
import functools
import glob
import openpyxl as px
import openpyxl.styles.alignment
import os
import re

VARIABLE_CONFIG_PATH = 'variables.ini'
COLUMN_CONFIG_PATHS = [
    'columns.yml',
    'columns.yaml',
]
IGNORE_FILE_PATH = 'ignore.txt'

THIN_BORDER = px.styles.Border(left=px.styles.Side(border_style='thin'),
                               right=px.styles.Side(border_style='thin'),
                               top=px.styles.Side(border_style='thin'),
                               bottom=px.styles.Side(border_style='thin'))

def trim_blank_lines(lines: list[str]) -> list[str]:
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
        elif self.type == ValueType.LIST:
            self.content_items = trim_blank_lines(self.content_items)
            return self.content_items


def build_excel(directory_path):
    # load column config
    column_config = ColumnConfig()
    for path in COLUMN_CONFIG_PATHS:
        file_path = os.path.join(directory_path, 'config', path)
        if os.path.exists(file_path):
            column_config.parse(file_path)
            break
    list_columns = column_config.list_columns()

    # load variables from ini
    variables = {}
    file_path = os.path.join(directory_path, 'config', VARIABLE_CONFIG_PATH)
    if os.path.exists(file_path):
        with open(file_path, 'r') as f:
            result = re.findall(r'^(?P<key>[^#].*)=(?P<value>.*)', f.read(), flags=re.MULTILINE)
            for key, value in result:
                variables[key.strip()] = value.strip()

    ignore_file_path = os.path.join(directory_path, 'config', IGNORE_FILE_PATH)
    ignore_file_names = []
    if os.path.exists(ignore_file_path):
        with open(ignore_file_path, 'r') as f:
            ignore_file_names = list(filter(lambda x: x != '', map(lambda x: x.strip(), f.readlines())))

    # create new excel book
    wb = px.Workbook()

    # build Excel file
    for scenario_file in glob.glob(directory_path + '/*.md'):
        if os.path.basename(scenario_file) in ignore_file_names:
            continue

        # add Excel sheet
        with open(scenario_file) as f:
            # set name
            name = None
            while True:
                line = f.readline()

                if not line:
                    break

                result = re.match(r'^#[^#]\s*(\S.*)\s*$', line.rstrip())
                if result:
                    name = result.group(1)
                    break

            if not name:
                raise Exception('Title is not set, which must begin with "#" at the top of the file.')

            ws = wb.create_sheet(name)

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

            row_index = 1
            cell = ws.cell(row=row_index, column=1)
            cell.value = 'Summary'
            cell.font = px.styles.Font(bold=True)
            row_index += 2

            # read summary lines
            summary_lines = []
            while True:
                line = f.readline()
                if re.match(r'^##\s*(List|Steps|Rows)\s*$', line):
                    break
                else:
                    summary_lines.append(line.rstrip())
            summary_lines = trim_blank_lines(summary_lines)

            # write summary lines
            for summary_line in summary_lines:
                ws.cell(row=row_index, column=1).value = summary_line
                row_index += 1

            row_index += 1

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

                result = re.match(r'^###\s*(.*)\s*$', line)
                if result:
                    if item:
                        step_dict[item.title] = item.get_content()
                    title = result.group(1)
                    item = StepItem(
                        title,
                        column_config.conditions[
                            title].type if title in column_config.conditions else ValueType.STRING
                    )
                    continue

                if item:
                    item.add_content_line(line.rstrip())

        # write steps
        columns = functools.reduce(lambda x, y: x + [z for z in y if z not in x], map(lambda x: x.keys(), steps), [])
        for column in list_columns:
            if column in columns:
                index = columns.index(column)
                count = functools.reduce(max, map(lambda x: len(x[column]) if column in x else 0, steps), 0)
                # add numbered column
                for i in range(count - 1, -1, -1):
                    columns.insert(index + 1, f'{column} ({i + 1})')
                # split list column
                for step in steps:
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

        # write header
        for column_index, column in enumerate(columns):
            cell = ws.cell(row=row_index, column=column_index + 1)
            cell.value = column
            cell.font = px.styles.Font(bold=True)
            cell.border = THIN_BORDER

            letter = get_column_letter(column_index + 1)

            # arrange column width
            all_conditions = column_config.all_conditions()
            if column in all_conditions:
                condition = all_conditions[column]
                if condition.width:
                    ws.column_dimensions[letter].width = condition.width
                cell.alignment = condition.alignment.excel_alignment()
            else:
                cell.alignment = Alignment.LEFT.excel_alignment()
        row_index += 1

        # write steps
        increment_columns = column_config.increment_columns()
        for index, step in enumerate(steps):
            increment_value = index + 1
            for column in increment_columns:
                step[column] = increment_value

            for column_index, column in enumerate(columns):
                cell = ws.cell(row=row_index, column=column_index + 1)
                if column in step:
                    value = step[column]
                    if isinstance(value, str):
                        for k, v in variables.items():
                            value = re.sub(r'{{\s*' + k + r'\s*}}', v, value)
                    cell.value = value
                cell.border = THIN_BORDER

            row_index += 1

    wb.remove(wb.worksheets[0])

    # save Excel file
    basename = os.path.basename(os.path.abspath(directory_path))
    filename = f'{basename}.xlsx'
    wb.save(os.path.join(directory_path, filename))
    print('Saved', filename)
    return wb
