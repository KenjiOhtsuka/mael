import openpyxl
import yaml
from enum import Enum

class ValueType(Enum):
    INCREMENT = 1
    STRING = 2
    LIST = 3

class Alignment(Enum):
    LEFT = 1
    CENTER = 2
    RIGHT = 3

    def excel_alignment(self):
        if self == Alignment.LEFT:
            return openpyxl.styles.alignment.Alignment(
                wrap_text=True, vertical='top', horizontal='left')
        if self == Alignment.CENTER:
            return openpyxl.styles.alignment.Alignment(
                wrap_text=True, vertical='top', horizontal='center')
        if self == Alignment.RIGHT:
            return openpyxl.styles.alignment.Alignment(
                wrap_text=True, vertical='top', horizontal='right')

class ColumnCondition:
    def __init__(
            self,
            value_type: ValueType | str = ValueType.STRING,
            width: int = None,
            alignment: Alignment | str = Alignment.LEFT
    ):
        self.type = value_type
        self.width = width
        if self.type == ValueType.INCREMENT:
            self.alignment = Alignment.RIGHT
        else:
            self.alignment = alignment


class ColumnConfig:
    def __init__(self):
        self.prepend_columns = {}
        self.conditions = {}
        self.append_columns = {}

    def all_conditions(self):
        return {**self.prepend_columns, **self.conditions, **self.append_columns}

    def list_columns(self):
        return [k for k, v in self.conditions.items() if v.type == ValueType.LIST]

    def increment_columns(self):
        return [k for k, v in self.conditions.items() if v.type == ValueType.INCREMENT]

    def parse(self, path: str):
        with open(path, 'r') as f:
            config = yaml.load(f, Loader=yaml.SafeLoader)

        if 'prepend' in config:
            for name, column in config['prepend'].items():
                self.prepend_columns[name] = self.parse_condition(column)

        if 'column_conditions' in config:
            for name, column in config['column_conditions'].items():
                self.conditions[name] = self.parse_condition(column)

        if 'append' in config:
            for name, column in config['append'].items():
                self.prepend_columns[name] = self.parse_condition(column)

    @staticmethod
    def parse_condition(condition: dict):
        """
        Parse a column condition from a dict

        :param condition:
        :return:

        Example:
        {
            'value_type': 'increment',
            'width': 10,
            'alignment': 'right'
        }

        >>> c = ColumnConfig.parse_condition({
        ...     'value_type': 'increment',
        ...     'width': 10,
        ...     'alignment': 'right'
        ... })
        >>> c.type
        <ValueType.INCREMENT: 1>
        >>> c.width
        10
        >>> ColumnConfig.parse_condition({
        ...     'value_type': 'increment',
        ...     'width': 10,
        ...     'alignment': 'right'
        ... }).alignment
        <Alignment.RIGHT: 3>
        """
        return ColumnCondition(
            ValueType[condition['value_type'].upper()] if 'value_type' in condition else ValueType.STRING,
            condition['width'] if 'width' in condition else None,
            Alignment[condition['alignment'].upper()] if 'alignment' in condition else Alignment.LEFT
        )
