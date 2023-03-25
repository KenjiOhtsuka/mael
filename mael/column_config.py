from enum import Enum
import openpyxl
import yaml


class ValueType(Enum):
    INCREMENT = 1
    STRING = 2
    LIST = 3


class Alignment(Enum):
    LEFT = 1
    CENTER = 2
    RIGHT = 3

    def excel_alignment(self):
        if self == Alignment.CENTER:
            return openpyxl.styles.alignment.Alignment(
                wrap_text=True, vertical='top', horizontal='center')
        if self == Alignment.RIGHT:
            return openpyxl.styles.alignment.Alignment(
                wrap_text=True, vertical='top', horizontal='right')
        return openpyxl.styles.alignment.Alignment(
            wrap_text=True, vertical='top', horizontal='left')


class ColumnCondition:
    def __init__(
            self,
            value_type: ValueType | str = ValueType.STRING,
            width: int = None,
            alignment: Alignment | str = Alignment.LEFT,
            duplicate_previous_for_blank: bool = None,
    ):
        self.type = value_type
        self.width = width
        if self.type == ValueType.INCREMENT:
            self.alignment = Alignment.RIGHT
        else:
            self.alignment = alignment
        self.duplicate_previous_for_blank = duplicate_previous_for_blank


class ColumnConfig:
    def __init__(self):
        self.prepend_columns = {}
        self.conditions = {}
        self.append_columns = {}
        self.overwrite_for_repeat = False
        self.duplicate_previous_for_blank = False

    def all_conditions(self) -> dict:
        return {**self.prepend_columns, **self.conditions, **self.append_columns}

    def list_columns(self) -> list[str]:
        return [k for k, v in self.conditions.items() if v.type == ValueType.LIST]

    def increment_columns(self) -> list[str]:
        return [
            k for k, v in {**self.prepend_columns, **self.append_columns}.items() if v.type == ValueType.INCREMENT
        ]

    def type_of(self, column: str) -> ValueType:
        return self.conditions[column].type if column in self.conditions else ValueType.STRING

    def parse(self, path: str) -> None:
        with open(path, 'r') as f:
            config = yaml.load(f, Loader=yaml.SafeLoader)
        if config is None:
            return
        # check dict value
        # refactor
        self.duplicate_previous_for_blank = \
            True == (config.get('global', {}).get('duplicate_previous_for_blank', False))
        self.overwrite_for_repeat = \
            True == (config.get('global', {}).get('overwrite_for_repeat', False))
        for name, column in config.get('prepend', {}).items():
            self.prepend_columns[name] = self.parse_condition(column)

        for name, column in config.get('column_conditions', {}).items():
            self.conditions[name] = self.parse_condition(column)

        for name, column in config.get('append', {}).items():
            self.append_columns[name] = self.parse_condition(column)

    def parse_condition(self, condition: dict):
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
            ValueType[condition['type'].upper()] if condition and 'type' in condition else ValueType.STRING,
            condition['width'] if condition and 'width' in condition else None,
            Alignment[condition['alignment'].upper()] if condition and 'alignment' in condition else Alignment.LEFT,
            condition.get('duplicate_previous_for_blank', self.duplicate_previous_for_blank) \
                if condition else self.duplicate_previous_for_blank,
        )


class Document:
    def __init__(self, file_path: str, variables = {}):
        self.title = None
        self.summary = None
        self.summary_lines = []
        self.list = []
