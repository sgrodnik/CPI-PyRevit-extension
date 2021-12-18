# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import BuiltInCategory as bic
from collections import namedtuple
import Autodesk.Revit.DB as db
import re

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


TABLE = [
    # Параметр целевой                         Источник 1                  Источник 2                  Источник 3               Источник 4              # noqa
    ('CPI_Отбойник_Номера помещений',        ['CPI_Отбойник_Описание']),                                                                                # noqa
    ('CPI_Потолок_Номера помещений',         ['CPI_Потолок_Тип']),                                                                                      # noqa
    ('CPI_Чистовая_Номера помещений',        ['CPI_Чистовая_Тип отделки']),                                                                             # noqa
    ('CPI_Пол2_Номера помещений',            ['CPI_Пол2_Тип конструкции', 'CPI_Плинтус_Описание']),                                                     # noqa
    ('CPI_Пол_Номера помещений',             ['CPI_Пол_Тип конструкции',  'CPI_Плинтус_Описание']),                                                     # noqa
    ('CPI_Фартук_Номера помещений',          ['CPI_Фартук_Описание',      'CPI_Фартук_Наличие']),                                                       # noqa
    ('CPI_Потолок и стены_Номера помещений', ['CPI_Потолок_Тип',          'CPI_Чистовая_Тип отделки', 'CPI_Отбойник_Описание']),                        # noqa
    ('CPI_Черновая_Номера помещений',        ['CPI_Черновая-Каркас_Тип',  'CPI_Черновая-ГБ_Тип',      'CPI_Черновая-КР_Тип',   'CPI_Черновая-ЖБ_Тип']), # noqa
    # target                                   source 1                    source 2                    source 3                 source 4                # noqa
]

EXCLUDED_NAMES = [
    'естничная клетка',
    'вакуационный выход',
]

SIMPLE_MODE = False
# SIMPLE_MODE = True


class Lookuper(object):  # https://stackoverflow.com/a/16185009
    """Wrapper for adding a bit of syntactic sugar to Elements.
    Allows to use the new method "el.Look" instead of a bulky
    "el.LookupParameter", due to it's unhandiness in term of necessity
    of thinkig about the type of a returning value."""

    def __init__(self, obj):
        self.obj = obj

    def __getattr__(self, name):
        if name == 'Look':
            return lambda s: None if not \
                self.obj.LookupParameter(s) \
                else self.obj.LookupParameter(s).AsDouble() if \
                str(self.obj.LookupParameter(s).StorageType) == 'Double' \
                else self.obj.LookupParameter(s).AsString() if \
                str(self.obj.LookupParameter(s).StorageType) == 'String' \
                else self.obj.LookupParameter(s).AsElementId() if \
                str(self.obj.LookupParameter(s).StorageType) == 'ElementId' \
                else self.obj.LookupParameter(s).AsInteger() if \
                str(self.obj.LookupParameter(s).StorageType) == 'Integer' \
                else None
        return getattr(self.obj, name)

    def __repr__(self):
        return self.obj.__repr__() + '*'

    def __str__(self):
        return self.obj.__repr__() + '**'


def get_collector(cat_name, to_elements=True):
    return list(db.FilteredElementCollector(doc)
                  .OfCategory(getattr(bic, cat_name))
                  .WhereElementIsNotElementType()
                  .ToElements())


def natural_sorted(list, key=lambda s: s):
    """
    Sort the list into natural alphanumeric order.
    """
    def get_alphanum_key_func(key):
        convert = lambda text: int(text) if text.isdigit() else text  # noqa
        return lambda s: [convert(c) for c in re.split('([0-9]+)', key(s))]
    sort_key = get_alphanum_key_func(key)
    return sorted(list, key=sort_key)


class Number:
    def __init__(self, number):
        self.origin = number
        self.prefix = '.'.join(number.split('.')[0:-1])
        self.base = number.split('.')[-1]
        self.int = int(self.base)

    def __str__(self):
        return 'origin {}||prefix {}||base {}'.format(self.origin, self.prefix,
                                                      self.base)


def get_grouped_numbers(rooms):
    nums_by_prefix = {}
    for room in natural_sorted(rooms, lambda r: r.Number):
        numo = Number(room.Number)
        if numo.prefix not in nums_by_prefix:
            nums_by_prefix[numo.prefix] = []
        nums_by_prefix[numo.prefix].append(numo)
    groups = [[]]
    for nums in nums_by_prefix.values():
        for i, numo in enumerate(nums):
            if SIMPLE_MODE:
                groups.append([])
            else:
                if len(nums) > 2:
                    if i > 0 and nums[i].int != nums[i - 1].int + 1:
                        groups.append([])
            groups[-1].append(numo)
            if i == len(nums) - 1:
                groups.append([])
    results = []
    groups = [group for group in groups if len(group) > 0]
    groups = natural_sorted(groups, lambda x: x[0].origin)
    temp = []
    filtered_groups = []
    for group in groups:
        if len(group) == 1 and str(group[0]) in temp:
            continue
        else:
            temp.append(str(group[0]))
            filtered_groups.append(group)
    for group in filtered_groups:
        numo = group[0]
        if numo.prefix:
            if len(group) == 2:
                s = '{0}.{1}, {0}.{2}'.format(numo.prefix, group[0].base, group[-1].base)
                if group[0].base == group[-1].base:
                    s = '{0}.{1}'.format(numo.prefix, group[0].base)
            else:
                if group[0].base != group[-1].base:
                    s = '{0}.{1}÷{0}.{2}'.format(numo.prefix, group[0].base, group[-1].base)
                else:
                    s = '{}.{}'.format(numo.prefix, group[0].base)
        else:
            if len(group) == 2:
                s = '{}, {}'.format(group[0].base, group[-1].base)
                if group[0].base == group[-1].base:
                    s = '{}'.format(group[0].base)
            else:
                if group[0].base != group[-1].base:
                    s = '{}÷{}'.format(group[0].base, group[-1].base)
                else:
                    s = '{}'.format(group[0].base)
        results.append(s)
    return ', '.join(results)

# ----------------------------------------------------------------------------
# ----------------------------------- Main -----------------------------------
# ----------------------------------------------------------------------------

rooms_num = [Lookuper(el) for el in get_collector('OST_Rooms') if el.Area > 0]
rooms_bad = [Lookuper(el) for el in get_collector('OST_Rooms') if el.Area == 0]

t = db.Transaction(doc, 'Группировка номеров помещений')
t.Start()
for target, sources in TABLE:
    rooms_by_kind = {}
    for room in rooms_num:
        kind = ' + '.join([src + (str(room.Look(src)) or '') for src in sources])
        room_name = room.Look('Имя')
        kind += str(any([name in room_name for name in EXCLUDED_NAMES]))
        if kind not in rooms_by_kind:
            rooms_by_kind[kind] = []
        rooms_by_kind[kind].append(room)
    for rooms in rooms_by_kind.values():
        s = get_grouped_numbers(rooms)
        for room in rooms:
            room.LookupParameter(target).Set(s)
    for room in rooms_bad:
        room.LookupParameter(target).Set('Не определено')
t.Commit()
