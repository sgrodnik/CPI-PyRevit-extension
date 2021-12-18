# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import BuiltInCategory as bic
from collections import namedtuple
from pyrevit import script, forms
from System.Collections.Generic import *
import Autodesk.Revit.DB as db
import re

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

REPORT_ON = not 0
JUST_SEL = __shiftclick__


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


def str_param(param):
    if str(param.StorageType) == 'Double':
        val = param.AsDouble()
        if val:
            val = '{:.2f}'.format(val)
        else:
            val = '<p style="color:Gainsboro">' + str(val) + '</p>'
        return val
    if str(param.StorageType) == 'String':
        val = param.AsString()
        if val:
            val = str(val)
        else:
            val = '<p style="color:Gainsboro">' + 'None' + '</p>'
        return val
    if str(param.StorageType) == 'ElementId':
        return 'Id{}'.format(param.AsElementId())
    if str(param.StorageType) == 'Integer':
        val = param.AsInteger()
        if param.AsValueString() == 'Да' or param.AsValueString() == 'Нет':
            if val:
                val = '<p style="color:SeaGreen">' + '✓' + '</p>'
            else:
                val = '<p style="color:LightPink">' + '✗' + '</p>'
            return val
        if val:
            val = str(val)
        else:
            val = '<p style="color:Gainsboro">' + str(val) + '</p>'
        return val


# ----------------------------------------------------------------------------
# ----------------------------------- Main -----------------------------------
# ----------------------------------------------------------------------------


ALLOWED = [
    'Фамилия',
    'Подпись',
    'Дата вручную',
    'Имя листа',
    'Время печати',
    'Выносные линии',
    'Количество измов для',
]

title_blocks = [Lookuper(el) for el in get_collector('OST_TitleBlocks')]
sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]
sel = [el.Id for el in sel if el.LookupParameter('Категория').AsValueString() == 'Листы']

if not sel:
    script.exit()

PARAMS = []
PARAM_NAMES = []
report = []  # Формирование отчёта
tbs = []
for tb in title_blocks:
    owner = Lookuper(doc.GetElement(tb.OwnerViewId))
    owner_id = tb.OwnerViewId
    parameters = natural_sorted(tb.Parameters, lambda p: p.Definition.Name)
    if owner_id in sel:
        tbs.append(tb)
        report.append([output.linkify(tb.Id, owner.SheetNumber)] + [str_param(p) for p in parameters if any([s in p.Definition.Name for s in ALLOWED])])
        PARAMS = PARAMS or parameters
        PARAM_NAMES = PARAM_NAMES or [p.Definition.Name for p in parameters if any([s in p.Definition.Name for s in ALLOWED])]

if JUST_SEL:
    el_ids = List[db.ElementId]([el.Id for el in tbs])
    uidoc.Selection.SetElementIds(el_ids)
    script.exit()

from collections import OrderedDict
options_dict = OrderedDict()
for p in PARAMS:
    options_dict[p.Definition.Name] = p
selected_param_name = forms.CommandSwitchWindow.show(
    options_dict,
    message='Выберите параметр (Esc для отчёта):',
    width=400
)
if selected_param_name:
    selected_param = options_dict[selected_param_name]
    value = forms.ask_for_string(
        default=' '.join(natural_sorted(list(set(['<' + str(tb.Look(selected_param.Definition.Name)) + '>' for tb in tbs])))),
        prompt='Введите новое значение для параметра ' + selected_param.Definition.Name,
        title='Параметры основной надписи')
    if value is None:
        script.exit()
    if str(selected_param.StorageType) == 'Double':
        value = float(value)
    elif str(selected_param.StorageType) == 'Integer':
        value = int(value)
    elif str(selected_param.StorageType) == 'ElementId':
        value = db.ElementId(int(value))

    t = db.Transaction(doc, 'Параметры основной надписи')
    t.Start()
    for tb in tbs:
        print(type(value))
        print(value)
        tb.LookupParameter(selected_param.Definition.Name).Set(value)
    t.Commit()

    if REPORT_ON:
        report = []
        for tb in tbs:
            owner = Lookuper(doc.GetElement(tb.OwnerViewId))
            parameters = natural_sorted(tb.Parameters, lambda p: p.Definition.Name)
            report.append([output.linkify(tb.Id, owner.SheetNumber)] + [str_param(p) for p in parameters if any([s in p.Definition.Name for s in ALLOWED])])

if REPORT_ON:
    output.print_table(  # Вывод отчёта
        table_data=report,
        columns=[
            'Номер листа',
        ] + PARAM_NAMES
    )
