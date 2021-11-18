# -*- coding: utf-8 -*-
"""Description"""

from Autodesk.Revit.DB import BuiltInCategory as bic
from collections import namedtuple
from pyrevit import script
import Autodesk.Revit.DB as db
import re

FEET_TO_MM = 304.8
MM_TO_FEET = 1 / FEET_TO_MM
F2_TO_M2 = FEET_TO_MM ** 2 / 1000000

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

seb_options = db.SpatialElementBoundaryOptions()


def flatten(two_dim_list):  # https://stackoverflow.com/a/952952
    return [item for sublist in two_dim_list for item in sublist]


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


def get_area(el):
    return get_width(el) * get_height(el)


def get_width(el):
    symbol = Lookuper(el.Symbol)
    width = el.Look('Ширина') or el.Look('Примерная ширина') or \
        symbol.Look('Ширина') or symbol.Look('Примерная ширина')
    return width


def get_height(el):
    symbol = Lookuper(el.Symbol)
    height = el.Look('Высота') or el.Look('Примерная высота') or \
        symbol.Look('Высота') or symbol.Look('Примерная высота')
    return height


Segment = namedtuple('Segment', [
    'length',
    'decor_base',
    'apertures',
    'host_id',
    'seg_prep_decor_area',
])

errs = {}  # Supposed to be {message: Set(element_ids_as_integer_value)}


def errors(message, element_ids=None):
    if message not in errs:
        errs[message] = set()
    if element_ids:
        if isinstance(element_ids, list):
            [errs[message].add(el_id) for el_id in element_ids]
        else:
            errs[message].add(element_ids)


def parse_baseboard_height(room):
    param = room.Look('CPI_Плинтус_Описание')
    if not param:
        return None
    description = param.replace('мм', '')
    digits = [int(s) for s in description.split() if s.isdigit()]
    return digits[0] if digits else None


def valid(instance):
    if instance.Category.Name == '<Разделитель помещений>':
        return False
    if instance.LookupParameter('Семейство').AsValueString() == 'Витраж':
        return False
    return True


class Room():
    """Wrapper for calculating the decorating of room"""
    objects = []

    def __init__(self, room):
        self.__class__.objects.append(self)
        self.origin = room
        self.Id = room.Id
        self.full_heigth = room.Look("Полная высота")
        self.ceiling_heigth = room.Look("W_Потолок_Высота") or self.full_heigth
        self.final_decor_heigth = min(self.ceiling_heigth + 100 * MM_TO_FEET,
                                      self.full_heigth)
        self.segments = []
        self.apertures_area = 0
        self.prep_decor_area = {}  # Supposed to be {decor_base: Area}
        self.final_decor_area = 0
        self.baseboard_lenth = 0
        # self.baseboard_height = parse_baseboard_height(room)
        for segment in flatten(room.GetBoundarySegments(seb_options)):
            instance = doc.GetElement(segment.ElementId)
            if not valid(instance):
                continue
            symbol = Lookuper(doc.GetElement(instance.GetTypeId()))
            decor_base = symbol.Look('WH_Основа черновой отделки')
            if not decor_base:
                errors('Параметр "WH_Основа черновой отделки" не заполнен, \
                        элемент исключён из расчёта',
                       instance.Id.IntegerValue)
                continue
            length = segment.GetCurve().Length
            if decor_base not in self.prep_decor_area:
                self.prep_decor_area[decor_base] = 0
            host_id = segment.ElementId.IntegerValue
            apertures = apertures_by_host.get(host_id, [])
            self.baseboard_lenth += length
            for ap in apertures:
                get_height(ap)
                # check_for_baseboard(высотаРазмещенияПроёма, высотаПлинтуса)
                self.baseboard_lenth -= 123
            apertures_area = sum([get_area(ap) for ap in apertures])
            seg_prep_decor_area = length * self.full_heigth - apertures_area
            self.prep_decor_area[decor_base] += seg_prep_decor_area
            final_decor_area = \
                length * self.final_decor_heigth - apertures_area
            self.final_decor_area += final_decor_area
            self.segments.append(Segment(
                length=length,
                decor_base=decor_base,
                apertures=apertures,
                host_id=db.ElementId(host_id),
                seg_prep_decor_area=seg_prep_decor_area,
            ))

    def commit(self):
        self.origin.LookupParameter('WS_Стены_Площадь чистовой отделки') \
            .Set(self.final_decor_area)
        areas = {}  # Supposed to be {decor_base: [Area, ElementIds]}
        for seg in self.segments:
            if seg.decor_base not in areas:
                areas[seg.decor_base] = [0, []]  # [Area, ElementIds]
            areas[seg.decor_base][0] += seg.seg_prep_decor_area
            areas[seg.decor_base][1].append(seg.host_id.IntegerValue)
        for decor_base in areas:
            par = self.origin.LookupParameter('WS_Стены_Площадь ' + decor_base)
            if par:
                par.Set(areas[decor_base][0])
            else:
                errors('Не найден параметр "WS_Стены_Площадь {}", \
                        значение не записано'.format(decor_base),
                       areas[decor_base][1])


def pack_apertures_by_host(apertures):
    apertures_by_host = {}
    for ap in apertures:
        host_id = ap.Host.Id.IntegerValue
        if host_id not in apertures_by_host:
            apertures_by_host[host_id] = []
        apertures_by_host[host_id].append(ap)
    return apertures_by_host


def get_collector(cat_name, to_elements=True):
    return list(db.FilteredElementCollector(doc)
                  .OfCategory(getattr(bic, cat_name))
                  .WhereElementIsNotElementType()
                  .ToElements())


# ----------------------------------------------------------------------------
# ----------------------------------- Main -----------------------------------
# ----------------------------------------------------------------------------

doors = get_collector('OST_Doors')
windows = get_collector('OST_Windows')
apertures = [Lookuper(el) for el in doors + windows]
apertures_by_host = pack_apertures_by_host(apertures)
sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]
rooms = [el for el in sel if el.Category.Name == 'Помещения']
rooms = rooms or get_collector('OST_Rooms')
rooms = [Room(Lookuper(el)) for el in rooms if el.Area > 0]

t = db.Transaction(doc, 'Отделка')
t.Start()
for room in rooms:
    room.commit()
t.Commit()

report = []
for room in rooms:
    finish_area = 'Sч = {:n}'.format(room.final_decor_area * F2_TO_M2)
    prep_areas = '<br>'\
        .join([finish_area] + ['S{} = {:n}'
              .format(decor_base.lower(),
                      room.prep_decor_area[decor_base] * F2_TO_M2)
              for decor_base in room.prep_decor_area])
    room_info = '{} {}<br>{}' \
        .format(output.linkify(room.origin.Id,
                               room.origin.Look('Номер')),
                room.origin.Look('Имя'),
                prep_areas,
                )
    walls_info = []
    apertures_info = []
    segs_area = 0
    aps_area = 0
    for i_seg, seg in enumerate(room.segments):
        seg_area = seg.length * room.final_decor_heigth * F2_TO_M2
        segs_area += seg_area
        walls_info.append(
            '{} L = {:n}, h = {:n} ({:n}), S = {:n} ({:n})'.format(
                output.linkify(seg.host_id,
                               '{} {}'.format(i_seg + 1, seg.decor_base)),
                seg.length * FEET_TO_MM,
                room.final_decor_heigth * FEET_TO_MM,
                room.full_heigth * FEET_TO_MM,
                seg_area,
                segs_area)
        )
        for i_ap, ap in enumerate(seg.apertures):
            ap_area = get_area(ap) * F2_TO_M2
            aps_area += ap_area
            apertures_info.append(
                '{} S = {:n} ({:n})'.format(
                    output.linkify(ap.Id, '{}.{}'.format(i_seg + 1, i_ap + 1)),
                    ap_area,
                    aps_area)
            )
    report.append([room_info,
                   '<br>'.join(walls_info),
                   '<br>'.join(apertures_info), ]
                  )

output.print_table(
    table_data=report,
    columns=['Помещение', 'Стены', 'Проёмы', ],
)

for message in errs:
    print('Предупреждение: ' + message)
    element_ids_as_integer_value = sorted(list(errs[message]))
    element_ids = [db.ElementId(val) for val in element_ids_as_integer_value]
    button_name = 'Выбрать {} шт.'.format(len(element_ids))
    sel_all_button = output.linkify(element_ids, button_name)
    print(sel_all_button + ' '.join([output.linkify(i) for i in element_ids]))
