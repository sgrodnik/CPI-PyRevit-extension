# -*- coding: utf-8 -*-
"""Description"""

from Autodesk.Revit.DB import BuiltInCategory as bic
from collections import namedtuple
from pyrevit import script, forms
import Autodesk.Revit.DB as db

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

FEET_TO_MM = 304.8
FEET_TO_M = 304.8 / 1000
MM_TO_FEET = 1 / FEET_TO_MM
M_TO_FEET = 1 / FEET_TO_M
F2_TO_M2 = FEET_TO_MM**2 / 10**6

# Значение, после превышения которого скрипт будет вычитать площадь отбойника
# из площади чистовой отделки
GUARD_THRESHOLD = 500 * MM_TO_FEET  # Пороговое значение учёта отбойника

# Следующую строку трогать не нужно
REPORT_ON = not __shiftclick__
# Поведение по умолчанию в части вывода отчёта можно переключить,
# раскоментировав последующую строку
# REPORT_ON = __shiftclick__  # Отчёт не выводится. Shift + Клик включает вывод отчёта
# ↑↑↑ ↑↑↑ ↑↑↑ ↑↑↑ Раскомментируй эту строку ↑↑↑ ↑↑↑ ↑↑↑ ↑↑↑


def to_mm(feet_val):
    return round(feet_val * FEET_TO_MM, 0)


def to_sq(sq):
    return round(sq * F2_TO_M2, 2)


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

    def __str__(self):
        return self.obj.__repr__()


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
        return 0
    description = param.replace('мм', '').replace('=', '').replace(',', '').replace('.', '')
    digits = [int(s) for s in description.split() if s.isdigit()]
    return digits[0] if digits else 0


def valid(instance):
    if not instance:
        return False
    if instance.Category.Name == '<Разделитель помещений>':
        return False
    if instance.LookupParameter('Семейство').AsValueString() == 'Витраж':
        return False
    return True


Segment = namedtuple('Segment', [
    'length',
    'decor_base',
    'apertures',
    'host_id',
    'seg_prep_decor_area',
])


class Room():  # Основной расчёт помещений
    """Wrapper for calculating the decorating of room"""
    objects = []

    def __init__(self, room):
        self.__class__.objects.append(self)
        self.origin = room
        self.Id = room.Id
        self.full_heigth = room.Look("Полная высота")
        self.perim = room.Look("Периметр")
        self.ceiling_heigth = room.Look("CPI_Потолок_Высота") \
            or self.full_heigth
        self.final_decor_heigth = min(self.ceiling_heigth + 100 * MM_TO_FEET,
                                      self.full_heigth)
        self.segments = []
        self.apertures_area = 0
        self.prep_decor_area = {}  # Supposed to be {decor_base: Area}
        self.final_decor_area = 0
        self.baseboard_on = room.Look("CPI_Плинтус_Наличие")
        self.baseboard_lenth = 0
        self.baseboard_height = parse_baseboard_height(room) * MM_TO_FEET
        self.guard_on = room.Look("CPI_Отбойник_Наличие")
        self.guard_width = room.Look("CPI_Отбойник_Ширина")
        self.guard_height = room.Look("CPI_Отбойник_Отметка верха")
        self.guard_reserve = room.Look("CPI_Отбойник_Запас") or 0
        self.guard_lenth = 0 + self.guard_reserve
        self.apron_on = room.Look("CPI_Фартук_Наличие")
        self.apron_width = room.Look("CPI_Фартук_Ширина")
        self.apron_height = room.Look("CPI_Фартук_Высота")
        self.apron_area = self.apron_width \
            * self.apron_height if self.apron_on else 0
        self.aperture_ids = []
        for segment in flatten(
                room.GetBoundarySegments(db.SpatialElementBoundaryOptions())):
            instance = doc.GetElement(segment.ElementId)
            if not valid(instance):
                continue
            symbol = Lookuper(doc.GetElement(instance.GetTypeId()))
            decor_base = symbol.Look('CPI_Основа черновой отделки')
            if not decor_base:
                errors('Параметр "CPI_Основа черновой отделки" не заполнен, \
                        расчёт черновой отделки некорректен',
                       instance.Id.IntegerValue)
                # continue
                decor_base = '???'
            length = segment.GetCurve().Length
            if decor_base not in self.prep_decor_area:
                self.prep_decor_area[decor_base] = 0
            host_id = segment.ElementId.IntegerValue
            apertures_ = apertures_by_host.get(host_id, [])
            apertures = []
            phase = doc.GetElement(room.Look('Стадия'))
            for ap in apertures_:
                if ap.Id in self.aperture_ids:
                    continue
                if (ap.FromRoom[phase] and ap.FromRoom[phase].Id.IntegerValue == room.Id.IntegerValue) \
                        or (ap.ToRoom[phase] and ap.ToRoom[phase].Id.IntegerValue == room.Id.IntegerValue):
                    self.aperture_ids.append(ap.Id)
                    apertures.append(ap)
            self.baseboard_lenth += length if self.baseboard_on else 0
            self.guard_lenth += length if self.guard_on else 0
            for ap in apertures:
                ap_sill_height = ap.get_Parameter(
                    db.BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).AsDouble()
                if self.baseboard_on:
                    if self.baseboard_height > ap_sill_height:
                        self.baseboard_lenth -= get_width(ap)
                if self.guard_on:
                    if self.guard_height > ap_sill_height:
                        self.guard_lenth -= get_width(ap)
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
        if self.guard_width >= GUARD_THRESHOLD:
            self.final_decor_area -= self.guard_width * self.guard_lenth
        self.final_decor_area -= self.apron_area
        self.number = room.Look("Номер")

    def commit(self):  # Прописывание значений параметров
        self.origin.LookupParameter('CPI_Чистовая_Площадь отделки') \
            .Set(self.final_decor_area)
        areas = {}  # Supposed to be {decor_base: [Area, ElementIds]}
        for seg in self.segments:
            if seg.decor_base not in areas:
                areas[seg.decor_base] = [0, []]  # [Area, ElementIds]
            areas[seg.decor_base][0] += seg.seg_prep_decor_area
            areas[seg.decor_base][1].append(seg.host_id.IntegerValue)
        for base in areas:
            par = self.origin.LookupParameter(
                'CPI_Черновая-' + base + '_Площадь')
            if par:
                par.Set(areas[base][0])
            else:
                errors('Не найден параметр "CPI_Черновая-{0}_Площадь", \
                        значение площади для "{0}" не записано'.format(base),
                       areas[base][1])
        self.origin.LookupParameter('CPI_Плинтус_Длина') \
            .Set(self.baseboard_lenth)
        self.origin.LookupParameter('CPI_Отбойник_Длина') \
            .Set(self.guard_lenth)


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
apertures = [Lookuper(el) for el in doors + windows if el.Host]
apertures_by_host = pack_apertures_by_host(apertures)
sel = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]
rooms = [el for el in sel if el.Category.Name == 'Помещения']
all_rooms = get_collector('OST_Rooms')
rooms = rooms or all_rooms

t = db.Transaction(doc, 'Отделка: Простановка галочек помещениям')
t.Start()
for room in rooms:
    for param_name in ['CPI_Плинтус_Наличие',
                       'CPI_Фартук_Наличие',
                       'CPI_Отбойник_Наличие',
                       'CPI_Подсчёт отделки']:
        param = room.LookupParameter(param_name)
        if not param.HasValue:
            param.Set(1)
t.Commit()

rooms = [r for r in rooms if r.LookupParameter('CPI_Подсчёт отделки').AsInteger()]

title = 'Основной расчёт'
rooms_ = []
with forms.ProgressBar(title=title, cancellable=True) as pb:
    i = 0
    for room in rooms:
        if room.Area > 0:
            rooms_.append(Room(Lookuper(room)))
            pb.title = '{}: {} из {}: Помещение № {}'.format(title,
                                                             i + 1,
                                                             len(rooms),
                                                             room.Number)
        if pb.cancelled:
            break
        else:
            pb.update_progress(i, len(rooms))
        i += 1
rooms = rooms_

if pb.cancelled:
    script.exit()

t = db.Transaction(doc, 'Отделка')
t.Start()
for room in rooms:
    room.commit()
t.Commit()

title = 'Формирование отчёта'
report = []  # Формирование отчёта
with forms.ProgressBar(title=title, cancellable=True) as pb:
    i = 0
    for room in rooms:
        if not REPORT_ON:
            continue
        finish_area = 'Sч = {:n} м²'.format(to_sq(room.final_decor_area))
        prep_areas = '<br>'\
            .join([finish_area] + ['S{} = {:n} м²'.format(
                decor_base.lower(),
                to_sq(room.prep_decor_area[decor_base])
            )
                for decor_base in room.prep_decor_area])
        room_info = '{}<br>{} {}<br>{}' \
            .format(i + 1,
                    output.linkify(room.origin.Id,
                                   room.origin.Look('Номер')),
                    room.origin.Look('Имя'),
                    prep_areas,
                    )
        walls_info = []
        apertures_info = []
        segs_area = 0
        aps_area = 0
        perim = 0
        for i_seg, seg in enumerate(room.segments):
            seg_area = to_sq(seg.length * room.final_decor_heigth)
            segs_area += seg_area
            perim += seg.length
            # if len(rooms) < 4:
            room_mark = output.linkify(seg.host_id, '{} {}'.format(i_seg + 1, seg.decor_base))
            # else:
                # room_mark = '{} {}'.format(i_seg + 1, seg.decor_base)
            walls_info.append(
                '{}: L = {:n} ({:n}), h = {:n} ({:n}), S = {:n} ({:n})'.format(
                    room_mark,
                    to_mm(seg.length),
                    to_mm(perim),
                    to_mm(room.final_decor_heigth),
                    to_mm(room.full_heigth),
                    seg_area,
                    segs_area)
            )
            for i_ap, ap in enumerate(seg.apertures):
                ap_area = to_sq(get_area(ap))
                aps_area += ap_area
                apertures_info.append(
                    '{} S = {:n} ({:n})'.format(
                        output.linkify(ap.Id, '{}.{}'.format(i_seg + 1,
                                                             i_ap + 1)),
                        ap_area,
                        aps_area)
                )
        baseboard_info = '{:n}<br>h={:n}'.format(
            to_mm(room.baseboard_lenth),
            to_mm(room.baseboard_height)
        )
        diff = room.guard_lenth - room.guard_reserve
        guardrail_info = \
            '{:n}{}<br>Ш = {:n} мм<br>Отм. в. {:n} мм<br>S = {:n} м²'.format(
                to_mm(room.guard_lenth),
                ' =<br>{:n}{}{:n}'.format(
                    to_mm(diff),
                    ' + ' if room.guard_reserve > 0 else ' ',
                    to_mm(room.guard_reserve)) if room.guard_reserve else '',
                to_mm(room.guard_width),
                to_mm(room.guard_height),
                to_sq(room.guard_width * room.guard_lenth),
            )
        apron_info = '{:n} м² =<br>{:n}×{:n}'.format(
            to_sq(room.apron_area),
            to_mm(room.apron_width),
            to_mm(room.apron_height),
        )
        report.append([room_info,
                       '<br>'.join(walls_info),
                       '<br>'.join(apertures_info),
                       baseboard_info,
                       guardrail_info,
                       apron_info,
                       ])

        pb.title = '{}: {} из {}: Помещение № {}'.format(title,
                                                         i + 1,
                                                         len(rooms),
                                                         room.number)
        if pb.cancelled:
            break
        else:
            pb.update_progress(i, len(rooms))
        i += 1

if REPORT_ON:
    if report:
        LIMIT = 10
        reports = []
        for i in range(len(report)):
            if len(report) >= LIMIT:
                reports.append(report[:LIMIT])
                report = report[LIMIT:]
            else:
                reports.append(report) if report else None
                break
        for report in reports:
            output.print_table(  # Вывод отчёта
                table_data=report,
                columns=[
                    'Помещение, м²',
                    'Стены: Длина, мм (Σмм); Высота (черновая), мм; Площадь, м² (Σм²)',
                    'Проёмы: площадь, м² (Σм²)',
                    'Плинтус',
                    '<p title="Пороговая ширина отбойника для учёта его площади в'
                    + 'чистовой отделке составляет {0:n} мм">Отбойник {0:n}</p>'
                    .format(GUARD_THRESHOLD * FEET_TO_MM),
                    'Фартук',
                ]
            )
    rooms_off = [r for r in all_rooms if not r.LookupParameter('CPI_Подсчёт отделки').AsInteger()]
    print('\nПомещений в проекте всего: {}'.format(len(all_rooms)))
    print('Помещений с нулевой площадью: {}'.format(len([r for r in all_rooms if r.Area == 0])))
    print('Помещений с выключенным "CPI_Подсчёт отделки": {}'.format(len(rooms_off)))
    print('Обработано {}'.format(len(rooms)))

LIMIT = 50
for message in errs:  # Вывод ошибок
    print('\nПредупреждение: ' + message)
    element_ids_as_integer_value = sorted(list(errs[message]))
    element_ids = [db.ElementId(val) for val in element_ids_as_integer_value]
    el_ids = element_ids[0:LIMIT]
    too_big = len(element_ids) > LIMIT
    button_name = 'Выбрать{} {}{}'.format(
        ' первые ' if too_big else '',
        len(el_ids),
        ' из {} шт.'.format(len(element_ids)) if too_big else ' шт.'
    )
    sel_all_button = output.linkify(el_ids, button_name)
    print(sel_all_button
          + ' '.join([output.linkify(i) for i in el_ids])
          + (' ...' if too_big else '')
          )
