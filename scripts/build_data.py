#!/usr/bin/env python3
"""Собирает data.json из Excel-файлов в папке input/.

Файлы определяются по СОДЕРЖИМУ, а не по имени:
  * баланс БМТС: книга с листом «баланс» (+ листы поступление / выдача / итого / цеха);
  * ME2N: книга, у которой на первом листе есть колонка «Год плана закупа».
Если в input/ есть несколько файлов одного вида, берётся самый свежий (по дате последнего коммита).
После успешной сборки распознанные файлы удаляются из input/ (в репозитории это фиксирует workflow).
Если файла какого-то вида нет, соответствующая часть data.json остаётся прежней.
Колонки ищутся по заголовкам. Цены и суммы в data.json не попадают.
"""
import datetime, glob, json, os, re, subprocess, sys
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT = os.path.join(ROOT, 'input')
OUT = os.path.join(ROOT, 'data.json')


def s(v):
    return '' if v is None else str(v).strip()


def n(v):
    try:
        x = round(float(v), 3)
    except (TypeError, ValueError):
        return 0
    return int(x) if x == int(x) else x


def d(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    return s(v)[:10]


def hn(x):
    return re.sub(r'\s+', ' ', s(x).lower())


def rows(ws):
    return [r for r in ws.iter_rows(values_only=True) if any(c not in (None, '') for c in r)]


def commit_time(path):
    try:
        out = subprocess.run(['git', 'log', '-1', '--format=%ct', '--', path], cwd=ROOT,
                             capture_output=True, text=True).stdout.strip()
        if out:
            return int(out)
    except Exception:
        pass
    return int(os.path.getmtime(path))


def sheet_by_name(wb, *names):
    low = {hn(x): x for x in wb.sheetnames}
    for nm in names:
        if hn(nm) in low:
            return wb[low[hn(nm)]]
    raise KeyError('нет листа: ' + ' / '.join(names))


def header_map(R):
    return {hn(x): i for i, x in enumerate(R[0]) if x not in (None, '')}


# Запасные названия колонок: в файле их переименовывают (например «Базовая ЕИ» -> «ЕИ», в резервах «Объем потребности» -> «Заложено»)
AL = {
    'базовая еи': ('ЕИ',),
    'обозначение склада': ('Склады Цехов',),
    'объем потребности': ('Заложено',),
    'сокращ. количество': ('Получено',),
    'размер разницы': ('Неполучено',),
}


def _expand(names):
    out = []
    for nm in names:
        out.append(nm)
        out.extend(AL.get(hn(nm), ()))
    return out


def col(H, *names):
    for nm in _expand(names):
        if hn(nm) in H:
            return H[hn(nm)]
    raise KeyError('нет колонки: ' + ' / '.join(names))


def colo(H, *names):
    """необязательная колонка: индекс или None, если такой колонки нет"""
    for nm in _expand(names):
        if hn(nm) in H:
            return H[hn(nm)]
    return None


def gv(r, i, conv=s):
    """безопасно взять значение ячейки (если колонки нет, пусто / 0)"""
    return conv(r[i]) if i is not None and i < len(r) else conv(None)


def find_files():
    bal, me = [], []
    for f in sorted(glob.glob(os.path.join(INPUT, '*.xls*'))):
        if os.path.basename(f).startswith('~$'):
            continue
        try:
            wb = load_workbook(f, read_only=True, data_only=True)
        except Exception as e:
            print('пропуск (не открывается):', f, e)
            continue
        names = [hn(x) for x in wb.sheetnames]
        first = None
        for r in wb.worksheets[0].iter_rows(values_only=True, max_row=3):
            if any(c not in (None, '') for c in r):
                first = [hn(c) for c in r]
                break
        if 'баланс' in names and any('поступление' in x for x in names):
            bal.append(f)
        elif first and 'год плана закупа' in first:
            me.append(f)
        else:
            print('пропуск (не похож ни на баланс, ни на ME2N):', f)
    pick = lambda L: max(L, key=commit_time) if L else None
    return pick(bal), pick(me), bal + me


def build_balance(f):
    wb = load_workbook(f, read_only=True, data_only=True)
    out = {}
    R = rows(sheet_by_name(wb, 'баланс')); H = header_map(R)
    c = {k: col(H, k) for k in ['Материал', 'ДлинТекстМатер', 'Базовая ЕИ', 'Обозначение склада', 'Доступный запас',
                                'Партия', 'Имя 1', 'Складское место', 'Склад', 'Старый номер КБМ']}
    out['bal'] = [[s(r[c['Материал']]), s(r[c['ДлинТекстМатер']]), s(r[c['Базовая ЕИ']]), s(r[c['Обозначение склада']]),
                   n(r[c['Доступный запас']]), s(r[c['Партия']]), s(r[c['Имя 1']]), s(r[c['Складское место']]),
                   s(r[c['Склад']]), s(r[c['Старый номер КБМ']])]
                  for r in R[1:] if s(r[c['Материал']])]

    R = rows(sheet_by_name(wb, 'поступление (вне баланса)')); H = header_map(R)
    c = dict(code=col(H, 'Код материала'), desc=col(H, 'Описание по БД SAP'), unit=col(H, 'Ед. изм. заказа по БД SAP'),
             qty=col(H, 'Кол-во получено'), date=col(H, 'Дата оформления Акта первичной приемки'),
             dog=col(H, 'Договор №'), po=col(H, 'PO №'), shelf=col(H, 'ТМЦ размещено на полке'),
             act=col(H, 'Акт первичной приемки №'), note=col(H, 'Примечание'))
    out['inc'] = [[s(r[c['code']]), s(r[c['desc']]), s(r[c['unit']]), n(r[c['qty']]), d(r[c['date']]),
                   s(r[c['dog']]), s(r[c['po']]), s(r[c['shelf']]), s(r[c['act']]), s(r[c['note']])] for r in R[1:]]

    R = rows(sheet_by_name(wb, 'выдача (МТ)')); H = header_map(R)
    c = dict(date=col(H, 'Дата документа'), ini=col(H, 'Инициатор'), wh=col(H, 'Склад'), code=col(H, 'Материал'),
             name=col(H, 'Наименование'), unit=col(H, 'ЕИ'), qty=col(H, 'кол-во'), why=col(H, 'Причина'),
             note=col(H, 'Примечание'))
    out['iss'] = [[d(r[c['date']]), s(r[c['ini']]), s(r[c['wh']]), s(r[c['code']]), s(r[c['name']]), s(r[c['unit']]),
                   n(r[c['qty']]), s(r[c['why']]), s(r[c['note']])] for r in R[1:] if s(r[c['code']])]

    R = rows(sheet_by_name(wb, 'баланс+поступление-выдача МТ'))
    out['tot'] = [[s(r[0]), s(r[1]), n(r[2])] for r in R[1:] if s(r[0])]

    R = rows(sheet_by_name(wb, 'баланс Цехов (Virtual)')); H = header_map(R)
    c = {k: col(H, k) for k in ['Материал', 'Краткий текст материала', 'Базовая ЕИ', 'Обозначение склада',
                                'Доступный запас', 'Партия', 'Имя 1']}
    out['vir'] = [[s(r[c['Материал']]), s(r[c['Краткий текст материала']]), s(r[c['Базовая ЕИ']]),
                   s(r[c['Обозначение склада']]), n(r[c['Доступный запас']]), s(r[c['Партия']]), s(r[c['Имя 1']])]
                  for r in R[1:] if s(r[c['Материал']])]

    # актуальные резервы цехов (лист ищется по слову «резерв» в названии: «актуальные резервы цехов 2026», в 2027 название сменится).
    # Обязательны только «Обозначение склада» и «Материал»; остальные колонки необязательны (их состав в файле меняется).
    # Если лист не прочитался, остаются прежние данные резервов, а остальное всё равно обновляется.
    res_name = next((nm for nm in wb.sheetnames if 'резерв' in hn(nm)), None)
    if res_name is None:
        print('ВНИМАНИЕ: лист резервов не найден, остаются прежние данные резервов')
    else:
        try:
            R = rows(wb[res_name]); H = header_map(R)
            c = dict(shop=col(H, 'Обозначение склада'), code=col(H, 'Материал'),
                     plan=colo(H, 'Имя плановика'), wh=colo(H, 'Принимающий склад'), no=colo(H, 'Резервирование'),
                     name=colo(H, 'Краткий текст материала'), unit=colo(H, 'Базовая ЕИ'),
                     need=col(H, 'Объем потребности'), red=colo(H, 'Сокращ. количество'),
                     diff=col(H, 'Размер разницы'), mv=colo(H, 'Вид движения'), date=colo(H, 'Дата потребности'))
            out['res'] = [[gv(r, c['shop']), gv(r, c['plan']), gv(r, c['wh']), gv(r, c['no']), gv(r, c['code']),
                           gv(r, c['name']), gv(r, c['unit']), gv(r, c['need'], n), gv(r, c['red'], n),
                           gv(r, c['diff'], n), gv(r, c['mv']), gv(r, c['date'], d)]
                          for r in R[1:] if gv(r, c['code'])]
        except Exception as e:
            print('ВНИМАНИЕ: лист резервов не прочитан, остаются прежние данные резервов:', e)

    m = re.search(r'(\d{2})[_.-](\d{2})[_.-](\d{4})', os.path.basename(f))
    out['asof'] = f'{m.group(1)}.{m.group(2)}.{m.group(3)}' if m else \
        datetime.datetime.fromtimestamp(commit_time(f)).strftime('%d.%m.%Y')
    return out


def build_me(f):
    wb = load_workbook(f, read_only=True, data_only=True)
    R = rows(wb.worksheets[0]); H = header_map(R)
    c = dict(plan=col(H, 'Год плана закупа'), sup=col(H, 'Имя поставщика'), st=col(H, 'Статус поставки'),
             mat=col(H, 'Материал'), txt=col(H, 'Краткий текст'), unit=col(H, 'ЕИ заказа на постав.'),
             ord=col(H, 'Объем заказа'), got=col(H, 'Поступило'), rem=col(H, 'еще поставить (количество)'),
             dd=col(H, 'Дата поставки'), act=col(H, 'Дата акта ответхранения'), why=col(H, 'Причина ответхранения'))
    me = [[s(r[c['plan']]), re.sub(r'^\d+\s+', '', s(r[c['sup']])), s(r[c['st']]), s(r[c['mat']]), s(r[c['txt']]),
           s(r[c['unit']]), n(r[c['ord']]), n(r[c['got']]), n(r[c['rem']]), d(r[c['dd']]), d(r[c['act']]),
           s(r[c['why']])] for r in R[1:] if s(r[c['mat']])]
    m = re.search(r'(\d{2})[_.-](\d{2})', os.path.basename(f))
    asof = f'{m.group(1)}.{m.group(2)}' if m else \
        datetime.datetime.fromtimestamp(commit_time(f)).strftime('%d.%m')
    return {'me': me, 'me_asof': asof}


def main():
    data = {}
    if os.path.exists(OUT):
        data = json.load(open(OUT, encoding='utf-8'))
    fb, fm, used = find_files()
    print('баланс:', fb and os.path.basename(fb), '| ME2N:', fm and os.path.basename(fm))
    if not fb and not fm:
        print('В input/ нет подходящих файлов, data.json не меняю.')
        return 0
    if fb:
        data.update(build_balance(fb))
    if fm:
        data.update(build_me(fm))
    with open(OUT, 'w', encoding='utf-8') as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(',', ':'))
    print({k: len(v) for k, v in data.items() if isinstance(v, list)}, data.get('asof'), data.get('me_asof'))
    # как в КБМ: после успешной сборки input/ очищается (README и нераспознанные файлы остаются)
    for f in used:
        try:
            os.remove(f)
            print('удалён из input/:', os.path.basename(f))
        except OSError as e:
            print('не удалось удалить', f, e)
    return 0


if __name__ == '__main__':
    sys.exit(main())
