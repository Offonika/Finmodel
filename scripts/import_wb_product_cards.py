from __future__ import annotations

import logging
import os
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
import xlwings as xw
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

SETTINGS_SHEET = 'НастройкиОрганизаций'
PRODUCTS_SHEET = 'Номенклатура_WB'
HEADERS = [
    'Организация', 'Артикул_WB', 'Артикул_поставщика',
    'Бренд', 'Название', 'Предмет',
    'Ширина', 'Высота', 'Длина', 'Вес_брутто', 'Объем_литр'
]
API_URL = 'https://content-api.wildberries.ru/content/v2/get/cards/list?locale=ru'
LIMIT = 100


def _configure_logger() -> tuple[logging.Logger, Path | None]:
    logger = logging.getLogger(__name__)
    if logger.handlers:
        log_path = getattr(logger, '_log_file_path', None)
        return logger, Path(log_path) if log_path else None

    logger.setLevel(logging.DEBUG)
    log_dir = Path(__file__).resolve().parent / 'log'
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_path = log_dir / f'import_wb_product_cards_{timestamp}.log'

    file_handler = logging.FileHandler(log_path, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    logger.propagate = False
    setattr(logger, '_log_file_path', str(log_path))
    logger.debug('Logging configured | log_path=%s', log_path)
    return logger, log_path


LOGGER, LOG_PATH = _configure_logger()


def _session_with_retries():
    session = requests.Session()
    retry = Retry(
        total=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 501, 502, 503, 504],
        allowed_methods=frozenset(["GET", "POST"]),
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('https://', adapter)
    session.mount('http://', adapter)
    return session


def get_idx(header_row):
    return {h.strip(): i for i, h in enumerate(header_row)}


@dataclass
class TokenCleanResult:
    token: str | None
    error: str | None
    details: dict[str, Any] = field(default_factory=dict)


def _mask_value(value: str) -> str:
    if not value:
        return ''
    if len(value) <= 2:
        return '*' * len(value)
    return f"{value[0]}{'*' * (len(value) - 2)}{value[-1]}"


def _describe_non_ascii(value: str):
    return [
        {
            'index': idx,
            'char': ch,
            'ord': ord(ch),
        }
        for idx, ch in enumerate(value)
        if not ch.isascii()
    ]


def _clean_token(token_raw):
    details: dict[str, Any] = {'original': token_raw}

    if token_raw is None:
        details['reason'] = 'missing'
        return TokenCleanResult(token=None, error='Token_WB отсутствует', details=details)

    raw_string = str(token_raw)
    details['raw'] = raw_string
    chars_to_remove = " \t\r\n\"'\u00a0\ufeff"
    translation_table = str.maketrans('', '', chars_to_remove)
    cleaned = raw_string.translate(translation_table).strip()
    details['cleaned'] = cleaned

    if cleaned.casefold() in {'', 'nan', 'none'}:
        details['reason'] = 'empty_after_cleaning'
        return TokenCleanResult(token=None, error='Token_WB пустой после очистки', details=details)

    if not cleaned.isascii():
        offending = _describe_non_ascii(cleaned)
        masked = _mask_value(cleaned)
        details.update({
            'reason': 'non_ascii',
            'masked_token': masked,
            'length': len(cleaned),
            'offending_characters': offending,
        })
        LOGGER.error(
            'Token содержит недопустимые символы (не-ASCII). Маска=%s длина=%s offending=%s',
            masked,
            len(cleaned),
            offending,
        )
        return TokenCleanResult(token=None, error='Token_WB содержит недопустимые символы', details=details)

    details.update({
        'masked_token': _mask_value(cleaned),
        'length': len(cleaned),
    })
    return TokenCleanResult(token=cleaned, error=None, details=details)


def _sanitize_token_details(details: dict[str, Any]) -> dict[str, Any]:
    sensitive_keys = {'cleaned', 'raw', 'original'}
    return {k: v for k, v in details.items() if k not in sensitive_keys}


def _validate_headers_ascii(headers, *, context: str = 'HTTP headers'):
    for key, value in headers.items():
        for label, text in (('ключ', key), ('значение', value)):
            text_str = '' if text is None else str(text)
            if not text_str.isascii():
                masked = _mask_value(text_str)
                offending = _describe_non_ascii(text_str)
                target = 'ключ заголовка' if label == 'ключ' else 'значение заголовка'
                LOGGER.error(
                    '%s: %s %r содержит недопустимые символы. Маска=%s длина=%s offending=%s',
                    context,
                    target,
                    key,
                    masked,
                    len(text_str),
                    offending,
                )
                return False
    return True

def main():
    LOGGER.info('=== START import_wb_product_cards ===')
    wb = xw.Book.caller()  # <-- ВАЖНО!
    sht_set = wb.sheets[SETTINGS_SHEET]

    # --- Подготовка листа с товарами ---
    sheet_names = [sht.name for sht in wb.sheets]
    if PRODUCTS_SHEET not in sheet_names:
        sht_prod = wb.sheets.add(PRODUCTS_SHEET, after=wb.sheets[wb.sheets.count-1])
        LOGGER.info('Создан новый лист для выгрузки карточек | sheet=%s', PRODUCTS_SHEET)
    else:
        sht_prod = wb.sheets[PRODUCTS_SHEET]
        LOGGER.info('Используется существующий лист для выгрузки карточек | sheet=%s', PRODUCTS_SHEET)

    sht_prod.clear()
    sht_prod.range('A1').value = HEADERS

    hdr_rng = sht_prod.range((1, 1), (1, len(HEADERS)))
    hdr_rng.api.Font.Bold = True
    hdr_rng.api.HorizontalAlignment = -4108  # xlCenter
    hdr_rng.api.Borders.Weight = 2           # xlThin

    for col in range(1, len(HEADERS) + 1):
        sht_prod.range((1, col)).api.EntireColumn.AutoFit()
    LOGGER.debug('Автоподбор ширины колонок выполнен')

    cfgHdr = sht_set.range('A1').expand('right').value
    LOGGER.debug('Загружена шапка листа настроек | header=%s', cfgHdr)
    idx = get_idx(cfgHdr)
    if 'Организация' not in idx or 'Token_WB' not in idx:
        LOGGER.error('В листе настроек отсутствуют обязательные колонки «Организация» и/или «Token_WB»')
        return

    org_col = idx['Организация']
    org_values = sht_set.range((2, org_col+1), (sht_set.cells.last_cell.row, org_col+1)).options(ndim=1).value
    org_rows_count = next((i for i, val in enumerate(org_values) if not val), len(org_values))
    if org_rows_count == 0:
        LOGGER.info('Нет организаций для обработки')
        return

    last_col = len(cfgHdr)
    settings = sht_set.range((2,1), (org_rows_count+1, last_col)).value

    LOGGER.info(
        'Стартовое окружение | cwd=%s python=%s log_path=%s detected_ranges=%s',
        os.getcwd(),
        sys.version.replace('\n', ' '),
        str(LOG_PATH) if LOG_PATH else '',
        {
            'settings_header': sht_set.range('A1').expand('right').address,
            'organizations_column': sht_set.range(
                (2, org_col + 1),
                (org_rows_count + 1, org_col + 1),
            ).address,
            'products_header': sht_prod.range((1, 1), (1, len(HEADERS))).address,
        },
    )

    allCards = []
    for i, row in enumerate(settings):
        org = row[idx['Организация']]
        token_raw = row[idx['Token_WB']]
        org_name = '' if org is None else str(org)
        LOGGER.info('Начало обработки организации | organization=%s row_index=%s', org_name, i + 2)
        token_result = _clean_token(token_raw)
        safe_details = _sanitize_token_details(token_result.details)
        if token_result.error:
            LOGGER.error(
                'Некорректный Token_WB у организации | organization=%s reason=%s details=%s',
                org_name,
                token_result.error,
                safe_details,
            )
            continue
        token_clean = token_result.token
        LOGGER.debug(
            'Token очищен | organization=%s masked=%s length=%s details=%s',
            org_name,
            safe_details.get('masked_token'),
            safe_details.get('length'),
            safe_details,
        )
        if not org:
            LOGGER.warning(
                'Пропуск строки из-за отсутствия значения в колонке «Организация» | organization=%s',
                org_name,
            )
            continue
        cursor = None
        page = 0
        existSet = set()
        session = None
        headers = {
            'Authorization': token_clean,
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'User-Agent': 'FinmodelWB/1.0',
        }
        if not _validate_headers_ascii(headers, context=f'HTTP-заголовки для организации "{org_name}"'):
            LOGGER.warning(
                'Пропуск организации из-за некорректных HTTP-заголовков | organization=%s',
                org_name,
            )
            continue
        LOGGER.debug('HTTP-заголовки прошли проверку | organization=%s', org_name)
        try:
            session = _session_with_retries()
            while True:
                page += 1
                payload = {
                    'settings': {
                        'cursor': cursor if cursor else {'limit': LIMIT},
                        'filter': {'withPhoto': -1}
                    }
                }
                LOGGER.debug(
                    'Параметры пагинации перед запросом | organization=%s page=%s cursor=%s limit=%s',
                    org_name,
                    page,
                    payload['settings']['cursor'],
                    LIMIT,
                )
                try:
                    resp = session.post(API_URL, json=payload, headers=headers, timeout=60)
                except Exception as e:
                    backoff_seconds = 10
                    LOGGER.warning(
                        'Сетевая ошибка при обращении к API | organization=%s page=%s backoff=%s error=%s',
                        org_name,
                        page,
                        backoff_seconds,
                        e,
                    )
                    time.sleep(backoff_seconds)
                    continue

                LOGGER.info(
                    'Ответ API | organization=%s page=%s status_code=%s',
                    org_name,
                    page,
                    resp.status_code,
                )
                if resp.status_code != 200:
                    LOGGER.error(
                        'Завершение обработки организации из-за статуса API | organization=%s page=%s status_code=%s response=%s',
                        org_name,
                        page,
                        resp.status_code,
                        resp.text,
                    )
                    break
                data = resp.json()
                cards = data.get('cards', [])
                LOGGER.info(
                    'Получены карточки | organization=%s page=%s count=%s',
                    org_name,
                    page,
                    len(cards),
                )
                for c in cards:
                    nm = str(c.get('nmID', ''))
                    if nm and nm not in existSet:
                        width = c.get('dimensions', {}).get('width', '')
                        height = c.get('dimensions', {}).get('height', '')
                        length = c.get('dimensions', {}).get('length', '')
                        # Считаем объем, если все размеры есть и являются числами
                        try:
                            vol_ltr = float(width) * float(height) * float(length) / 1000
                            vol_ltr = round(vol_ltr, 3)
                        except Exception:
                            vol_ltr = ''
                        allCards.append([
                            org,
                            nm,
                            c.get('vendorCode', ''),
                            c.get('brand', ''),
                            c.get('title', ''),
                            c.get('subjectName', ''),
                            width, height, length,
                            c.get('dimensions', {}).get('weightBrutto', ''),
                            vol_ltr
                        ])
                        existSet.add(nm)

                cur = data.get('cursor', {})
                if cur.get('total') is None or cur.get('total', 0) < LIMIT:
                    LOGGER.info(
                        'Завершение пагинации для организации | organization=%s page=%s reason=%s',
                        org_name,
                        page,
                        'cursor_total_below_limit',
                    )
                    break
                cursor = {k: cur[k] for k in ('updatedAt','nmID') if k in cur}
                cursor['limit'] = LIMIT
        finally:
            if session is not None:
                session.close()

    if allCards:
        sht_prod.range((2, 1)).value = allCards
        sht_prod.range('B:B').api.NumberFormat = '@'
        LOGGER.info('Добавлено новых карточек | count=%s', len(allCards))
    else:
        LOGGER.info('Новых карточек не найдено')

    LOGGER.info('=== END import_wb_product_cards ===')

if __name__ == '__main__':
    pass
