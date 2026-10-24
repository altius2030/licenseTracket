"""
Data models and enums for License Tracker.
"""
from enum import Enum
from typing import List


class LicenseStatus(str, Enum):
    ACTIVE = "active"
    EXPIRING = "expiring"
    EXPIRED = "expired"
    MANUAL = "manual"


class ManualStatus(str, Enum):
    ACTIVE = "active"
    EXPIRING = "expiring"
    EXPIRED = "expired"
    NONE = ""


# Field definitions for each sheet type
SHEET_FIELDS = {
    'GBS': [
        ('install_date', 'Дата установки', 'date'),
        ('full_name', 'Фамилия и имя', 'text'),
        ('store_name', 'Название магазина', 'text'),
        ('phone', 'Телефон', 'text'),
        ('email', 'Почта', 'email'),
        ('gbs_id', 'GBS ID', 'text'),
        ('activation_period', 'Срок активации', 'text'),
        ('id_expiry', 'Срок ID', 'date'),
        ('payment_date', 'Дата оплаты', 'date'),
        ('renewal_date', 'Дата продления', 'date'),
        ('next_payment', 'Следующая оплата', 'date'),
        ('address', 'Адрес', 'text'),
    ],
    'ШопУчет': [
        ('full_name', 'ФИО', 'text'),
        ('store_name', 'Магазин', 'text'),
        ('license_key', 'Ключ', 'text'),
        ('payment_date', 'Срок когда оплатил', 'date'),
        ('activation_period', 'Срок активации', 'text'),
        ('comment', 'Комментарий', 'text'),
    ],
    'шоп учет': [
        ('install_date', 'Дата установки', 'date'),
        ('row_number', '№', 'number'),
        ('full_name', 'ФИО', 'text'),
        ('store_name', 'Название магазина', 'text'),
        ('phone', 'Телефон', 'text'),
        ('email', 'Почта', 'email'),
        ('sent_info', 'Отправлено', 'text'),
        ('activation_info', 'Активация и срок', 'text'),
        ('payment_info', 'Оплата', 'text'),
        ('receipt_number', '№ квитанции', 'text'),
        ('migrated_from', 'Переход с другой программы', 'text'),
        ('comment', 'Комментарий', 'text'),
    ],
    'АСТАНА': [
        ('install_date', 'Дата установки', 'date'),
        ('row_number', '№', 'number'),
        ('full_name', 'Фамилия и имя', 'text'),
        ('store_name', 'Название магазина', 'text'),
        ('phone', 'Телефон', 'text'),
        ('email', 'Почта', 'email'),
        ('gbs_id', 'GBS ID', 'text'),
        ('sent_info', 'Отправлено', 'text'),
        ('activation_info', 'Активация и срок', 'text'),
        ('payment_info', 'Оплата', 'text'),
        ('receipt_number', '№ квитанции', 'text'),
        ('migrated_from', 'Переход с другой программы', 'text'),
        ('comment', 'Комментарий', 'text'),
    ],
    'Алматы': [
        ('install_date', 'Дата установки', 'date'),
        ('row_number', '№', 'number'),
        ('full_name', 'Фамилия и имя', 'text'),
        ('store_name', 'Название магазина', 'text'),
        ('phone', 'Телефон', 'text'),
        ('email', 'Почта', 'email'),
        ('gbs_id', 'GBS ID', 'text'),
        ('sent_info', 'Отправлено', 'text'),
        ('activation_info', 'Активация и срок', 'text'),
        ('payment_info', 'Оплата', 'text'),
        ('receipt_number', '№ квитанции', 'text'),
        ('migrated_from', 'Переход с другой программы', 'text'),
        ('comment', 'Комментарий', 'text'),
    ],
    'отвалы': [
        ('install_date', 'Дата установки', 'date'),
        ('row_number', '№', 'number'),
        ('full_name', 'Фамилия и имя', 'text'),
        ('store_name', 'Название магазина', 'text'),
        ('phone', 'Телефон', 'text'),
        ('email', 'Почта', 'email'),
        ('gbs_id', 'GBS ID', 'text'),
        ('sent_info', 'Отправлено', 'text'),
        ('activation_info', 'Активация и срок', 'text'),
        ('payment_info', 'Оплата', 'text'),
        # Many extra columns - we'll store in comment or ignore
    ],
}

# Default fields for unknown sheets
DEFAULT_FIELDS = [
    ('install_date', 'Дата установки', 'date'),
    ('full_name', 'ФИО', 'text'),
    ('store_name', 'Магазин', 'text'),
    ('phone', 'Телефон', 'text'),
    ('email', 'Почта', 'email'),
    ('gbs_id', 'GBS ID', 'text'),
    ('license_key', 'Ключ лицензии', 'text'),
    ('activation_period', 'Срок активации', 'text'),
    ('id_expiry', 'Срок ID', 'date'),
    ('payment_date', 'Дата оплаты', 'date'),
    ('renewal_date', 'Дата продления', 'date'),
    ('next_payment', 'Следующая оплата', 'date'),
    ('address', 'Адрес', 'text'),
    ('sent_info', 'Отправлено', 'text'),
    ('activation_info', 'Активация и срок', 'text'),
    ('payment_info', 'Оплата', 'text'),
    ('receipt_number', '№ квитанции', 'text'),
    ('migrated_from', 'Переход с другой программы', 'text'),
    ('comment', 'Комментарий', 'text'),
]


def get_fields_for_sheet(sheet_name: str) -> List[tuple]:
    """Get field definitions for a sheet."""
    return SHEET_FIELDS.get(sheet_name, DEFAULT_FIELDS)


# Status colors for UI
STATUS_COLORS = {
    LicenseStatus.ACTIVE: "#2ecc71",      # Green
    LicenseStatus.EXPIRING: "#f39c12",    # Orange
    LicenseStatus.EXPIRED: "#e74c3c",     # Red
    LicenseStatus.MANUAL: "#3498db",      # Blue
}

STATUS_LABELS = {
    LicenseStatus.ACTIVE: "Активна",
    LicenseStatus.EXPIRING: "Скоро истекает",
    LicenseStatus.EXPIRED: "Истекла",
    LicenseStatus.MANUAL: "Ручная",
}

MANUAL_STATUS_LABELS = {
    ManualStatus.ACTIVE: "Активна (руч.)",
    ManualStatus.EXPIRING: "Скоро истекает (руч.)",
    ManualStatus.EXPIRED: "Истекла (руч.)",
    ManualStatus.NONE: "Авто",
}