from lxml.etree import Element


def clean_patronymic(ot_obj: Element) -> str:
    """Возвращает обработанное отчество."""
    if ot_obj is None:
        return ''

    ot = ot_obj.text
    if ot.upper() == 'НЕТ':
        return ''
    
    return ot


def clean_phone(phone: str) -> str:
    """Дополняет номер телефона кодом страны - `+7`."""
    if phone.startswith('+7'):
        return phone
    else:
        return f'+7{phone}'