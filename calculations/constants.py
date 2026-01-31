"""
Единый стандарт именования ключей (ТЗ V.9 — Этап 3)

ВСЕ обращения к словарям в проекте идут через эти константы.
Запрещено использовать строковые литералы для ключей словарей изделий/материалов.
"""


class ProductKeys:
    """Ключи для позиций изделий"""
    TYPE            = "product_type"       # Тип изделия
    SYSTEM          = "system_id"          # Система профиля
    WIDTH           = "width"              # Ширина (мм)
    HEIGHT          = "height"             # Высота (мм)
    COUNT           = "count"              # Количество
    FILL_CATEGORY   = "fill_category"      # Тип заполнения (Стеклопакет / Ламбри)
    GLASS_TYPE      = "glass_type"         # Тип стеклопакета
    TONING          = "toning"             # Тонировка (Есть / Нет)
    ASSEMBLY        = "assembly"           # Сборка
    INSTALLATION    = "installation"       # Монтаж
    IMPOSTS         = "imposts"            # Импосты
    SASHES          = "sashes"             # Створки
    SASH_COUNT      = "sash_count"         # Количество створок
    CODE            = "code"               # CODE для поиска в Справочнике-1


class MaterialKeys:
    """Ключи для позиций материалов (Справочник-1)"""
    ARTICLE         = "Артикул"
    ELEMENT         = "Элемент"
    SYSTEM          = "Система"
    CODE            = "CODE"
    UNIT            = "Единица"
    PRICE           = "Цена за единицу"
    FORMULA         = "Формула"
    PACKAGE_SIZE    = "Кратность"          # Кратность упаковки (хлыст)


class ServiceKeys:
    """Ключи для Справочника-2 (услуги и стеклопакеты)"""
    LAMBRI_NO_THERMO  = "ламбри без термо"
    LAMBRI_THERMO     = "ламбри с термо"
    TONING            = "тонировка"
    ASSEMBLY          = "сборка"
    MONTAGE           = "монтаж"
    DEMONTAGE         = "демонтаж"
    DEMONTAGE_MONTAGE = "демонтаж/монтаж"
    COMPLEX_MONTAGE   = "сложный монтаж"
    NASCHEL           = "нащельник"        # Ключ ищется по подстроке


class ResultKeys:
    """Ключи для словарей результатов расчёта"""
    PART1_GABARITS      = "part1_gabarits"
    PART2_MATERIALS     = "part2_materials"
    PART3_FINAL         = "part3_final"
    MATERIALS_COST      = "materials_cost"       # Себестоимость БЕЗ маржи
    TOTAL_WITH_MARGIN   = "total_with_margin"    # DEPRECATED
    METRICS             = "metrics"
    TOTAL_AREA          = "total_area"
    TOTAL_PERIMETER     = "total_perimeter"
    QUANTITY_RAW        = "quantity_raw"          # НЕТТО (до округления)
    QUANTITY            = "Количество"            # БРУТТО (после округления)


class FacadeKeys:
    """Категории корзины MaterialAggregator"""
    FRAME       = "facade_frame"        # Каркас фасада
    INSERTS     = "facade_inserts"      # Вставки (окна/двери в фасаде)
    WINDOWS     = "windows_doors"       # Окна/Двери (автономные)
    TAMBOUR     = "tambour"             # Тамбур


# Артикулы тамбура (для поиска в ref1)
class TambourArticles:
    GUIDE       = "2-00-5581"           # Направляющий
    PIPE        = "2-00-2010"           # Соединительная труба 90°
