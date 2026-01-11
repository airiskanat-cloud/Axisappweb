# calculations/test_geometry.py

from calculations.geometry import (
    geometry_positions,
    geometry_positions_extended
)

from calculations.facade import facade_glass_and_panels
from calculations.materials import materials_positions, materials_facade
from calculations.pricing import price_windows_doors, price_facade

from references.sheets_reader import load_reference_1, load_reference_2
from config.settings import GOOGLE_CREDENTIALS_PATH, SPREADSHEET_ID


# ======================================================
# ТЕСТОВЫЕ ДАННЫЕ — ОКНО
# ======================================================

positions = [
    {
        "width": 2000,
        "height": 1560,
        "imposts": {
            "left": 1000,
            "center": 0,
            "right": 1000,
            "tor": 560
        },
        "sashes": [
            {"width": 1000, "height": 1000},
            {"width": 1000, "height": 1560}
        ]
    }
]

print("Базовая геометрия:")
geo_base = geometry_positions(positions)
print(geo_base)

print("\nРасширенная геометрия:")
geo_ext = geometry_positions_extended(positions)
print(geo_ext)


# ======================================================
# ТЕСТОВЫЕ ДАННЫЕ — ФАСАД
# ======================================================

facade_test = {
    "width": 6000,
    "height": 3000,
    "grid": {"cols": 4, "rows": 3},
    "inserts": [{}, {}],   # 2 вставки
    "panels": [{}]         # 1 панель
}

print("\nФасад — площади:")
facade_areas = facade_glass_and_panels(facade_test)
print(facade_areas)


# ======================================================
# МАТЕРИАЛЫ
# ======================================================

print("\nМатериалы — окна / двери:")
materials_win = materials_positions(geo_ext)
print(materials_win)

print("\nМатериалы — фасад:")
materials_fac = materials_facade(facade_test)
print(materials_fac)


# ======================================================
# СПРАВОЧНИКИ
# ======================================================

print("\nСправочник №1:")
ref1 = load_reference_1(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
print(list(ref1.keys())[:5])

print("\nСправочник №2:")
ref2 = load_reference_2(SPREADSHEET_ID, GOOGLE_CREDENTIALS_PATH)
print(list(ref2.keys())[:5])


# ======================================================
# ПРОСТАЯ ЦЕНА (ШАГ 6.1)
# ======================================================

print("\nЦена — окна / двери:")
print(price_windows_doors(
    geometry=geometry_positions(positions),
    materials=materials_positions(
        geometry_positions_extended(positions)
    ),
    ref2=ref2,
    glass_type="двойной",
    profile_system="ALG 2030-63C"
))
print("\nDEBUG ref2['двойной']:")
print(ref2["двойной"])

from calculations.pricing import price_options_windows

print("\nОпции — окна / двери:")
print(
    price_options_windows(
        geometry=geometry_positions(positions),
        ref2=ref2,
        glass_type="двойной",
        toning="Есть",
        assembly="Есть",
        installation="Монтаж"
    )
)

from calculations.final import calc_windows_doors_final

print("\nФИНАЛ — окна / двери:")
print(
    calc_windows_doors_final(
        positions=positions,
        ref2=ref2,
        glass_type="двойной",
        profile_system="ALG 2030-73C",
        toning="Есть",
        assembly="Есть",
        installation="Монтаж"
    )
)
from calculations.final import calc_facade_final

print("\nФИНАЛ — ФАСАД:")
print(
    calc_facade_final(
        facade_data=facade_test,
        ref2=ref2,
        glass_type="двойной",
        toning="Есть",
        assembly="Есть",
        installation="Монтаж"
    )
)

