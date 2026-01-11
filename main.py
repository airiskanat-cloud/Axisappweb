# main.py
"""
Точка входа backend-логики.
Используется app/app.py (Streamlit).
"""

from calculations.final import run_calculation


def main():
    """
    Заглушка для локального запуска без Streamlit.
    Можно использовать для тестов.
    """
    test_order = {
        "meta": {
            "order_number": "TEST-001"
        },
        "common": {
            "product_type": "window_fixed",
            "profile_system": "ALG 63",
            "glass_type": "double",
            "panel_type": None,
            "assembly": True,
            "installation": False
        },
        "positions": [
            {
                "width": 1200,
                "height": 1400,
                "imposts": {
                    "left": 0,
                    "right": 0,
                    "center": 0,
                    "top": 0
                },
                "sashes": []
            }
        ],
        "facade": None
    }

    result = run_calculation(test_order)
    print("RESULT:")
    print(result)


if __name__ == "__main__":
    main()
