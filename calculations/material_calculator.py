"""
Material Calculator - Formula-Based System (FULL VERSION)
Полный расчёт материалов как в engine_windows.py
С формулами из справочника + правильный итоговый расчёт
"""

import math
from typing import Dict, List, Any
from .product_model import (
    Product, ProductGeometry, ProductMaterials,
    FrameMaterial, SealMaterial, HardwareItem,
    UsageMode, ProductType
)


def safe_eval(formula: str, context: dict) -> float:
    """Безопасное вычисление формул Python"""
    try:
        f = str(formula).replace(",", ".").replace(" ", "")
        return float(eval(f, {"__builtins__": None, "math": math}, context))
    except Exception as e:
        print(f"⚠️ Ошибка в формуле '{formula}': {e}")
        return 0.0


def safe_float(value, default=0.0):
    """Безопасное преобразование в float"""
    try:
        if value is None:
            return default
        s = str(value).replace(",", ".").replace(" ", "").replace("\xa0", "")
        if s == "":
            return default
        return float(s)
    except:
        return default


class MaterialCalculator:
    """
    Калькулятор материалов - ПОЛНАЯ ВЕРСИЯ
    
    Как в engine_windows.py:
    1. Формулы из справочника
    2. Округление до упаковок
    3. part2_materials с qty_fact, norm, qty_ship
    4. part3_final со всеми категориями
    5. Использует CODE для точного поиска
    """
    
    def __init__(self, ref1: List[Dict], ref2: Dict[str, float], ref3: List[Dict]):
        self.ref1 = ref1
        self.ref2 = ref2
        self.ref3 = ref3
    
    @staticmethod
    def _get_price(item: Dict, default: float = 0.0) -> float:
        """Извлекает цену с поддержкой всех вариантов колонок"""
        price = item.get("Цена за единицу",
                item.get("цена за ед.",
                item.get("цена за ед ",
                item.get("Цена", default))))
        return safe_float(price, default)
    
    def calculate_materials_full(
        self,
        product_data: Dict,
        usage_mode: UsageMode = UsageMode.STANDALONE
    ) -> Dict:
        """
        ПОЛНЫЙ расчёт как в engine_windows.py
        
        Возвращает:
        {
            "part1_gabarits": [],
            "part2_materials": [],
            "part3_final": {...},
            "metrics": {...},
            "total_with_margin": float
        }
        """
        from .product_model import create_product_from_form_data
        
        # 1. Создание модели изделия
        product = create_product_from_form_data(
            product_type=product_data.get("product_type", "Окно"),
            system=product_data.get("system", "ALG 2030-45C"),
            data=product_data.get("data", {}),
            usage_mode=usage_mode,
            code=product_data.get("code", "")
        )
        
        print(f"\n{'='*70}")
        print(f"🔧 ПОЛНЫЙ РАСЧЁТ МАТЕРИАЛОВ")
        print(f"{'='*70}")
        print(f"Тип: {product.product_type.value}")
        print(f"Система: {product.system}")
        print(f"CODE: {product.code}")
        print(f"Габариты: {product.geometry.width_m}м × {product.geometry.height_m}м")
        
        # 2. Контекст для формул
        context = self._create_formula_context(product)
        
        print(f"\n📊 Контекст:")
        print(f"   W={context['W']:.2f}м, H={context['H']:.2f}м")
        print(f"   Периметр={context['perimeter']:.2f}м, Площадь={context['area']:.2f}м²")
        
        # 3. Расчёт материалов по формулам
        materials_dict = {}
        
        print(f"\n🔍 Поиск материалов по CODE={product.code}...")
        
        for row in self.ref1:
            row_code = str(row.get("CODE") or row.get("code") or "").strip()
            
            if row_code and product.code and row_code == product.code:
                formula = row.get("Формула_Python", "")
                if not formula:
                    formula = row.get("формула фактического расхода", "")
                if not formula:
                    continue
                
                qty_fact = safe_eval(formula, context)
                
                if qty_fact <= 0:
                    continue
                
                товар = str(row.get("Товар", ""))
                артикул = str(row.get("Артикул", ""))
                key = f"{товар}|{артикул}"
                
                if key not in materials_dict:
                    materials_dict[key] = {
                        "товар": товар,
                        "артикул": артикул,
                        "тип_эл": row.get("Тип элемента", row.get("тип элемент", "")),
                        "qty_fact": 0,
                        "norm": safe_float(row.get("кол-во норм к упаковке", 1)),
                        "price": self._get_price(row, 0),
                        "unit": str(row.get("Ед.", "шт"))
                    }
                
                materials_dict[key]["qty_fact"] += qty_fact
        
        print(f"Найдено: {len(materials_dict)} позиций")
        
        # 4. Округление до упаковок (как в engine_windows)
        part2_materials = []
        materials_sum = 0.0
        
        print(f"\n💰 Округление до упаковок:")
        
        for key, mat in materials_dict.items():
            qty_fact = mat["qty_fact"]
            norm = mat["norm"]
            price = mat["price"]
            
            # ФОРМУЛА ИЗ engine_windows.py (строки 339-344):
            if norm > 0:
                qty_ship = math.ceil(qty_fact / norm)
            else:
                qty_ship = math.ceil(qty_fact)
            
            row_sum = (price * norm) * qty_ship
            materials_sum += row_sum
            
            if qty_fact > 0:
                part2_materials.append({
                    "Товар": mat["товар"],
                    "Артикул": mat["артикул"],
                    "Тип элемента": mat["тип_эл"],
                    "Цена": price,
                    "Ед.": mat["unit"],
                    "Расход факт.": round(qty_fact, 2),
                    "Норма": norm,
                    "К отгрузке": qty_ship,
                    "Сумма": round(row_sum, 0)
                })
                
                print(f"   {mat['тип_эл']}: {qty_fact:.2f}{mat['unit']} → {qty_ship} упак × {norm}{mat['unit']} = {row_sum:,.0f}₸")
        
        print(f"\nИТОГО МАТЕРИАЛЫ: {materials_sum:,.0f}₸")
        
        # 5. Расчёт стекла/ламбри
        total_area = context["area"]
        
        cost_glass = 0.0
        cost_lambri = 0.0
        
        data = product_data.get("data", {})
        fill_category = data.get("fill_category", "Стеклопакет")
        glass_type = data.get("glass_type", "Двойной")
        
        print(f"\n💎 Заполнение: {fill_category}")
        
        if fill_category == "Стеклопакет":
            price_glass = self._get_price_from_ref2(glass_type)
            cost_glass = total_area * price_glass
            print(f"   Стеклопакет: {total_area:.3f}м² × {price_glass:,.0f}₸/м² = {cost_glass:,.0f}₸")
        elif "Ламбри" in fill_category:
            price_lambri = self._get_price_from_ref2(fill_category)
            qty_hlysti = math.ceil(total_area / 6) if total_area > 0 else 0
            total_meters = qty_hlysti * 6
            cost_lambri = total_meters * price_lambri
            print(f"   Ламбри: {total_area:.3f}м² → {qty_hlysti} хлыстов × 6м = {total_meters}м × {price_lambri:,.0f}₸/м = {cost_lambri:,.0f}₸")
        
        # 6. Тонировка, сборка, монтаж
        cost_toning = 0.0
        cost_assembly = 0.0
        cost_installation = 0.0
        
        # ИСПРАВЛЕНО: Берём из common (для окон/дверей) или data (для вставок)
        common = product_data.get("common", {})
        
        toning = common.get("toning") or data.get("toning", "Нет")
        if toning == "Есть":
            price_toning = self._get_price_from_ref2("Тонировка")
            cost_toning = total_area * price_toning
            print(f"   ✅ Тонировка: {total_area:.3f}м² × {price_toning:,.0f}₸/м² = {cost_toning:,.0f}₸")
        
        assembly = common.get("assembly") or data.get("assembly", "Нет")
        if assembly == "Есть":
            price_assembly = self._get_price_from_ref2("Сборка")
            cost_assembly = total_area * price_assembly
            print(f"   ✅ Сборка: {total_area:.3f}м² × {price_assembly:,.0f}₸/м² = {cost_assembly:,.0f}₸")
        
        installation = common.get("installation") or data.get("installation", "Нет")
        if installation != "Нет":
            installation_clean = " ".join(installation.split())
            price_installation = self._get_price_from_ref2(installation_clean)
            cost_installation = total_area * price_installation
            print(f"   ✅ Монтаж: {total_area:.3f}м² × {price_installation:,.0f}₸/м² = {cost_installation:,.0f}₸")
        
        # 7. Дополнительные детали (нащельник)
        cost_additional = 0.0
        total_perimeter = context["perimeter"]
        
        additional_name = None
        for key in self.ref2.keys():
            if "нащельник" in key.lower():
                additional_name = key
                break
        
        if additional_name:
            price_additional = self.ref2.get(additional_name, 0)
            cost_additional = math.ceil(total_perimeter / 3) * price_additional
            print(f"\n🔧 Доп. детали: ⌈{total_perimeter:.2f}/3⌉ × {price_additional:,.0f}₸ = {cost_additional:,.0f}₸")
        
        # 8. ИТОГОВЫЙ РАСЧЁТ (как в engine_windows строки 530-543)
        part3_final = {
            "Стеклопакет": round(cost_glass, 0),
            "Ламбри": round(cost_lambri, 0),
            "Тонировка": round(cost_toning, 0),
            "Сборка": round(cost_assembly, 0),
            "Монтаж": round(cost_installation, 0),
            "Дополнительные детали": round(cost_additional, 0),
            "Материалы": round(materials_sum, 0)
        }
        
        subtotal = sum(part3_final.values())
        margin = subtotal * 0.81
        part3_final["Обеспечение"] = round(margin, 0)
        total_with_margin = round(subtotal + margin, 0)
        
        print(f"\n{'='*70}")
        print(f"📊 ИТОГОВЫЙ РАСЧЁТ:")
        for key, val in part3_final.items():
            if val > 0:
                print(f"   {key}: {val:,.0f}₸")
        print(f"{'='*70}")
        print(f"К ОПЛАТЕ: {total_with_margin:,.0f}₸")
        print(f"{'='*70}\n")
        
        # 9. Возвращаем в формате engine_windows
        return {
            "part1_gabarits": [],  # Пока пусто, можно добавить позже
            "part2_materials": part2_materials,
            "part3_final": part3_final,
            "metrics": {
                "total_area": round(total_area, 3),
                "total_perimeter": round(total_perimeter, 3)
            },
            "total_with_margin": total_with_margin,
            "materials_cost": round(subtotal, 0)  # Для фасадов
        }
    
    def _create_formula_context(self, product: Product) -> Dict:
        """Создаёт контекст для формул"""
        geometry = product.geometry
        
        W = geometry.width_m
        H = geometry.height_m
        
        n_sash = len(geometry.sashes) if geometry.sashes else 1  # глухое окно = 1 пакет
        
        if geometry.sashes:
            w_s = sum(s.width for s in geometry.sashes) / len(geometry.sashes) / 1000
            h_s = sum(s.height for s in geometry.sashes) / len(geometry.sashes) / 1000
            w_s_total = sum(s.width for s in geometry.sashes) / 1000
        else:
            w_s = W
            h_s = H
            w_s_total = W
        
        offset = 0.073
        w_g = W - 2 * offset
        h_g = H - 2 * offset
        
        imp_vertical = 1 if geometry.has_vertical_impost else 0
        imp_horizontal = 1 if geometry.has_horizontal_impost else 0
        total_imposts = imp_vertical + imp_horizontal
        
        if h_s * 1000 < 1200:
            n_lp = 2
        elif h_s * 1000 < 2000:
            n_lp = 3
        else:
            n_lp = 4
        
        return {
            "W": W, "H": H, "w": W, "h": H,
            "count": 1, "qty": 1, "Nwin": 1,
            "n_sash": n_sash,
            "w_s": w_s, "h_s": h_s,
            "w_stvor": w_s, "h_stvor": h_s,
            "w_s_total": w_s_total,
            "w_g": w_g, "h_g": h_g,
            "w_glass": w_g, "h_glass": h_g,
            "imp_vertical": imp_vertical,
            "imp_horizontal": imp_horizontal,
            "total_imposts": total_imposts,
            "n_lp": n_lp, "lock_points": n_lp,
            "area": W * H, "area_m2": W * H,
            "perimeter": 2 * (W + H),
            "perimeter_m": 2 * (W + H)
        }
    
    def _get_price_from_ref2(self, key_word: str) -> float:
        """Поиск цены в Справочнике-2"""
        key_normalized = key_word.lower().replace(" / ", "/")
        price = self.ref2.get(key_normalized)
        
        if price is None:
            defaults = {
                "ламбри без термо": 2248,
                "ламбри с термо": 2800,
                "двойной": 9500,
                "тройной": 12000
            }
            price = defaults.get(key_normalized)
        
        if price is None:
            return 0.0
        return float(price)


def calculate_product_materials(
    product_data: Dict,
    ref1: List[Dict],
    ref2: Dict[str, float],
    ref3: List[Dict],
    usage_mode: UsageMode = UsageMode.STANDALONE
) -> Dict:
    """
    ПОЛНАЯ функция расчёта материалов
    
    Возвращает результат в формате engine_windows.py:
    {
        "part1_gabarits": [...],
        "part2_materials": [...],
        "part3_final": {...},
        "metrics": {...},
        "total_with_margin": float
    }
    """
    calculator = MaterialCalculator(ref1, ref2, ref3)
    return calculator.calculate_materials_full(product_data, usage_mode)
