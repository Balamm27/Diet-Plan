from __future__ import annotations

import json
import re
from datetime import datetime
from html import escape
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path("/Users/bala/Documents/Codex/Diet Plan")
WORKBOOK_PATH = ROOT / "Diet Plan.xlsx"
SITE_PATH = ROOT / "index.html"
DATA_PATH = ROOT / "site-data.json"

DAY_NAMES = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
]

PANTRY_BASICS = ["Salt", "Black Pepper", "Water", "Neutral Oil", "Olive Oil"]

CANONICAL_OVERRIDES = {
    "mangoes": "mango",
    "mango": "mango",
    "bananas": "banana",
    "banana": "banana",
    "coocnut water": "coconut water",
    "coconut water": "coconut water",
    "dragon fruit": "dragon fruit",
    "blue berries": "blueberries",
    "blueberries": "blueberries",
    "brocolli": "broccoli",
    "broccoli": "broccoli",
    "chicken": "chicken breast",
    "chicken breast": "chicken breast",
    "chick pea": "chickpeas",
    "chickpeas": "chickpeas",
    "walnut": "walnuts",
    "walnuts": "walnuts",
    "coriander": "coriander",
    "mint": "mint",
    "basil": "basil",
    "thyme": "thyme",
    "sesame seeds": "sesame seeds",
    "coconut oil": "coconut oil",
    "olive oil": "olive oil",
    "navel orange": "navel orange",
    "beans": "green beans",
    "cauliflower": "cauliflower",
    "lady's finger": "okra",
    "lady's finger okra": "okra",
    "okra": "okra",
    "celery juice": "celery",
    "flax seed water": "flax seeds",
    "sweet potatoes": "sweet potatoes",
}

DISPLAY_NAMES = {
    "mango": "Mango",
    "banana": "Banana",
    "coconut water": "Coconut Water",
    "dragon fruit": "Dragon Fruit",
    "blueberries": "Blueberries",
    "chicken breast": "Chicken Breast",
    "coconut oil": "Coconut Oil",
    "navel orange": "Navel Orange",
    "broccoli": "Broccoli",
    "olive oil": "Olive Oil",
    "sesame seeds": "Sesame Seeds",
    "sweet potatoes": "Sweet Potatoes",
    "chickpeas": "Chickpeas",
    "walnuts": "Walnuts",
    "coriander": "Coriander",
    "mint": "Mint",
    "basil": "Basil",
    "thyme": "Thyme",
    "green beans": "Green Beans",
    "cauliflower": "Cauliflower",
    "okra": "Lady's Finger (Okra)",
    "celery": "Celery",
    "flax seeds": "Flax Seeds",
}

DISH_VISUALS = {
    "flax-seed-water": {"icon": "💧", "avatar": "FS"},
    "celery-juice": {"icon": "🥬", "avatar": "CJ"},
    "mango-smoothie": {"icon": "🥭", "avatar": "MS"},
    "berry-smoothie": {"icon": "🫐", "avatar": "BS"},
    "orange-chicken": {"icon": "🍊", "avatar": "OC"},
    "sweet-potatoes": {"icon": "🍠", "avatar": "SP"},
    "beans": {"icon": "🫛", "avatar": "GB"},
    "cauliflower": {"icon": "🥦", "avatar": "CF"},
    "ladys-finger": {"icon": "✨", "avatar": "OK"},
}

RECIPE_DETAILS = {
    "flax-seed-water": {
        "title": "Flax Seed Water",
        "summary": "A simple soaked drink that feels light but grounding first thing in the morning.",
        "method": "Soak and Stir",
        "prepMinutes": 2,
        "cookMinutes": 0,
        "ingredients": [
            {"name": "Flax Seeds", "amount": "1 tbsp"},
            {"name": "Water", "amount": "12 to 14 oz"},
        ],
        "steps": [
            "Stir 1 tablespoon flax seeds into a large glass of water.",
            "Let it soak overnight or at least 6 hours so the water softens and thickens slightly.",
            "Stir once more before drinking. Sip slowly while the seeds are suspended.",
        ],
        "tip": "Chilled water makes this easier to drink, and the brief soak softens the texture noticeably.",
    },
    "celery-juice": {
        "title": "Celery Juice",
        "summary": "Clean, fresh, and mild, with the option to keep the fiber in for a fuller drink.",
        "method": "Blend or Juice",
        "prepMinutes": 5,
        "cookMinutes": 0,
        "ingredients": [
            {"name": "Celery", "amount": "4 to 5 stalks"},
            {"name": "Water", "amount": "2 to 3 tbsp if blending"},
        ],
        "steps": [
            "Roughly chop celery and add it to a blender with a small splash of water.",
            "Blend until smooth, adding just enough water to help it move.",
            "Drink as-is for more fiber, or strain for a lighter juice texture.",
        ],
        "tip": "Use only a little water so the celery flavor stays fresh instead of diluted.",
    },
    "mango-smoothie": {
        "title": "Mango Smoothie",
        "summary": "Thick, naturally sweet, and creamy from banana without needing anything extra.",
        "method": "Blender",
        "prepMinutes": 5,
        "cookMinutes": 0,
        "ingredients": [
            {"name": "Mango", "amount": "1 heaping cup"},
            {"name": "Banana", "amount": "1 medium"},
            {"name": "Coconut Water", "amount": "3/4 to 1 cup"},
        ],
        "steps": [
            "Add chopped mango, banana, and cold coconut water to a blender.",
            "Blend until silky and thick, scraping down the sides if needed.",
            "Pour immediately for the freshest taste and creamiest texture.",
        ],
        "tip": "If the fruit is cold, the smoothie tastes richer without needing ice.",
    },
    "berry-smoothie": {
        "title": "Berry Smoothie",
        "summary": "Bright and juicy, with dragon fruit keeping it light and blueberries adding depth.",
        "method": "Blender",
        "prepMinutes": 5,
        "cookMinutes": 0,
        "ingredients": [
            {"name": "Dragon Fruit", "amount": "1 cup"},
            {"name": "Blueberries", "amount": "3/4 cup"},
            {"name": "Banana", "amount": "1 medium"},
            {"name": "Coconut Water", "amount": "3/4 cup"},
        ],
        "steps": [
            "Add dragon fruit, blueberries, banana, and coconut water to a blender.",
            "Blend until the berries disappear and the color turns evenly vibrant.",
            "Taste for thickness and add a small splash of coconut water only if needed.",
        ],
        "tip": "Blend the blueberries fully so the drink stays smooth instead of slightly seedy.",
    },
    "orange-chicken": {
        "title": "Orange Chicken",
        "summary": "Citrusy and savory with caramelized edges on the chicken and crisp-tender broccoli.",
        "method": "Pan and Oven",
        "prepMinutes": 12,
        "cookMinutes": 20,
        "ingredients": [
            {"name": "Chicken Breast", "amount": "6 to 8 oz"},
            {"name": "Navel Orange", "amount": "1 whole"},
            {"name": "Broccoli", "amount": "1 1/2 to 2 cups florets"},
            {"name": "Coconut Oil", "amount": "1 tsp"},
            {"name": "Olive Oil", "amount": "1 tsp"},
            {"name": "Sesame Seeds", "amount": "1 tsp"},
        ],
        "steps": [
            "Season chicken breast with salt and pepper, then sear it in a hot pan with coconut oil until lightly golden.",
            "Transfer the chicken to the oven to finish cooking while you toss broccoli with olive oil, salt, and pepper and roast it until the edges darken.",
            "Return the pan to medium heat, squeeze in the navel orange juice, and let it reduce to a glossy citrus coating.",
            "Slice the chicken, toss it in the reduced orange glaze, and plate with the roasted broccoli.",
            "Finish with sesame seeds for crunch.",
        ],
        "tip": "Roasting the broccoli separately keeps it crisp instead of watery, which makes the orange glaze taste sharper and cleaner.",
    },
    "sweet-potatoes": {
        "title": "Sweet Potatoes",
        "summary": "Roasted sweet potatoes layered with warm chickpeas, toasted walnuts, and a fresh herb finish.",
        "method": "Oven and Pan",
        "prepMinutes": 15,
        "cookMinutes": 30,
        "ingredients": [
            {"name": "Sweet Potatoes", "amount": "1 large or 2 small"},
            {"name": "Chickpeas", "amount": "3/4 cup"},
            {"name": "Walnuts", "amount": "2 tbsp"},
            {"name": "Coriander", "amount": "1 tbsp chopped"},
            {"name": "Mint", "amount": "1 tbsp chopped"},
            {"name": "Basil", "amount": "1 tbsp chopped"},
            {"name": "Thyme", "amount": "1 tsp leaves"},
            {"name": "Olive Oil", "amount": "2 tsp"},
        ],
        "steps": [
            "Roast halved or cubed sweet potatoes with olive oil, salt, pepper, and thyme until deeply tender and browned at the edges.",
            "Warm chickpeas in a pan with a little olive oil, salt, and pepper just until they start to dry and lightly crisp.",
            "Toast walnuts briefly in a dry pan so they smell nutty and become more fragrant.",
            "Roughly chop coriander, mint, and basil together to make a bright herb finish.",
            "Pile the chickpeas and walnuts over the roasted sweet potatoes and scatter the fresh herbs over everything just before serving.",
        ],
        "tip": "Keep the herbs off the heat until the end so the bowl tastes fresh instead of muted.",
    },
    "beans": {
        "title": "Beans",
        "summary": "A fast pan-cooked green bean side with blistered spots and a clean savory finish.",
        "method": "Stovetop Pan",
        "prepMinutes": 5,
        "cookMinutes": 10,
        "ingredients": [
            {"name": "Green Beans", "amount": "1 1/2 cups"},
            {"name": "Olive Oil", "amount": "1 tsp"},
        ],
        "steps": [
            "Heat a wide pan with olive oil until hot.",
            "Add green beans in one layer and season with salt and pepper.",
            "Cook, tossing only occasionally, until the beans blister in places but stay slightly crisp.",
            "Serve right away while the beans still have a little snap.",
        ],
        "tip": "Leaving the beans mostly undisturbed at first gives better browning and more flavor.",
    },
    "cauliflower": {
        "title": "Cauliflower",
        "summary": "Golden roasted cauliflower that turns sweet and nutty without needing extra seasoning.",
        "method": "Oven or Air Fryer",
        "prepMinutes": 8,
        "cookMinutes": 18,
        "ingredients": [
            {"name": "Cauliflower", "amount": "2 cups florets"},
            {"name": "Olive Oil", "amount": "1 tsp"},
        ],
        "steps": [
            "Cut cauliflower into medium florets and coat with olive oil, salt, and pepper.",
            "Roast or air fry until the edges are deeply golden and the centers are tender.",
            "Turn once halfway through so the color develops evenly.",
            "Serve hot while the edges are still crisp.",
        ],
        "tip": "Do not overcrowd the tray or basket or the cauliflower will steam instead of roast.",
    },
    "ladys-finger": {
        "title": "Lady's Finger",
        "summary": "Dry-cooked okra with lightly charred edges and none of the slimy texture people worry about.",
        "method": "Stovetop Pan",
        "prepMinutes": 8,
        "cookMinutes": 12,
        "ingredients": [
            {"name": "Lady's Finger (Okra)", "amount": "1 1/2 cups"},
            {"name": "Olive Oil", "amount": "1 tsp"},
        ],
        "steps": [
            "Dry the okra thoroughly, then slice it into larger pieces so it browns instead of steams.",
            "Heat olive oil in a wide pan and add the okra in a single layer.",
            "Season with salt and pepper and cook over medium-high heat, stirring only occasionally.",
            "Keep cooking until the cut edges darken and the pieces feel dry and tender.",
        ],
        "tip": "The drier the okra is before it hits the pan, the better the final texture will be.",
    },
}

WEEKLY_QUANTITY_GUIDE = {
    "flax seeds": {"quantity": "1 cup (about 250 g)", "category": "Pantry"},
    "celery": {"quantity": "3 bunches", "category": "Produce"},
    "mango": {"quantity": "4 mangoes", "category": "Fruit"},
    "banana": {"quantity": "7 bananas", "category": "Fruit"},
    "coconut water": {"quantity": "2 to 3 liters", "category": "Beverages"},
    "dragon fruit": {"quantity": "3 dragon fruits", "category": "Fruit"},
    "blueberries": {"quantity": "3 cups", "category": "Fruit"},
    "chicken breast": {"quantity": "2 lb", "category": "Protein"},
    "navel orange": {"quantity": "4 oranges", "category": "Fruit"},
    "broccoli": {"quantity": "4 medium heads", "category": "Produce"},
    "sesame seeds": {"quantity": "1 small packet or jar", "category": "Pantry"},
    "sweet potatoes": {"quantity": "6 medium sweet potatoes", "category": "Produce"},
    "chickpeas": {"quantity": "2 cans or 3 cups cooked", "category": "Pantry"},
    "walnuts": {"quantity": "1 1/2 cups", "category": "Pantry"},
    "coriander": {"quantity": "1 bunch", "category": "Herbs"},
    "mint": {"quantity": "1 bunch", "category": "Herbs"},
    "basil": {"quantity": "1 bunch", "category": "Herbs"},
    "thyme": {"quantity": "1 small bunch", "category": "Herbs"},
    "green beans": {"quantity": "1 1/2 lb", "category": "Produce"},
    "cauliflower": {"quantity": "2 medium heads", "category": "Produce"},
    "okra": {"quantity": "1 1/2 lb", "category": "Produce"},
    "olive oil": {"quantity": "1 bottle", "category": "Pantry"},
    "coconut oil": {"quantity": "1 small jar", "category": "Pantry"},
}

DEFAULT_STORE = "Whole Foods / Fred Meyer / QFC"


def clean(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\n", " ").split()).strip()


def slugify(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return slug or "item"


def dish_id_from_label(label: str) -> str:
    recipe_key = canonicalize(label)
    recipe_slug_map = {
        "flax seeds": "flax-seed-water",
        "celery": "celery-juice",
        "green beans": "beans",
        "okra": "ladys-finger",
    }
    return recipe_slug_map.get(recipe_key, slugify(label))


def canonicalize(name: str) -> str:
    normalized = re.sub(r"[^a-z0-9' ]+", "", name.lower()).strip()
    return CANONICAL_OVERRIDES.get(normalized, normalized)


def display_name(name: str) -> str:
    if name in DISPLAY_NAMES:
        return DISPLAY_NAMES[name]
    return " ".join(part.capitalize() for part in name.split())


def load_sheet():
    workbook = load_workbook(WORKBOOK_PATH, data_only=True)
    return workbook["Sheet1"]


def parse_meal_plan(ws) -> list[dict]:
    meal_rows = [
        ("Morning Cleanse", 7, 8),
        ("Breakfast", 10),
        ("Lunch", 12),
        ("Dinner", 14),
    ]
    days = []
    for index, col in enumerate(range(4, 11)):
        meals = []
        for row_info in meal_rows:
            name = row_info[0]
            rows = row_info[1:]
            items = [clean(ws.cell(row, col).value) for row in rows if clean(ws.cell(row, col).value)]
            meals.append(
                {
                    "mealType": name,
                    "dishIds": [dish_id_from_label(item) for item in items],
                    "labels": items,
                }
            )
        days.append(
            {
                "id": DAY_NAMES[index].lower(),
                "name": DAY_NAMES[index],
                "meals": meals,
            }
        )
    return days


def parse_workbook_recipe_blocks(ws) -> dict[str, list[str]]:
    recipes = {}
    for col in range(13, ws.max_column + 1):
        row = 1
        while row <= ws.max_row:
            current = clean(ws.cell(row, col).value)
            previous = clean(ws.cell(row - 1, col).value) if row > 1 else ""
            next_value = clean(ws.cell(row + 1, col).value) if row < ws.max_row else ""
            if current and not previous and next_value:
                items = []
                cursor = row + 1
                while cursor <= ws.max_row:
                    ingredient = clean(ws.cell(cursor, col).value)
                    if not ingredient:
                        break
                    items.append(ingredient)
                    cursor += 1
                recipes[current] = items
                row = cursor
            else:
                row += 1
    return recipes


def parse_store_lists(ws) -> list[dict]:
    stores = []
    for col in range(4, ws.max_column + 1):
        store = clean(ws.cell(23, col).value)
        if not store:
            continue
        items = []
        row = 24
        while row <= ws.max_row:
            item = clean(ws.cell(row, col).value)
            if not item:
                break
            canonical = canonicalize(item)
            items.append(
                {
                    "name": display_name(canonical),
                    "canonical": canonical,
                    "source": item,
                }
            )
            row += 1
        stores.append({"store": store, "items": items})
    return stores


def build_recipes(workbook_recipes: dict[str, list[str]]) -> tuple[list[dict], dict[str, dict]]:
    derived = []
    recipe_by_id = {}
    for recipe_id, details in RECIPE_DETAILS.items():
        workbook_ingredients = workbook_recipes.get(details["title"], [])
        listed_ingredients = []
        for ingredient in details["ingredients"]:
            canonical = canonicalize(ingredient["name"])
            listed_ingredients.append(
                {
                    "name": display_name(canonical),
                    "canonical": canonical,
                    "amount": ingredient["amount"],
                }
            )
        recipe = {
            "id": recipe_id,
            "title": details["title"],
            "summary": details["summary"],
            "method": details["method"],
            "icon": DISH_VISUALS.get(recipe_id, {}).get("icon", "🍽️"),
            "avatar": DISH_VISUALS.get(recipe_id, {}).get("avatar", "PL"),
            "prepMinutes": details["prepMinutes"],
            "cookMinutes": details["cookMinutes"],
            "ingredients": listed_ingredients,
            "steps": details["steps"],
            "tip": details["tip"],
            "workbookIngredients": workbook_ingredients,
        }
        derived.append(recipe)
        recipe_by_id[recipe_id] = recipe
    return derived, recipe_by_id


def dish_recipe_id(dish_label: str) -> str:
    return dish_id_from_label(dish_label)


def build_dishes(days: list[dict], recipe_by_id: dict[str, dict]) -> list[dict]:
    seen = {}
    for day in days:
        for meal in day["meals"]:
            for label in meal["labels"]:
                dish_id = slugify(label)
                dish_id = dish_id_from_label(label)
                if dish_id in seen:
                    continue
                recipe = recipe_by_id[dish_recipe_id(label)]
                seen[dish_id] = {
                    "id": dish_id,
                    "title": label,
                    "recipeId": recipe["id"],
                    "summary": recipe["summary"],
                    "icon": recipe["icon"],
                    "avatar": recipe["avatar"],
                }
    return list(seen.values())


def build_store_lookup(store_lists: list[dict]) -> dict[str, list[str]]:
    lookup = {}
    for store in store_lists:
        for item in store["items"]:
            lookup.setdefault(item["canonical"], []).append(store["store"])
    return lookup


def build_weekly_supply_list(recipe_by_id: dict[str, dict], store_lookup: dict[str, list[str]]) -> list[dict]:
    ingredient_keys = []
    for recipe in recipe_by_id.values():
        for ingredient in recipe["ingredients"]:
            canonical = ingredient["canonical"]
            if canonical in {"water"}:
                continue
            if canonical not in ingredient_keys:
                ingredient_keys.append(canonical)

    grouped = {}
    for canonical in ingredient_keys:
        guide = WEEKLY_QUANTITY_GUIDE.get(canonical)
        if not guide:
            continue
        store = store_lookup.get(canonical, [DEFAULT_STORE])[0]
        grouped.setdefault(guide["category"], []).append(
            {
                "name": display_name(canonical),
                "quantity": guide["quantity"],
                "store": store,
            }
        )

    return [
        {
            "category": category,
            "items": sorted(items, key=lambda item: item["name"].lower()),
        }
        for category, items in sorted(grouped.items(), key=lambda pair: pair[0].lower())
    ]


def build_daily_supply_list(days: list[dict], dishes: list[dict], recipe_by_id: dict[str, dict], store_lookup: dict[str, list[str]]) -> dict[str, list[dict]]:
    dish_lookup = {dish["id"]: dish for dish in dishes}
    output = {}
    for day in days:
        grouped = {}
        seen = set()
        for meal in day["meals"]:
            for dish_id in meal["dishIds"]:
                recipe = recipe_by_id[dish_lookup[dish_id]["recipeId"]]
                for ingredient in recipe["ingredients"]:
                    key = (ingredient["canonical"], ingredient["amount"])
                    if key in seen or ingredient["canonical"] == "water":
                        continue
                    seen.add(key)
                    guide = WEEKLY_QUANTITY_GUIDE.get(ingredient["canonical"])
                    category = guide["category"] if guide else "Needs Review"
                    store = store_lookup.get(ingredient["canonical"], [DEFAULT_STORE])[0]
                    grouped.setdefault(category, []).append(
                        {
                            "name": ingredient["name"],
                            "used": ingredient["amount"],
                            "buy": guide["quantity"] if guide else "See weekly supply list",
                            "store": store,
                        }
                    )
        output[day["id"]] = [
            {"category": category, "items": sorted(items, key=lambda item: item["name"].lower())}
            for category, items in sorted(grouped.items(), key=lambda pair: pair[0].lower())
        ]
    return output


def build_day_shopping(days: list[dict], recipe_by_id: dict[str, dict], store_lookup: dict[str, list[str]]) -> dict[str, dict]:
    result = {}
    all_unassigned = set()
    for day in days:
        store_map = {}
        unassigned = set()
        for meal in day["meals"]:
            for dish_id in meal["dishIds"]:
                recipe = recipe_by_id[dish_id]
                for ingredient in recipe["ingredients"]:
                    canonical = ingredient["canonical"]
                    if canonical in {"water", "salt", "black pepper", "neutral oil", "olive oil"}:
                        continue
                    stores = store_lookup.get(canonical, [])
                    if stores:
                        for store in stores:
                            store_map.setdefault(store, set()).add(ingredient["name"])
                    else:
                        unassigned.add(ingredient["name"])
                        all_unassigned.add(ingredient["name"])
        result[day["id"]] = {
            "stores": [
                {"store": store, "items": sorted(items)}
                for store, items in sorted(store_map.items(), key=lambda pair: pair[0].lower())
            ],
            "unassigned": sorted(unassigned),
        }
    result["weekly"] = {"unassigned": sorted(all_unassigned)}
    return result


def build_weekly_shopping(store_lists: list[dict], all_unassigned: list[str]) -> list[dict]:
    weekly = []
    for store in store_lists:
        weekly.append(
            {
                "store": store["store"],
                "items": [item["name"] for item in store["items"]],
            }
        )
    if all_unassigned:
        weekly.append({"store": "Needs Store Assignment", "items": all_unassigned})
    return weekly


def build_data() -> dict:
    ws = load_sheet()
    days = parse_meal_plan(ws)
    workbook_recipes = parse_workbook_recipe_blocks(ws)
    store_lists = parse_store_lists(ws)
    recipes, recipe_by_id = build_recipes(workbook_recipes)
    dishes = build_dishes(days, recipe_by_id)
    store_lookup = build_store_lookup(store_lists)
    day_shopping = build_day_shopping(days, recipe_by_id, store_lookup)
    weekly_shopping = build_weekly_shopping(store_lists, day_shopping["weekly"]["unassigned"])
    weekly_supply = build_weekly_supply_list(recipe_by_id, store_lookup)
    daily_supply = build_daily_supply_list(days, dishes, recipe_by_id, store_lookup)

    return {
        "title": "Weekly Diet Dashboard",
        "subtitle": "Tap a day to see the meals, recipes, and shopping details that matter right then.",
        "generatedAt": datetime.now().strftime("%B %d, %Y"),
        "defaultDay": "monday",
        "pantryBasics": PANTRY_BASICS,
        "days": days,
        "dishes": dishes,
        "recipes": recipes,
        "dayShopping": day_shopping,
        "dailySupply": daily_supply,
        "weeklySupply": weekly_supply,
        "weeklyShopping": weekly_shopping,
        "pdfPath": "output/weekly-diet-dashboard.pdf",
    }


def build_html(data: dict) -> str:
    title = escape(data["title"])
    subtitle = escape(data["subtitle"])
    embedded_json = json.dumps(data, indent=2)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{title}</title>
  <meta name="description" content="{escape(data['subtitle'])}" />
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700;800&family=Space+Grotesk:wght@500;700&display=swap" rel="stylesheet">
  <style>
    :root {{
      --bg-1: #fff0de;
      --bg-2: #ffe4f2;
      --bg-3: #e5fff1;
      --panel: rgba(255, 250, 247, 0.88);
      --panel-strong: rgba(255, 255, 255, 0.92);
      --ink: #182329;
      --muted: #56616a;
      --accent: #ff5c8a;
      --accent-strong: #ff3f74;
      --accent-soft: #ffe3ee;
      --mint: #7ce7b8;
      --sky: #c9f4ff;
      --warm: #ff9b42;
      --line: rgba(24, 35, 41, 0.08);
      --shadow: 0 24px 48px rgba(100, 67, 95, 0.12);
      --radius-xl: 28px;
      --radius-lg: 22px;
      --radius-md: 16px;
      --max: 1120px;
    }}
    * {{ box-sizing: border-box; }}
    html {{ scroll-behavior: smooth; }}
    body {{
      margin: 0;
      color: var(--ink);
      font-family: "DM Sans", "Segoe UI", sans-serif;
      background:
        radial-gradient(circle at top left, rgba(255, 92, 138, 0.18), transparent 22%),
        radial-gradient(circle at top right, rgba(124, 231, 184, 0.24), transparent 25%),
        radial-gradient(circle at bottom left, rgba(201, 244, 255, 0.24), transparent 20%),
        linear-gradient(180deg, var(--bg-1) 0%, #fff7ef 42%, var(--bg-2) 78%, var(--bg-3) 100%);
      min-height: 100vh;
    }}
    .shell {{
      width: min(calc(100% - 20px), var(--max));
      margin: 0 auto;
      padding: 18px 0 40px;
    }}
    .hero {{
      position: relative;
      overflow: hidden;
      border-radius: 32px;
      padding: 24px;
      background:
        linear-gradient(135deg, rgba(255, 255, 255, 0.95), rgba(255, 244, 249, 0.82)),
        radial-gradient(circle at top right, rgba(124, 231, 184, 0.24), transparent 35%);
      border: 1px solid rgba(255, 255, 255, 0.75);
      box-shadow: var(--shadow);
    }}
    .hero::after {{
      content: "";
      position: absolute;
      width: 260px;
      height: 260px;
      right: -60px;
      bottom: -120px;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(255, 92, 138, 0.18), transparent 68%);
      pointer-events: none;
    }}
    @keyframes floaty {{
      0%, 100% {{ transform: translateY(0px); }}
      50% {{ transform: translateY(-4px); }}
    }}
    @keyframes fadeRise {{
      from {{ opacity: 0; transform: translateY(12px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    .eyebrow {{
      display: inline-flex;
      gap: 8px;
      align-items: center;
      padding: 7px 12px;
      border-radius: 999px;
      background: rgba(255, 92, 138, 0.1);
      color: var(--accent-strong);
      font-size: 12px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
      font-weight: 700;
    }}
    h1 {{
      margin: 14px 0 10px;
      font-family: "Space Grotesk", "DM Sans", sans-serif;
      font-size: clamp(30px, 7vw, 60px);
      line-height: 0.95;
      max-width: 9ch;
      letter-spacing: -0.04em;
    }}
    .subtitle {{
      max-width: 52ch;
      margin: 0;
      color: var(--muted);
      font-size: 15px;
      line-height: 1.55;
    }}
    .hero-meta {{
      margin-top: 18px;
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
    }}
    .meta-pill, .method-pill, .store-pill, .pantry-pill {{
      display: inline-flex;
      align-items: center;
      padding: 8px 12px;
      border-radius: 999px;
      border: 1px solid rgba(32, 49, 39, 0.08);
      background: rgba(255, 255, 255, 0.78);
      font-size: 12px;
      color: var(--muted);
    }}
    .top-actions {{
      margin-top: 16px;
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
    }}
    .action-link {{
      border: 0;
      text-decoration: none;
      background: linear-gradient(135deg, var(--accent), var(--accent-strong));
      color: white;
      border-radius: 999px;
      padding: 12px 16px;
      font-weight: 700;
      box-shadow: 0 12px 26px rgba(255, 92, 138, 0.24);
      cursor: pointer;
      font: inherit;
    }}
    .action-link.secondary {{
      background: rgba(255, 255, 255, 0.86);
      color: var(--ink);
      box-shadow: none;
    }}
    .day-bar {{
      position: sticky;
      top: 0;
      z-index: 5;
      margin: 18px 0 22px;
      padding: 10px 0 4px;
      backdrop-filter: blur(16px);
    }}
    .day-track {{
      display: flex;
      gap: 10px;
      overflow-x: auto;
      padding-bottom: 6px;
      scrollbar-width: none;
    }}
    .day-track::-webkit-scrollbar {{ display: none; }}
    .day-chip {{
      flex: 0 0 auto;
      border: 0;
      border-radius: 999px;
      padding: 13px 16px;
      background: rgba(255, 255, 255, 0.76);
      color: var(--ink);
      font-weight: 800;
      font-size: 13px;
      box-shadow: 0 10px 20px rgba(100, 67, 95, 0.08);
      cursor: pointer;
    }}
    .day-chip.active {{
      background: linear-gradient(135deg, var(--accent), var(--accent-strong));
      color: white;
      box-shadow: 0 14px 26px rgba(255, 92, 138, 0.28);
    }}
    .focus-layout, .logistics-layout {{
      display: grid;
      gap: 18px;
    }}
    .recipe-stage, .pantry-stage {{
      min-width: 0;
    }}
    .section-heading {{
      display: flex;
      justify-content: space-between;
      align-items: end;
      gap: 16px;
      margin: 22px 0 14px;
    }}
    .section-heading.compact {{
      margin-top: 0;
    }}
    .section-heading h2 {{
      margin: 0;
      font-family: "Space Grotesk", "DM Sans", sans-serif;
      font-size: 24px;
      letter-spacing: -0.03em;
    }}
    .section-heading p {{
      margin: 6px 0 0;
      color: var(--muted);
      line-height: 1.45;
      max-width: 58ch;
    }}
    .selected-day-shell, .panel {{
      background: var(--panel);
      border: 1px solid rgba(255, 255, 255, 0.7);
      box-shadow: var(--shadow);
      border-radius: var(--radius-xl);
    }}
    .recipe-stage .panel, .selected-day-shell, .shopping-card, .pantry-stage .panel {{
      animation: fadeRise 240ms ease;
    }}
    .selected-day-shell {{
      padding: 18px;
    }}
    .selected-day-head {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: end;
      margin-bottom: 16px;
    }}
    .selected-day-head h3 {{
      margin: 0;
      font-family: "Space Grotesk", "DM Sans", sans-serif;
      font-size: clamp(26px, 5vw, 40px);
      letter-spacing: -0.04em;
    }}
    .selected-day-head small {{
      display: block;
      margin-bottom: 6px;
      color: var(--warm);
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.08em;
    }}
    .meals-grid, .recipes-grid, .shopping-grid {{
      display: grid;
      gap: 16px;
    }}
    .meals-grid {{
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }}
    .meal-card, .panel {{
      padding: 16px;
    }}
    .meal-card {{
      border-radius: var(--radius-lg);
      background: rgba(255, 255, 255, 0.7);
      border: 1px solid var(--line);
    }}
    .meal-card h4 {{
      margin: 0 0 10px;
      color: var(--muted);
      font-size: 13px;
      letter-spacing: 0.08em;
      text-transform: uppercase;
    }}
    .meal-list {{
      display: grid;
      gap: 10px;
    }}
    .meal-topline {{
      display: flex;
      gap: 10px;
      align-items: center;
    }}
    .dish-avatar {{
      width: 40px;
      height: 40px;
      border-radius: 14px;
      background: linear-gradient(135deg, rgba(255, 92, 138, 0.16), rgba(124, 231, 184, 0.28));
      display: grid;
      place-items: center;
      font-weight: 800;
      font-size: 12px;
      color: var(--ink);
      flex: 0 0 auto;
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.7);
      animation: floaty 3.2s ease-in-out infinite;
    }}
    .dish-icon {{
      font-size: 20px;
      line-height: 1;
    }}
    .meal-item {{
      display: block;
      text-decoration: none;
      color: inherit;
      width: 100%;
      text-align: left;
      padding: 12px;
      border-radius: 14px;
      background: var(--panel-strong);
      border: 1px solid rgba(32, 49, 39, 0.06);
      transition: transform 120ms ease, box-shadow 120ms ease, border-color 120ms ease;
    }}
    .meal-item:hover, .meal-item:focus {{
      transform: translateY(-1px);
      box-shadow: 0 10px 18px rgba(32, 49, 39, 0.08);
      border-color: rgba(31, 106, 82, 0.28);
      outline: none;
    }}
    .meal-item strong {{
      display: block;
      font-size: 15px;
      margin-bottom: 2px;
    }}
    .meal-item span {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.4;
    }}
    .recipes-grid, .shopping-grid {{
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }}
    .panel h3 {{
      margin: 0 0 10px;
      font-family: "Space Grotesk", "DM Sans", sans-serif;
      font-size: 20px;
      letter-spacing: -0.03em;
    }}
    .recipe-card.flash {{
      outline: 3px solid rgba(31, 106, 82, 0.3);
      box-shadow: 0 0 0 6px rgba(31, 106, 82, 0.08);
    }}
    .panel p {{
      margin: 0;
      color: var(--muted);
      line-height: 1.5;
    }}
    .panel-top {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: start;
      margin-bottom: 12px;
    }}
    .recipe-title-group {{
      display: flex;
      gap: 12px;
      align-items: flex-start;
    }}
    .recipe-avatar {{
      width: 48px;
      height: 48px;
      border-radius: 18px;
      background: linear-gradient(135deg, rgba(255, 92, 138, 0.16), rgba(124, 231, 184, 0.28));
      display: grid;
      place-items: center;
      flex: 0 0 auto;
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.76);
      animation: floaty 3.6s ease-in-out infinite;
    }}
    .recipe-avatar .dish-icon {{
      font-size: 24px;
    }}
    .recipe-meta {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin: 12px 0;
    }}
    .ingredient-list, .steps-list, .shopping-list, .pantry-list {{
      margin: 0;
      padding: 0;
      list-style: none;
    }}
    .ingredient-list {{
      display: grid;
      gap: 8px;
      margin-bottom: 14px;
    }}
    .ingredient-list li {{
      background: linear-gradient(135deg, rgba(255, 227, 238, 0.9), rgba(201, 244, 255, 0.85));
      border-radius: 16px;
      padding: 10px 12px;
      font-size: 13px;
      border: 1px solid rgba(24, 35, 41, 0.06);
    }}
    .ingredient-list li strong {{
      display: block;
      font-size: 14px;
      margin-bottom: 3px;
    }}
    .steps-list {{
      display: grid;
      gap: 10px;
      counter-reset: recipe-step;
      margin-bottom: 12px;
    }}
    .steps-list li {{
      position: relative;
      padding-left: 38px;
      line-height: 1.45;
      color: var(--ink);
    }}
    .steps-list li::before {{
      counter-increment: recipe-step;
      content: counter(recipe-step);
      position: absolute;
      left: 0;
      top: 0;
      width: 26px;
      height: 26px;
      border-radius: 50%;
      background: rgba(31, 106, 82, 0.12);
      color: var(--accent-strong);
      font-size: 13px;
      font-weight: 800;
      display: grid;
      place-items: center;
    }}
    .tip {{
      margin-top: 10px;
      padding: 12px;
      border-radius: 14px;
      background: linear-gradient(135deg, rgba(255, 155, 66, 0.12), rgba(255, 92, 138, 0.08));
      color: #8a4c1a;
      font-size: 14px;
      line-height: 1.45;
    }}
    .section-sticker {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 6px 10px;
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.72);
      font-size: 12px;
      font-weight: 700;
      color: var(--muted);
      margin-bottom: 8px;
    }}
    .shopping-card h3 {{
      margin-bottom: 8px;
    }}
    .shopping-toggle-row {{
      margin: 0 0 14px;
    }}
    .toggle-button {{
      border: 0;
      border-radius: 999px;
      padding: 12px 16px;
      font: inherit;
      font-weight: 700;
      cursor: pointer;
      background: linear-gradient(135deg, var(--accent), var(--accent-strong));
      color: white;
      box-shadow: 0 10px 20px rgba(255, 92, 138, 0.2);
    }}
    .toggle-button.subtle {{
      background: rgba(255,255,255,0.82);
      color: var(--ink);
      box-shadow: none;
      border: 1px solid var(--line);
    }}
    .collapsible-section[hidden] {{
      display: none !important;
    }}
    .collapsible-section {{
      margin-bottom: 10px;
    }}
    .shopping-list {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 10px;
    }}
    .shopping-list li {{
      padding: 8px 12px;
      background: rgba(255, 255, 255, 0.8);
      border: 1px solid rgba(32, 49, 39, 0.08);
      border-radius: 999px;
      font-size: 14px;
    }}
    .shopping-note {{
      margin-top: 12px;
      color: var(--muted);
      line-height: 1.45;
    }}
    .weekly-supply-grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 16px;
    }}
    .supply-list {{
      list-style: none;
      margin: 12px 0 0;
      padding: 0;
      display: grid;
      gap: 10px;
    }}
    .supply-list li {{
      padding: 12px;
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.78);
      border: 1px solid rgba(32, 49, 39, 0.08);
    }}
    .supply-list strong {{
      display: block;
      margin-bottom: 4px;
    }}
    .supply-list span {{
      display: block;
      color: var(--muted);
      line-height: 1.4;
      font-size: 13px;
    }}
    .buy-quantity {{
      color: var(--accent-strong);
      font-weight: 800;
      font-size: 15px;
      margin: 3px 0;
    }}
    .used-quantity {{
      color: var(--ink);
      font-weight: 700;
      font-size: 14px;
      margin: 3px 0;
    }}
    details.weekly {{
      margin-top: 18px;
      background: rgba(255, 255, 255, 0.52);
      border: 1px solid rgba(32, 49, 39, 0.08);
      border-radius: 20px;
      overflow: hidden;
    }}
    details.weekly summary {{
      cursor: pointer;
      list-style: none;
      padding: 16px 18px;
      font-weight: 800;
    }}
    details.weekly summary::-webkit-details-marker {{ display: none; }}
    .weekly-body {{
      padding: 0 18px 18px;
    }}
    .pantry-list {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 12px;
    }}
    .footer-note {{
      margin-top: 26px;
      color: var(--muted);
      text-align: center;
      font-size: 13px;
    }}
    @media (max-width: 880px) {{
      .meals-grid, .recipes-grid, .shopping-grid, .weekly-supply-grid {{
        grid-template-columns: 1fr;
      }}
      .selected-day-head {{
        flex-direction: column;
        align-items: start;
      }}
    }}
    @media (min-width: 881px) {{
      .focus-layout {{
        grid-template-columns: minmax(0, 1.08fr) minmax(360px, 0.92fr);
        align-items: start;
      }}
      .recipe-stage {{
        position: sticky;
        top: 88px;
      }}
      .logistics-layout {{
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }}
      .pantry-stage {{
        grid-column: 1 / -1;
      }}
    }}
    @media (min-width: 1240px) {{
      .shell {{
        width: min(calc(100% - 36px), 1320px);
      }}
      .logistics-layout {{
        grid-template-columns: 1fr 1fr 0.78fr;
        align-items: start;
      }}
      .pantry-stage {{
        grid-column: auto;
      }}
    }}
    @media print {{
      @page {{
        size: A4;
        margin: 10mm;
      }}
      body {{
        background: white;
      }}
      .shell {{
        width: 100%;
        padding: 0;
      }}
      .hero, .selected-day-shell, .panel {{
        box-shadow: none;
        background: white;
      }}
      .day-bar {{
        position: static;
        backdrop-filter: none;
      }}
      .meals-grid, .recipes-grid, .shopping-grid, .weekly-supply-grid {{
        grid-template-columns: 1fr;
      }}
      .action-link {{
        color: var(--ink);
        border: 1px solid var(--line);
        background: white;
        box-shadow: none;
      }}
    }}
  </style>
</head>
<body>
  <main class="shell">
    <section class="hero">
      <span class="eyebrow">Public Share Link Ready</span>
      <h1>{title}</h1>
      <p class="subtitle">{subtitle}</p>
      <div class="top-actions">
        <a class="action-link" href="#day-view">Jump to Day View</a>
        <button class="action-link secondary" type="button" id="open-day-shopping">Day Shopping</button>
        <button class="action-link secondary" type="button" id="open-weekly-shopping">Weekly Shopping</button>
        <a class="action-link secondary" href="{escape(data['pdfPath'])}">Open PDF Backup</a>
      </div>
    </section>

    <nav class="day-bar" aria-label="Pick a day">
      <div class="day-track" id="day-track"></div>
    </nav>

    <section class="focus-layout">
      <section id="day-view" class="selected-day-shell">
        <div class="selected-day-head">
          <div>
            <small>Selected Day</small>
            <h3 id="selected-day-title">Monday</h3>
          </div>
          <p id="selected-day-copy" class="subtitle">Meals, recipes, and shopping are focused on the day you choose.</p>
        </div>
        <div class="meals-grid" id="meals-grid"></div>
      </section>

      <section class="recipe-stage">
        <div class="section-heading compact">
          <div>
            <div class="section-sticker">✨ Tap a meal to open its recipe</div>
            <h2>Recipe Focus</h2>
            <p>Only one recipe opens at a time so the page stays calm and easy to scan.</p>
          </div>
        </div>
        <div id="recipe-empty" class="panel">
          <h3>Choose a meal card above</h3>
          <p>Tap any breakfast, lunch, dinner, or cleanse item to open the matching recipe with the exact amount used.</p>
        </div>
        <div class="recipes-grid" id="recipes-grid"></div>
      </section>
    </section>

    <section class="logistics-layout">
      <section>
        <div class="section-heading">
          <div>
            <div class="section-sticker">🛒 One-person weekly buy guide</div>
            <h2>Weekly Supply Buy List</h2>
            <p>These quantities are estimated for one person following this exact weekly plan.</p>
          </div>
        </div>
        <div class="shopping-toggle-row">
          <button class="toggle-button" type="button" id="toggle-weekly-supply">Open Weekly Supply List</button>
        </div>
        <div class="collapsible-section" id="weekly-supply-section" hidden>
          <div class="weekly-supply-grid" id="weekly-supply-grid"></div>
        </div>
      </section>

      <section>
        <div class="section-heading">
          <div>
            <div class="section-sticker">🧺 Filtered by selected day</div>
            <h2>Shopping For This Day</h2>
            <p>Open this only when you need to buy for the currently selected day.</p>
          </div>
        </div>
        <div class="shopping-toggle-row">
          <button class="toggle-button" type="button" id="toggle-day-shopping">Open Day Shopping</button>
        </div>
        <div class="collapsible-section" id="day-shopping-section" hidden>
          <div class="weekly-supply-grid" id="daily-supply-grid"></div>
          <div class="shopping-grid" id="shopping-grid"></div>
        </div>
        <div class="shopping-toggle-row">
          <button class="toggle-button subtle" type="button" id="toggle-weekly-shopping">Open Weekly Store List</button>
        </div>
        <div class="collapsible-section" id="weekly-shopping-section" hidden>
          <div class="shopping-grid" id="weekly-shopping-grid"></div>
        </div>
      </section>

      <section class="pantry-stage">
        <div class="section-heading">
          <div>
            <h2>Pantry Basics Used In Recipes</h2>
            <p>These are the only non-workbook basics assumed anywhere on the site.</p>
          </div>
        </div>
        <div class="panel">
          <ul class="pantry-list" id="pantry-list"></ul>
        </div>
      </section>
    </section>

    <p class="footer-note">Generated from your workbook on {escape(data['generatedAt'])}. Share the public site link with anyone who needs the plan.</p>
  </main>

  <script>
    const data = {embedded_json};
    const stateKey = "diet-plan-selected-day-v1";
    const storedDay = localStorage.getItem(stateKey);
    const dishesById = Object.fromEntries(data.dishes.map((dish) => [dish.id, dish]));
    const recipesById = Object.fromEntries(data.recipes.map((recipe) => [recipe.id, recipe]));

    const dayTrack = document.getElementById("day-track");
    const mealsGrid = document.getElementById("meals-grid");
    const recipesGrid = document.getElementById("recipes-grid");
    const shoppingGrid = document.getElementById("shopping-grid");
    const dailySupplyGrid = document.getElementById("daily-supply-grid");
    const weeklySupplyGrid = document.getElementById("weekly-supply-grid");
    const weeklyShoppingGrid = document.getElementById("weekly-shopping-grid");
    const pantryList = document.getElementById("pantry-list");
    const selectedDayTitle = document.getElementById("selected-day-title");
    const selectedDayCopy = document.getElementById("selected-day-copy");
    const recipeEmpty = document.getElementById("recipe-empty");
    const dayShoppingSection = document.getElementById("day-shopping-section");
    const weeklyShoppingSection = document.getElementById("weekly-shopping-section");
    const weeklySupplySection = document.getElementById("weekly-supply-section");
    const toggleDayShopping = document.getElementById("toggle-day-shopping");
    const toggleWeeklyShopping = document.getElementById("toggle-weekly-shopping");
    const toggleWeeklySupply = document.getElementById("toggle-weekly-supply");
    const openDayShopping = document.getElementById("open-day-shopping");
    const openWeeklyShopping = document.getElementById("open-weekly-shopping");

    let activeDayId = storedDay && data.days.some((day) => day.id === storedDay) ? storedDay : data.defaultDay;
    let activeRecipeId = null;

    function renderDayChips() {{
      dayTrack.innerHTML = "";
      data.days.forEach((day) => {{
        const button = document.createElement("button");
        button.className = "day-chip" + (day.id === activeDayId ? " active" : "");
        button.textContent = day.name.slice(0, 3);
        button.setAttribute("aria-label", day.name);
        button.addEventListener("click", () => {{
          activeDayId = day.id;
          localStorage.setItem(stateKey, activeDayId);
          render();
        }});
        dayTrack.appendChild(button);
      }});
    }}

    function renderMeals(day) {{
      mealsGrid.innerHTML = "";
      day.meals.forEach((meal) => {{
        const card = document.createElement("article");
        card.className = "meal-card";
        card.innerHTML = `<h4>${{meal.mealType}}</h4>`;
        const list = document.createElement("div");
        list.className = "meal-list";
        meal.dishIds.forEach((dishId) => {{
          const dish = dishesById[dishId];
          const recipe = recipesById[dish.recipeId];
        const item = document.createElement("button");
        item.className = "meal-item";
        item.type = "button";
        item.addEventListener("click", (event) => {{
          event.preventDefault();
            openRecipe(recipe.id);
        }});
          item.innerHTML = `
            <div class="meal-topline">
              <div class="dish-avatar">${{dish.avatar}}</div>
              <div>
                <strong>${{dish.icon}} ${{dish.title}}</strong>
                <span>${{recipe.summary}}</span>
              </div>
            </div>
          `;
          list.appendChild(item);
        }});
        card.appendChild(list);
        mealsGrid.appendChild(card);
      }});
    }}

    function renderRecipes(day) {{
      recipesGrid.innerHTML = "";
      const recipeIds = [...new Set(day.meals.flatMap((meal) => meal.dishIds.map((dishId) => dishesById[dishId].recipeId)))];
      if (!activeRecipeId || !recipeIds.includes(activeRecipeId)) {{
        recipeEmpty.hidden = false;
        return;
      }}
      const recipe = recipesById[activeRecipeId];
      recipeEmpty.hidden = true;
      const card = document.createElement("article");
      card.className = "panel recipe-card";
      card.id = `recipe-${{recipe.id}}`;
      const ingredients = recipe.ingredients.map((ingredient) => `
        <li>
          <strong>${{ingredient.name}}</strong>
          <span class="buy-quantity">Use: ${{ingredient.amount}}</span>
        </li>
      `).join("");
      const steps = recipe.steps.map((step) => `<li>${{step}}</li>`).join("");
      card.innerHTML = `
        <div class="panel-top">
          <div class="recipe-title-group">
            <div class="recipe-avatar"><span class="dish-icon">${{recipe.icon}}</span></div>
            <div>
              <h3>${{recipe.title}}</h3>
              <p>${{recipe.summary}}</p>
            </div>
          </div>
          <span class="method-pill">${{recipe.method}}</span>
        </div>
        <div class="recipe-meta">
          <span class="meta-pill">Prep ${{recipe.prepMinutes}} min</span>
          <span class="meta-pill">Cook ${{recipe.cookMinutes}} min</span>
        </div>
        <ul class="ingredient-list">${{ingredients}}</ul>
        <ol class="steps-list">${{steps}}</ol>
        <div class="tip">${{recipe.tip}}</div>
      `;
      recipesGrid.appendChild(card);
    }}

    function renderDailySupply(day) {{
      dailySupplyGrid.innerHTML = "";
      const groups = data.dailySupply[day.id] || [];
      groups.forEach((group) => {{
        const card = document.createElement("article");
        card.className = "panel shopping-card";
        card.innerHTML = `
          <span class="store-pill">${{group.category}}</span>
          <ul class="supply-list">
            ${{group.items.map((item) => `
              <li>
                <strong>${{item.name}}</strong>
                <span class="used-quantity">Use today: ${{item.used}}</span>
                <span class="buy-quantity">Buy for the week: ${{item.buy}}</span>
                <span>Best place to buy: ${{item.store}}</span>
              </li>
            `).join("")}}
          </ul>
        `;
        dailySupplyGrid.appendChild(card);
      }});
    }}

    function renderWeeklySupply() {{
      weeklySupplyGrid.innerHTML = "";
      data.weeklySupply.forEach((group) => {{
        const card = document.createElement("article");
        card.className = "panel shopping-card";
        card.innerHTML = `
          <span class="store-pill">${{group.category}}</span>
          <ul class="supply-list">
            ${{group.items.map((item) => `
              <li>
                <strong>${{item.name}}</strong>
                <span class="buy-quantity">Buy for the week: ${{item.quantity}}</span>
                <span>Best place to buy: ${{item.store}}</span>
              </li>
            `).join("")}}
          </ul>
        `;
        weeklySupplyGrid.appendChild(card);
      }});
    }}

    function renderShopping(day) {{
      shoppingGrid.innerHTML = "";
      const shopping = data.dayShopping[day.id];
      const cards = [];
      shopping.stores.forEach((store) => {{
        cards.push({{ store: store.store, items: store.items }});
      }});
      if (shopping.unassigned.length) {{
        cards.push({{ store: "Needs Store Assignment", items: shopping.unassigned }});
      }}

      if (!cards.length) {{
        const empty = document.createElement("article");
        empty.className = "panel shopping-card";
        empty.innerHTML = `<h3>No shopping list mapped</h3><p>The workbook does not tie store items to this day strongly enough to filter them.</p>`;
        shoppingGrid.appendChild(empty);
        return;
      }}

      cards.forEach((store) => {{
        const card = document.createElement("article");
        card.className = "panel shopping-card";
        card.innerHTML = `
          <span class="store-pill">${{store.store}}</span>
          <ul class="shopping-list">${{store.items.map((item) => `<li>${{item}}</li>`).join("")}}</ul>
        `;
        shoppingGrid.appendChild(card);
      }});
    }}

    function renderWeeklyShopping() {{
      weeklyShoppingGrid.innerHTML = "";
      data.weeklyShopping.forEach((store) => {{
        const card = document.createElement("article");
        card.className = "panel shopping-card";
        card.innerHTML = `
          <span class="store-pill">${{store.store}}</span>
          <ul class="shopping-list">${{store.items.map((item) => `<li>${{item}}</li>`).join("")}}</ul>
        `;
        weeklyShoppingGrid.appendChild(card);
      }});
    }}

    function renderPantry() {{
      pantryList.innerHTML = data.pantryBasics.map((item) => `<li class="pantry-pill">${{item}}</li>`).join("");
    }}

    function openRecipe(recipeId) {{
      activeRecipeId = recipeId;
      render();
      const recipeNode = document.getElementById(`recipe-${{recipeId}}`);
      if (recipeNode) {{
        recipeNode.scrollIntoView({{ behavior: "smooth", block: "start" }});
        recipeNode.classList.add("flash");
        window.setTimeout(() => recipeNode.classList.remove("flash"), 1200);
      }}
    }}

    function setExpanded(section, button, expanded, openLabel, closeLabel) {{
      section.hidden = !expanded;
      button.textContent = expanded ? closeLabel : openLabel;
    }}

    function openSection(section, button, openLabel, closeLabel) {{
      setExpanded(section, button, true, openLabel, closeLabel);
      section.scrollIntoView({{ behavior: "smooth", block: "start" }});
    }}

    function render() {{
      const day = data.days.find((entry) => entry.id === activeDayId) || data.days[0];
      selectedDayTitle.textContent = day.name;
      selectedDayCopy.textContent = `${{day.name}} is focused for fast scanning: meals first, recipes next, shopping last.`;
      renderDayChips();
      renderMeals(day);
      renderRecipes(day);
      renderDailySupply(day);
      renderShopping(day);
    }}

    toggleDayShopping.addEventListener("click", () => {{
      setExpanded(dayShoppingSection, toggleDayShopping, dayShoppingSection.hidden, "Open Day Shopping", "Close Day Shopping");
    }});
    toggleWeeklyShopping.addEventListener("click", () => {{
      setExpanded(weeklyShoppingSection, toggleWeeklyShopping, weeklyShoppingSection.hidden, "Open Weekly Store List", "Close Weekly Store List");
    }});
    toggleWeeklySupply.addEventListener("click", () => {{
      setExpanded(weeklySupplySection, toggleWeeklySupply, weeklySupplySection.hidden, "Open Weekly Supply List", "Close Weekly Supply List");
    }});
    openDayShopping.addEventListener("click", () => {{
      openSection(dayShoppingSection, toggleDayShopping, "Open Day Shopping", "Close Day Shopping");
    }});
    openWeeklyShopping.addEventListener("click", () => {{
      openSection(weeklySupplySection, toggleWeeklySupply, "Open Weekly Supply List", "Close Weekly Supply List");
      openSection(weeklyShoppingSection, toggleWeeklyShopping, "Open Weekly Store List", "Close Weekly Store List");
    }});

    renderWeeklySupply();
    renderWeeklyShopping();
    renderPantry();
    render();
  </script>
</body>
</html>
"""


def main():
    data = build_data()
    DATA_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    SITE_PATH.write_text(build_html(data), encoding="utf-8")
    print(f"Wrote {SITE_PATH}")
    print(f"Wrote {DATA_PATH}")


if __name__ == "__main__":
    main()
