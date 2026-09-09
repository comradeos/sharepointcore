#!/usr/bin/env python3
"""Генерує тестовий CSV для імпорту в список SharePoint.

Файл створюється в тій самій папці, що й цей скрипт. Дані мають UTF-8 BOM,
тому коректно відкриваються в Excel та розпізнаються під час імпорту CSV.
"""

from __future__ import annotations

import csv
import random
import sys
from datetime import date, datetime, time, timedelta
from pathlib import Path
from urllib.parse import quote


SCRIPT_DIRECTORY = Path(__file__).resolve().parent

FIRST_NAMES = (
    "Олена", "Андрій", "Наталія", "Віктор", "Ірина", "Дмитро",
    "Марія", "Сергій", "Тетяна", "Олексій",
)
LAST_NAMES = (
    "Коваль", "Шевченко", "Бондаренко", "Мельник", "Ткаченко",
    "Кравченко", "Олійник", "Петренко", "Савчук", "Романенко",
)
STATUSES = ("Новий", "У роботі", "На погодженні", "Виконано", "Призупинено")
CATEGORIES = ("Енергетика", "ІТ", "Закупівлі", "Безпека", "Документообіг")
AUDIENCES = ("Усі працівники", "Керівники", "ІТ-відділ", "Відділ закупівель")
LOCATIONS = (
    "Київ, вул. Симона Петлюри, 25",
    "Львів, вул. Городоцька, 179",
    "Дніпро, просп. Дмитра Яворницького, 75",
    "Одеса, вул. Канатна, 83",
)

# Назви відповідають призначенню полів зі списку типів SharePoint.
COLUMNS = (
    "Назва",
    "Детальний опис",
    "Кількість",
    "Активний",
    "Відповідальна особа",
    "Дата та час",
    "Статус",
    "Посилання",
    "Вартість, грн",
    "Розташування",
    "Фото",
    "Категорія",
    "Цільова аудиторія",
    "Швидка дія",
)


def ask_file_name() -> str:
    """Запитує безпечне ім'я CSV-файлу без шляху до іншої папки."""
    while True:
        value = input("Назва CSV-файлу (наприклад, test-sharepoint): ").strip()
        if not value:
            print("Вкажіть назву файлу.")
            continue

        candidate = Path(value)
        if candidate.name != value or value in {".", ".."}:
            print("Вкажіть лише назву файлу без шляху до папки.")
            continue

        return value if value.lower().endswith(".csv") else f"{value}.csv"


def ask_row_count() -> int:
    """Запитує додатну кількість рядків."""
    while True:
        value = input("Кількість рядків: ").strip()
        try:
            count = int(value)
        except ValueError:
            print("Введіть ціле додатне число, наприклад 1000.")
            continue

        if count < 1:
            print("Кількість рядків має бути більшою за нуль.")
            continue
        return count


def confirm_overwrite(path: Path) -> bool:
    if not path.exists():
        return True

    while True:
        answer = input(f"Файл «{path.name}» вже існує. Перезаписати? [так/ні]: ").strip().lower()
        if answer in {"так", "т", "yes", "y", "д", "дa"}:
            return True
        if answer in {"ні", "н", "no", "n"}:
            return False
        print("Введіть «так» або «ні».")


def progress_bar(current: int, total: int, width: int = 34) -> None:
    """Виводить компактний індикатор, не потребуючи сторонніх бібліотек."""
    filled = int(width * current / total)
    bar = "█" * filled + "░" * (width - filled)
    percent = current * 100 / total
    print(f"\r[{bar}] {percent:6.2f}%  {current:,}/{total:,} рядків", end="", flush=True)


def create_row(index: int, rng: random.Random) -> list[str | int | float]:
    """Створює один рядок із прикладами даних для всіх типів полів."""
    first_name = rng.choice(FIRST_NAMES)
    last_name = rng.choice(LAST_NAMES)
    status = rng.choice(STATUSES)
    category = rng.choice(CATEGORIES)
    timestamp = datetime.combine(
        date(2026, 1, 1) + timedelta(days=index % 365),
        time(hour=8 + index % 10, minute=(index * 7) % 60),
    )
    slug = quote(f"запис-{index:06d}")

    return [
        f"Запис {index:06d}",
        f"Тестовий багаторядковий опис для запису {index:06d}.\n"
        f"Статус: {status}. Категорія: {category}.",
        index * 3,
        "Так" if index % 2 else "Ні",
        f"{first_name} {last_name} ({first_name.lower()}.{last_name.lower()}@example.com)",
        timestamp.strftime("%Y-%m-%d %H:%M"),
        status,
        f"https://example.com/records/{slug}",
        f"{rng.uniform(1000, 250000):.2f}",
        rng.choice(LOCATIONS),
        f"https://picsum.photos/seed/sharepoint-{index}/640/480",
        category,
        rng.choice(AUDIENCES),
        f"Відкрити запис|https://example.com/actions/{slug}",
    ]


def generate_csv(path: Path, row_count: int) -> None:
    rng = random.Random(20260909)
    refresh_every = max(1, min(10_000, row_count // 200))

    with path.open("w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(COLUMNS)

        for index in range(1, row_count + 1):
            writer.writerow(create_row(index, rng))
            if index == row_count or index % refresh_every == 0:
                progress_bar(index, row_count)

    print()


def main() -> int:
    print("Генератор тестового CSV для SharePoint")
    print(f"Папка результату: {SCRIPT_DIRECTORY}")

    file_name = ask_file_name()
    output_path = SCRIPT_DIRECTORY / file_name
    row_count = ask_row_count()

    if not confirm_overwrite(output_path):
        print("Генерацію скасовано. Файл не змінено.")
        return 0

    print(f"Створення «{output_path.name}»…")
    generate_csv(output_path, row_count)
    print(f"Готово: {output_path}")
    print(f"Створено {row_count:,} рядків і {len(COLUMNS)} колонок.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nГенерацію скасовано користувачем.")
        raise SystemExit(130)
