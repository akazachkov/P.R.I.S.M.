# core/elements/pdf_finder.py

from pathlib import Path
import re
import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Optional, Tuple


class PDFFinder:
    def __init__(
        self, root_folder: str, year_labels: List[str],
        subfolder_name: str
    ):
        self.root_folder = Path(root_folder)
        self.year_labels = year_labels
        self.subfolder_name = subfolder_name

    @staticmethod
    def _get_file_mod_time(filepath: Path) -> str:
        try:
            mtime = filepath.stat().st_mtime
            dt = datetime.datetime.fromtimestamp(mtime)
            return dt.strftime("%d.%m.%Y %H:%M")
        except Exception:
            return "дата неизв."

    def normalize_number(self, number_str: str) -> str:
        """
        Приводит входную строку к нормализованному 4-значному номеру.
        - Если есть 4 цифры подряд, берёт их.
        - Если цифр меньше, дополняет нулями слева до 4.
        - Если цифр нет, возвращает исходную строку (без изменений).
        """
        original = number_str.strip()
        # 1. Пытаемся найти первые 4 цифры подряд
        match = re.search(r'(\d{4})', original)
        if match:
            return match.group(1)
        # 2. Удаляем всё, кроме цифр
        digits_only = re.sub(r'\D', '', original)
        if digits_only:
            if len(digits_only) >= 4:
                return digits_only[:4]
            else:
                return digits_only.zfill(4)
        # 3. Цифр нет — возвращаем как есть
        return original

    def find_for_number(
        self, raw_number: str
    ) -> Dict[str, List[Tuple[Path, str]]]:
        padded = self.normalize_number(raw_number)

        def search_in_year(label):
            year_path = self.root_folder / label
            if not year_path.exists():
                return label, []
            found = []
            num_pattern = re.compile(r'^(\d+)')
            for item in year_path.iterdir():
                if not item.is_dir():
                    continue
                match = num_pattern.match(item.name)
                if not match:
                    continue
                folder_num = match.group(1)
                folder_padded = folder_num.zfill(4)
                if folder_padded == padded:
                    subfolder = item / self.subfolder_name
                    if subfolder.exists():
                        for pdf in subfolder.glob("*.pdf"):
                            if pdf.is_file():
                                mod_time = self._get_file_mod_time(pdf)
                                found.append((pdf, mod_time))
            return label, found

        with ThreadPoolExecutor(max_workers=len(self.year_labels)) as executor:
            futures = [
                executor.submit(search_in_year, label)
                for label in self.year_labels
            ]
            result = {}
            for future in as_completed(futures):
                label, files = future.result()
                if files:
                    result[label] = files
            return result

    def find_for_numbers(
        self, numbers: List[str], target_year: Optional[str] = None
    ) -> Tuple[Dict[str, Dict[str, List[Tuple[Path, str]]]], List[str]]:
        # Нормализуем входные номера
        normalized_numbers = [self.normalize_number(num) for num in numbers]
        years_to_search = [target_year] if target_year else self.year_labels
        result_by_number = {norm: {} for norm in normalized_numbers}
        not_found_set = set(normalized_numbers)

        for year in years_to_search:
            year_path = self.root_folder / year
            if not year_path.exists():
                continue

            for item in year_path.iterdir():
                if not item.is_dir():
                    continue
                match = re.match(r'^(\d+)', item.name)
                if not match:
                    continue
                folder_num = match.group(1)
                folder_padded = folder_num.zfill(4)
                if folder_padded not in result_by_number:
                    continue

                subfolder = item / self.subfolder_name
                if not subfolder.exists():
                    continue

                files = list(subfolder.glob("*.pdf"))
                if files:
                    for pdf in files:
                        mod_time = self._get_file_mod_time(pdf)
                        if year not in result_by_number[folder_padded]:
                            result_by_number[folder_padded][year] = []
                        result_by_number[folder_padded][year].append((
                            pdf, mod_time
                        ))
                    not_found_set.discard(folder_padded)

        return result_by_number, list(not_found_set)
