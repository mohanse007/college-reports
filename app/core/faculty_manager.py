"""
Faculty Directory Manager
Handles loading, matching, and updating faculty contact information.
"""
import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DEFAULT_DIRECTORY_FILENAME = "faculty_directory.xlsx"


def normalize_name(name: str) -> str:
    """Normalize faculty name for fuzzy/flexible matching."""
    if not name:
        return ""
    # Lowercase, replace dots and hyphens with spaces, collapse multiple whitespace
    cleaned = re.sub(r'[\.\-_]', ' ', str(name).lower())
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned


class FacultyManager:
    def __init__(self, directory_path: str = None):
        if directory_path is None:
            # Look in parent directory of app (i.e. Reports/) or current working directory
            candidate = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "..", DEFAULT_DIRECTORY_FILENAME)
            candidate = os.path.normpath(candidate)
            if os.path.exists(candidate):
                self.directory_path = candidate
            else:
                self.directory_path = os.path.abspath(DEFAULT_DIRECTORY_FILENAME)
        else:
            self.directory_path = directory_path
            
        self.faculty_map = {} # normalized_name -> (original_name, phone, notes)
        self.load_directory()
        # Auto-import from Staff List Report if found
        self.import_from_staff_list_report()

    def import_from_staff_list_report(self, file_path: str = None) -> int:
        """
        Imports and merges staff details from a 'Staff List Report.xlsx' file.
        Updates missing phone numbers and adds new staff entries.
        Returns the number of updated/added entries.
        """
        if file_path is None:
            # Check candidate locations
            candidates = [
                os.path.join(os.path.dirname(self.directory_path), "Staff List Report.xlsx"),
                os.path.join(os.path.expanduser("~"), "Downloads", "Staff List Report.xlsx"),
                "Staff List Report.xlsx"
            ]
            for c in candidates:
                if os.path.exists(c):
                    file_path = c
                    break

        if not file_path or not os.path.exists(file_path):
            return 0

        try:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            existing = self.get_all()
            existing_map = {normalize_name(r["name"]): r for r in existing}
            modified_count = 0

            for r in range(2, ws.max_row + 1):
                name_val = ws.cell(row=r, column=2).value
                phone_val = ws.cell(row=r, column=3).value
                if not name_val:
                    continue

                name_str = str(name_val).strip()
                phone_str = str(phone_val).replace('.0', '').strip() if phone_val is not None else ""
                norm = normalize_name(name_str)
                if not norm:
                    continue

                matched = False
                if norm in existing_map:
                    if phone_str and not existing_map[norm]["phone"]:
                        existing_map[norm]["phone"] = phone_str
                        modified_count += 1
                    matched = True
                else:
                    # Ignore 1-letter initials when matching to prevent collisions
                    tokens = set(w for w in norm.split() if len(w) > 1)
                    for k, rec in existing_map.items():
                        k_tokens = set(w for w in k.split() if len(w) > 1)
                        if tokens and k_tokens:
                            if (tokens.issubset(k_tokens) or k_tokens.issubset(tokens)) and len(tokens & k_tokens) >= 2:
                                if phone_str and not rec["phone"]:
                                    rec["phone"] = phone_str
                                    modified_count += 1
                                matched = True
                                break
                            elif len(tokens) == 1 and tokens == k_tokens:
                                if phone_str and not rec["phone"]:
                                    rec["phone"] = phone_str
                                    modified_count += 1
                                matched = True
                                break

                if not matched:
                    new_rec = {"name": name_str.title(), "phone": phone_str, "notes": "From Staff List Report"}
                    existing.append(new_rec)
                    existing_map[norm] = new_rec
                    modified_count += 1

            if modified_count > 0:
                self.save_all(existing)

            return modified_count
        except Exception as e:
            print(f"Error importing from staff list report: {e}")
            return 0

    def load_directory(self):
        """Loads directory from Excel if it exists."""
        self.faculty_map.clear()
        if not os.path.exists(self.directory_path):
            return

        try:
            wb = openpyxl.load_workbook(self.directory_path)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                name = str(row[0]).strip()
                phone = str(row[1]).replace('.0', '').strip() if row[1] is not None else ""
                notes = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                norm = normalize_name(name)
                self.faculty_map[norm] = {
                    "name": name,
                    "phone": phone,
                    "notes": notes
                }
        except Exception as e:
            print(f"Warning: Could not read faculty directory from {self.directory_path}: {e}")

    def _lookup_single(self, faculty_name: str) -> str:
        """Looks up phone number for an individual faculty member."""
        if not faculty_name:
            return ""
        norm = normalize_name(faculty_name)
        
        # 1. Exact match
        if norm in self.faculty_map:
            return self.faculty_map[norm]["phone"]
        
        # 2. Substring / Token matching (ignoring 1-letter initials)
        norm_tokens = set(w for w in norm.split() if len(w) > 1)
        for k, info in self.faculty_map.items():
            k_tokens = set(w for w in k.split() if len(w) > 1)
            if norm_tokens and k_tokens:
                if (norm_tokens.issubset(k_tokens) or k_tokens.issubset(norm_tokens)) and len(norm_tokens & k_tokens) >= 2:
                    return info["phone"]
                elif len(norm_tokens) == 1 and norm_tokens == k_tokens:
                    return info["phone"]
        
        return ""

    def get_phone(self, faculty_name: str) -> str:
        """Looks up phone number by faculty name with fuzzy matching and multi-faculty support."""
        if not faculty_name:
            return ""
        # Handle comma-separated multiple faculty
        if ',' in str(faculty_name):
            names = [n.strip() for n in str(faculty_name).split(',')]
            phones = [self._lookup_single(n) for n in names]
            valid_phones = [p for p in phones if p]
            return ", ".join(valid_phones)
        
        return self._lookup_single(str(faculty_name).strip())

    def get_all(self):
        """Returns sorted list of all faculty records."""
        records = list(self.faculty_map.values())
        records.sort(key=lambda x: x["name"].lower())
        return records

    def save_all(self, records):
        """Saves a list of records back to faculty_directory.xlsx."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Faculty Directory"

        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        ws.append(["Faculty Name", "Phone Number", "Department / Notes"])
        for col_num in range(1, 4):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for r in records:
            ws.append([r.get("name", ""), str(r.get("phone", "")), r.get("notes", "")])

        for row_idx in range(2, ws.max_row + 1):
            for col_idx in range(1, 4):
                ws.cell(row=row_idx, column=col_idx).border = border
                ws.cell(row=row_idx, column=col_idx).font = Font(name="Calibri", size=11)

        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 25

        wb.save(self.directory_path)
        self.load_directory()
