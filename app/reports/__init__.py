"""
Reports Registry Module
Automatically discovers and registers all report generators.
"""
from typing import Dict, List, Type
from .base_report import BaseReport
from .course_faculty_report import CourseFacultyReport

# Registry list of available reports
# When you create a new report, simply import it and add it to AVAILABLE_REPORTS!
AVAILABLE_REPORTS: List[Type[BaseReport]] = [
    CourseFacultyReport,
]


def get_available_reports() -> Dict[str, BaseReport]:
    """Returns an instantiated dictionary of available report generators keyed by report_id."""
    reports = {}
    for report_cls in AVAILABLE_REPORTS:
        instance = report_cls()
        reports[instance.report_id] = instance
    return reports


__all__ = ["BaseReport", "CourseFacultyReport", "get_available_reports", "AVAILABLE_REPORTS"]
