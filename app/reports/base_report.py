"""
Base Report Module
Abstract interface for all report generator plugins.
"""
from abc import ABC, abstractmethod
from typing import Dict, Any


class BaseReport(ABC):
    """Abstract base class that all report generators must implement."""
    
    @property
    @abstractmethod
    def report_id(self) -> str:
        """Unique identifier for the report type."""
        pass

    @property
    @abstractmethod
    def display_name(self) -> str:
        """Human-readable name shown in the UI dropdown."""
        pass

    @property
    @abstractmethod
    def description(self) -> str:
        """Short description of the report and its input."""
        pass

    @abstractmethod
    def extract_metadata(self, raw_file_path: str) -> Dict[str, Any]:
        """
        Inspects the raw file and extracts dynamic parameters such as
        college name, semester number/title, admitted year, etc.
        """
        pass

    @abstractmethod
    def generate(self, raw_file_path: str, output_dir: str, options: Dict[str, Any]) -> str:
        """
        Processes raw_file_path with given options and writes the formatted
        report into output_dir. Returns the path to the generated file.
        """
        pass
