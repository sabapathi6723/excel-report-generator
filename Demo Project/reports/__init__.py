"""
Reports module for generating Excel reports.
"""

from .participation import generate_participation_report
from .performance import generate_performance_report

__all__ = [
    'generate_participation_report',
    'generate_performance_report',
]

