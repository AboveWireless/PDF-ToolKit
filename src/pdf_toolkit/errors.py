from __future__ import annotations


class PdfToolkitError(Exception):
    exit_code = 4


class ValidationError(PdfToolkitError):
    exit_code = 2


class DependencyMissingError(PdfToolkitError):
    exit_code = 3


class ProcessingFailureError(PdfToolkitError):
    exit_code = 4
