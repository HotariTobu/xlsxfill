"""Exceptions raised by xlsxfill."""


class XlsxfillError(Exception):
    """Base class for all exceptions raised by xlsxfill."""


class DataError(XlsxfillError):
    """The input data as a whole is unusable; substitution never starts."""
