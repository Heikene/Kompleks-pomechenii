# file_utils.py
from tempfile import NamedTemporaryFile
from contextlib import contextmanager

@contextmanager
def temp_docx(suffix: str = ".docx"):
    f = NamedTemporaryFile(suffix=suffix, delete=False)
    try:
        yield f.name
    finally:
        try:
            f.close()
            import os
            os.remove(f.name)
        except OSError:
            pass
