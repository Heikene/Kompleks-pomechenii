# logger.py
import logging

logger = logging.getLogger("OQGen")
logger.setLevel(logging.DEBUG)

# файл
fh = logging.FileHandler("oq_generator.log", encoding="utf-8")
fmt = logging.Formatter("%(asctime)s %(levelname)s: %(message)s")
fh.setFormatter(fmt)
logger.addHandler(fh)

# консоль (по желанию)
ch = logging.StreamHandler()
ch.setFormatter(fmt)
logger.addHandler(ch)
