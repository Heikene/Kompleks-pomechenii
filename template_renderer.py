# template_renderer.py

from typing import Dict, Any, List, Optional
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Mm
from logger import logger
from file_utils import temp_docx

AREA_THRESHOLDS = [
    (2,1),(4,2),(6,3),(8,4),(10,5),
    (24,6),(28,7),(32,8),(36,9),(52,10),
    (56,11),(64,12),(68,13),(72,14),(76,15),
    (104,16),(108,17),(116,18),(148,19),(156,20),
    (192,21),(232,22),(276,23),(352,24),(436,25),
    (636,26),(1000,27),
]

def _calc_points(area: float) -> int:
    for thr, n in AREA_THRESHOLDS:
        if area <= thr:
            return n
    return 27
# -------------------------------------------

def build_context(
    ctx_fields: Dict[str, str],
    rooms: List[Dict[str, str]],
    tests_placeholder: str = "{{TABLE}}"
) -> Dict[str, Any]:
    """
    Собирает контекст для рендеринга Jinja-плейсхолдеров: добавляет в rooms поле point.
    """
    rooms_ext: List[Dict[str, Any]] = []
    for r in rooms:
        a = str(r.get("area", "")).replace(",", ".")
        try:
            nl = _calc_points(float(a))
        except ValueError:
            nl = 1  # дефолт, если площадь не парсится
        rooms_ext.append({**r, "point": nl})

    ctx = ctx_fields.copy()
    ctx.update({
        "rooms": rooms_ext,
        "TABLE": tests_placeholder,
        "TABLE5": "{{TABLE5}}",
    })
    logger.debug("Контекст для Jinja собран (rooms с point)")
    return ctx

def render_template(
    tpl_path: str,
    context: Dict[str, Any],
    out_path: Optional[str] = None
) -> Document:
    """
    Рендерит DOCX по шаблону.

    :param tpl_path: путь к исходному .docx-шаблону
    :param context: словарь полей для подстановки, включая 'Scan_paths'
    :param out_path: если указан — сохранить итоговый документ сразу по этому пути
    :return: объект python-docx Document для дальнейшей обработки
    """
    tpl = DocxTemplate(tpl_path)

    # --- 1) Подготовка InlineImage для нескольких сканов ---
    if "Scan_paths" in context:
        paths = context.pop("Scan_paths")
        images: List[InlineImage] = []
        for p in paths:
            try:
                img = InlineImage(tpl, p, width=Mm(150))
                images.append(img)
                logger.debug(f"InlineImage создан для {p}")
            except Exception as ex:
                logger.error(f"Не удалось создать InlineImage для {p}: {ex}")
        # в шаблоне используем {% for img in Scans %}{{ img }}{% endfor %}
        context["Scans"] = images

    # --- 2) (если нужен единичный режим) InlineImage для одного скана ---
    if "Scan_path" in context:
        img_path = context.pop("Scan_path")
        try:
            img = InlineImage(tpl, img_path, width=Mm(100))
            context["Scan"] = img
            logger.debug(f"InlineImage создан для {img_path}")
        except Exception as ex:
            logger.error(f"Не удалось создать InlineImage для {img_path}: {ex}")

    # --- 3) рендер всех плейсхолдеров ---
    tpl.render(context)

    # --- 4) сохранение и возврат Document ---
    if out_path:
        tpl.save(out_path)
        logger.info(f"Шаблон отререндерен и сохранён в {out_path}")
        return Document(out_path)

    with temp_docx() as tmp:
        tpl.save(tmp)
        doc = Document(tmp)
        logger.info("Шаблон отререндерен во временный файл")
        return doc
