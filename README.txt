Изменения (18.12.2025)

1) Исправлена ошибка:
   AttributeError: module 'table_processor' has no attribute 'postprocess_equipment_dates'
   -> функция postprocess_equipment_dates добавлена в table_processor.py.

2) Добавлено разбиение Таблицы 3 (помещения):
   Если строк больше ROOM_TABLE_ROWS_PER_PAGE (по умолчанию 12),
   таблица автоматически режется на несколько таблиц с разрывом страницы.
   На каждой новой странице перед таблицей добавляется строка:
   "Продолжение таблицы 3"

3) Отчет:
   Для генерации отчета используйте тот же пайплайн, что и для протокола,
   просто подайте другой шаблон (Шаблон ОТЧ-OQ.docx) и сохраните отдельным файлом.
   Пример кода вставки в RenderWorker.run() (см. ниже):

   def render_one(template_path, out_path):
       doc = render_template(template_path, context)  # как в протоколе
       process_rooms_table(doc, rooms)  # теперь уже с автопереносом Табл.3
       process_equipment_table(doc, equipment_rows)
       postprocess_equipment_dates(doc)
       ... (остальные ваши постпроцедуры как для протокола)
       doc.save(out_path)

   render_one(protocol_tpl, protocol_out)
   render_one(report_tpl, report_out)

Если нужно — поменяйте ROOM_TABLE_ROWS_PER_PAGE в table_processor.py.
