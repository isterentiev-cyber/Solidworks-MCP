# SolidWorks MCP — заметки для Claude

Форк `Solidworks-MCP` (апстрим: Samsaam Ali Baig) с точечным хардненгом.
Даёт Claude MCP-инструменты `mcp__solidworks__*` для управления живым
SolidWorks через COM (win32com). Репозиторий отдельный от `D:\Claude\infra`
— здесь всё специфично для этого инструмента.

## Как это запущено

MCP-сервер — обычный python-процесс (`solidworks_mcp/server.py`, stdio),
поднятый хостом Claude Code и живущий всю сессию. **Правки .py-файлов не
подхватываются на лету** — модуль уже импортирован в памяти процесса.
После любой правки в `solidworks_mcp/` нужно попросить пользователя
переподключить/перезапустить MCP-сервер (или перезапустить Claude Code),
и только потом проверять — иначе будешь тестировать старый код и решишь,
что фикс не сработал.

## Главная ловушка COM-обёртки (win32com dynamic dispatch)

SolidWorks API отдаёт часть zero-arg членов (`GetTypeName2`, `GetTitle`,
`FirstFeature`, `GetNextFeature`, `IsSuppressed`, `GetType`...) то как
**property** (значение уже готово при обращении к атрибуту), то как метод,
который надо вызывать `()` — в зависимости от версии SW и конкретного
члена, предсказать нельзя, надо проверять эмпирически через
`execute_python`.

Наивный фикс `if callable(x): x = x()` **не работает**, потому что
`win32com.client.CDispatch` **всегда** callable (реализует `__call__` для
COM default-member), включая случай, когда `x` — уже готовый результат
(например, следующая фича в обходе дерева). Вызов уже-готового объекта
кидает `(-2147352573, 'Member not found.', ...)`, которая выглядит как
«элементов больше нет» и рвёт обход после первого же элемента. Это и было
причиной сломанных `list_features` (падал на 1 фиче), `get_document_info`,
`list_open_documents` — см. коммит с фиксом.

**Используй `com_get(obj, name)` из `solidworks_mcp/utils/com_helpers.py`**
для любого zero-arg члена: читает атрибут, и только если результат
callable — пробует вызвать, а при ошибке вызова просто возвращает то, что
уже получил. Для методов с аргументами вызывай напрямую как обычно
(`obj.Method(arg1, arg2)`) — там неоднозначности нет.

Отдельно: методы, возвращающие коллекции/массивы (`body.GetFaces()`,
`body.GetEdges()`), в этой среде **требуют явного вызова** — они не
попадают под com_get, это настоящие методы.

Если добавляешь новый обход COM-дерева — сразу используй `com_get`,
не изобретай свою версию property/method фолбэка.

## Извлечение геометрии тела (B-Rep) — рабочий рецепт

Дерево фич (`list_features`) не показывает скругления/фаски, если они
зашиты в эскиз, а не сделаны отдельной Fillet/Chamfer-фичей. Чтобы найти
реальную геометрию — обход граней/рёбер тела через `execute_python`:

```python
doc = sw.ActiveDoc
body = doc.GetBodies2(0, True)[0]      # 0 = solid bodies, True = visible only
faces = body.GetFaces()                 # метод, нужны скобки
edges = body.GetEdges()                 # метод, нужны скобки

# ISurface.Identity (swSurfaceTypes_e)
SURF_TYPES = {4001:"PLANE",4002:"CYLINDER",4003:"CONE",4004:"SPHERE",
              4005:"TORUS",4006:"BSURF",4007:"BLENDSURF",4008:"OFFSETSURF",
              4010:"EXTRUSION",4011:"REVOLUTION"}
# ICurve.Identity — эмпирически в этой версии SW: 3001=LINE, 3002=CIRCLE
# (официальный enum swCurveTypes_e другой — не доверять доке, сверять по факту)

for face in faces:
    surf = face.GetSurface          # property, без скобок
    st = surf.Identity
    area_mm2 = face.GetArea * 1e6   # GetArea — property, без скобок
    if SURF_TYPES.get(st) == "CYLINDER":
        radius_mm = surf.CylinderParams[6] * 1000  # [ox,oy,oz,ax,ay,az,radius]
```

Кластеризация радиусов цилиндрических граней = отпечаток отверстий/бобышек/
галтелей, даже когда в дереве нет именованной Fillet/Chamfer-фичи. Конусные
грани (CONE) = фаски. Полный пример разбора см. в истории сессии
(деталь `Part Coin Hopper.SLDPRT`, разобрана 2026-09-07).

## Структура

- `solidworks_mcp/server.py` — MCP tool-хендлеры (низкоуровневый `Server` API,
  поэтому `mcp` запинен на 1.29.0, 2.x ломает).
- `solidworks_mcp/automation/` — миксины по темам (`documents.py`,
  `features.py`, `sketches.py`, `base.py` с общими хелперами и коннектом).
- `solidworks_mcp/utils/` — `com_helpers.py` (см. выше), `units.py`,
  `sw_finder.py` (автопоиск инсталляции SW через реестр).
- `execute_python` (MCP-тул) — сырой доступ к `sw` (SldWorks.Application) и
  `doc` (ActiveDoc) в том же процессе; используй для разведки перед тем как
  чинить типизированный тул, как в этой сессии.
- `backup_20260209_*/` — снапшоты старых версий файлов, не апстрим и не
  рабочий код — не редактировать, не смотреть как источник истины.

## Известные открытые вопросы

- Feature `GetTypeName2` иногда возвращает нестандартные имена (например
  `"ICE"` для Boss-Extrude в детали Coin Hopper) — не нашёл, что это
  означает; не тратить время на угадывание, если не критично для задачи.
- Параметры фич (`Depth`, `DraftAngle` и т.п. через `IExtrudeFeatureData2`)
  недоступны через `CastTo` — win32com просит `makepy`-кэш типобиблиотеки
  SolidWorks, которого в этом окружении нет. Если понадобится — сначала
  сгенерировать кэш (`makepy.py` по typelib SW), это разовая настройка
  окружения, не код фикс.
