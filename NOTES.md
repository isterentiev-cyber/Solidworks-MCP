# SolidWorks MCP — подробные заметки

Находки по темам. `CLAUDE.md` — короткий индекс со ссылками сюда; этот файл
целиком не читать, искать по заголовку. Здесь только то, что верно сейчас:
как нашлось и чем доказано — в `git log` (сжато 2026-09-23, прежняя
версия — коммит до этой даты).

## win32com dynamic dispatch: property vs method — неоднозначность

Часть zero-arg членов (`GetTypeName2`, `GetTitle`, `FirstFeature`,
`GetNextFeature`, `IsSuppressed`, `GetType`…) в динамической обёртке то
свойство, то метод. `if callable(x): x = x()` **не работает**: `CDispatch`
всегда callable, вызов готового объекта кидает `Member not found`, и обход
дерева рвётся после первого элемента.

- Старый код: `com_get(obj, name)` (`utils/com_helpers.py`) — читает, при
  callable пробует вызвать, при ошибке возвращает прочитанное.
- Новый код: `T(obj, "IFace2")` + `v(obj, "Name")` — см. § Типизированные
  обёртки.
- Методы с аргументами и коллекции (`body.GetFaces()`, `GetEdges()`) —
  обычный явный вызов.

## Скрин детали + визуальная проверка геометрии

`capture_view` → PNG (`SaveBMP` + Pillow) → смотреть через `Read`. Это
качественная проверка; авторитетны цифры из B-Rep (`list_faces`/
`list_edges`, § Извлечение геометрии). Имени файла не доверять: скрин —
обязательный шаг перед описанием детали. Для чертежа надёжнее
`export_pdf` + рендер страницы (pymupdf): `capture_view` снимает окно SW и
обрезает лист по его пропорциям.

## Извлечение геометрии тела (B-Rep) — рабочий рецепт

Нужен, когда дерево фич не показывает форму (скругления в эскизе и т.п.).
Готовые таблицы — тулы `inspect`, `list_faces`, `list_edges`. Вручную:

```python
body = T(md, "IPartDoc").GetBodies2(0, True)[0]   # 0 = solid, True = visible
for f in body.GetFaces():                          # метод, со скобками
    s = T(T(f, "IFace2").GetSurface(), "ISurface")
    # s.Identity(): 4001 PLANE, 4002 CYLINDER, 4003 CONE, 4004 SPHERE,
    # 4005 TORUS, 4006 BSURF, 4007 BLEND, 4008 OFFSET, 4009 EXTRU, 4010 SREV
    # s.CylinderParams -> [ox,oy,oz, ax,ay,az, r] (м)
# ICurve.Identity: 3001 LINE, 3002 CIRCLE, 3003 ELLIPSE, 3004 INTERSECTION,
# 3005 BCURVE, 3006 SPCURVE, 3008 CONSTPARAM, 3009 TRIMMED
```

Кластеры радиусов цилиндров = отверстия/бобышки/галтели; CONE = фаски.

## Справочная геометрия, масса, ошибки перестроения, правка фич

- **`create_reference_plane`**: `InsertRefPlane(constraint, value, 0,0,0,0)`,
  значение в м / рад. `swRefPlaneReferenceConstraint_*`: Parallel 1,
  Perpendicular 2, Coincident 4, Distance 8, Angle 16, Tangent 32,
  MidPlane 128, флаг OptionFlip **256** (OR с основной). Созданную фичу
  API не отдаёт надёжно — берётся разницей дерева до/после.
- **`create_reference_axis`**: `InsertAxis2(True)` по выбранному. Две
  параллельные плоскости ось не задают (False). Ось из двух плоскостей
  смотрит в -Z. Ссылки обоих тулов: `front`/`top`/`right`, имя фичи,
  `face:N`/`edge:N`.
- **`mass_properties`**: `CreateMassProperty2` (откат на
  `CreateMassProperty`), **`UseSystemUnits = True`** — иначе числа в единицах
  документа.
- **`get_rebuild_errors`**: `IFeature.GetErrorCode2()` по дереву + статус
  эскизов; коды → имена `swFeatureError*` (`FilletRadiusTooBig2 (19)`).
  `GetWhatsWrong` (три by-ref массива, молча пустой) не используется, только
  `GetWhatsWrongCount` как подтверждение. SW переваривает многое —
  надёжно ломается только невозможный радиус уже существующего скругления.
- **`edit_feature`**: размеры по коротким именам, одно `EditRebuild3`. Полное
  имя (`D1@Sketch1`) уже умеет `set_parameter`, список — `get_parameters`.
  Правка определения фичи (граничное условие, направление) не покрыта.

## Перезагрузка сервера: новые тулы без рестарта клиента

- `reload_api` перезагружает `automation/*`, `ext.py`, `toolsets.py`
  (COM-коннект сохраняется) и шлёт `tools/list_changed`.
- **Claude Desktop это уведомление игнорирует**: список тулов фиксируется на
  старте разговора. Новый тул виден только в НОВОМ чате, рестарт сервера не
  помогает. В текущем чате — `execute_python` → `ext.t_имя(sw, {...})`.
- `server.py` не перезагружается никогда — его правки требуют рестарта
  процесса: `scripts/restart_mcp.ps1` (`-List`, `-Id <PID>`, без аргументов —
  все). SW не трогается, документы остаются открытыми.
- Серверов несколько (по одному на сессию), все маршалят COM в один
  SLDWORKS.exe → межпроцессные гонки возможны. После убийства процесса
  клиент поднимает сервер лениво, на первом вызове;
  `session_connectors_status` показывает кэш, а не живое состояние.

## Проба занятости: отказ вместо зависания

COM-вызов в модальный или занятый SW не падает, а блокирует весь сервер.
Перед каждым тулом `ext.sw_busy()`:
`SendMessageTimeout(hwnd, WM_NULL, SMTO_BLOCK|SMTO_ABORTIFHUNG, timeout)`,
~0,2–0,3 мс в простое.

- Порог `SW_MCP_BUSY_TIMEOUT_MS` (по умолчанию 1200, `0` — выкл).
- hwnd (`Frame().GetHWndx64()`, откат на `GetHWnd()`) кэшируется и
  проверяется `IsWindow` — его получение само COM-вызов и зависло бы.
- Проба не бросает и при сомнении отвечает «готов».
- `BUSY_EXEMPT` (`server.py`): connect, get_solidworks_info, lookup_*,
  reload_api, set_units.
- Не ловит SW, ставший модальным во время вызова: для этого нужен
  `IMessageFilter` с повтором на `RPC_E_CALL_REJECTED` (не сделан).

## Индекс API (scripts/swapi.py)

```
python scripts/swapi.py find shell          # поиск по всему API
python scripts/swapi.py find drawing -k m   # только методы (m/p/c)
python scripts/swapi.py sig IModelDoc2 InsertFeatureShell
python scripts/swapi.py const swEndCondBlind
python scripts/swapi.py build | stats
```

Отвечает на «каким методом это делается» (у `lookup_api_signature` имя уже
надо знать). ~9,8k методов, 3,4k свойств, 8,3k констант, ~900 КБ в
`api-index/`. SW запускать не нужно — строится из makepy-кэша.
Ранжирование: совпадение в имени члена выше, чем в параметрах.

Кэш makepy — `<repo>/.gen_py` (переопределить: `SW_MCP_GEN_PY`), логика в
`utils/typelib.py::_relocate_gen_cache`. Две ловушки: `dicts.dat`
перечитывается явно (загружен при импорте по старому пути);
`gencache.Rebuild(verbose=0)` — иначе прогресс уходит в stdout = протокол
MCP.

## Фильтр тулов по разделам (SW_MCP_TOOLSETS)

`solidworks_mcp/toolsets.py`. Разделы: `core` (всегда), `select`, `sketch`,
`features`, `analysis`, `drawings`. `SW_MCP_TOOLSETS=sketch,features` в
`env` сервера в `claude_desktop_config.json`, пусто = все. Неизвестный
раздел игнорируется с предупреждением; тул, не попавший ни в один раздел,
остаётся видимым. Нужен рестарт MCP.

## Не угадывай сигнатуры — используй lookup_api_signature / lookup_api_constant

Память и документация расходятся с этой инсталляцией. Так нашлись
`FeatureRevolve2` — **20** параметров (не 18: `OffsetDistance1/2` между
`OffsetReverse2` и `ThinType`) и `InsertFeatureChamfer` — **8** (не 7).
Реализация — `utils/typelib.py`. Типобиблиотеки — `sldworks.tlb` +
`swconst.tlb` в каталоге SW. `makepy.GenerateFromTypeLibSpec` принимает
объект из `selecttlb.EnumTlbs()`, голый кортеж → `Library not registered`.
Кэш не коммитить (машинно-зависимый).

## Построение вала/тела вращения

- **Основной путь — `revolve_sketch`** по одному полностью определённому
  профилю. Осевую не рисовать: `revolve_sketch(axis="X"|"Y"|"Z")`. Если
  осевая в эскизе есть, она должна быть ровно одна. См. § Revolve вокруг оси.
- Запасной путь — цепочка соосных Boss-Extrude от торца к торцу.
- Слишком близкие шаги профиля (Δr ~0,5 мм) иногда не дают `CreateLine`
  (None без ошибки) — объединять шаги.

## Sketch на цилиндрической (изогнутой) грани — не работает через InsertSketch2

- `InsertSketch2(True)` на CYLINDER-грани тихо не открывает эскиз
  (`ActiveSketch` = None) при любом способе выбора, на новом документе и
  после рестарта SW. Плоские грани работают. **Тупик — не повторять.**
- **Обход для пазов**: смещённая плоскость, касательная к валу
  (`create_reference_plane kind=distance`), эскиз на ней.
- **Открыто (2026-09-08)**: `FeatureCut3/4` по прямоугольнику на такой
  плоскости режет на правильную глубину, но по всей длине вала, игнорируя
  границы профиля по оси. `UseFeatScope`/`UseAutoSelect` не помогли. Что
  пробовать: паз уже, чтобы все углы профиля лежали в материале; рез до
  граней вместо blind; иначе — паз вручную в UI.
- Координаты сущностей эскиза — локальные 2D, по ним нельзя узнать
  ориентацию плоскости. Рамку пишет `create_sketch*`. В этом шаблоне
  Front — нормаль X, Right — Y, Top — Z.

## CreateCircle тихо возвращает None без AddToDB/AutoInference

`CreateCircle`, `CreateCircleByRadius`, `CreateCornerRectangle` и
автоинференс при `CreateLine` (точки тихо съезжают к соседним) лечатся одним
обходом — `with ext.NoInference(sm):` (`AddToDB=True`,
`AutoInference=False`). Уже внутри `draw_circle`, `draw_profile`,
`draw_rectangle`, разрезов и выносных видов. В сыром коде через
`execute_python` — оборачивать самому.

## FeatureCut3/4 на эскизе, лежащем на референс-плоскости (не на грани тела)

Эскиз на Front/Top/Right режет только с `Dir=True`, на грани — с
`Dir=False`. `cut_extrude` проверяет, что объём уменьшился, и при пустом
вырезе повторяет с другим `Dir` (в ответе `[Dir=…]`). Сырой вызов:
`FeatureCut3(True, False, True, 1, 0,0,0, False,False,False,False, 0,0,
False,False,False,False,False, True,True,True,True, False, 0,0, False)`
(`T1=1` = ThroughAll).

## Сборки (IAssemblyDoc): вставка компонентов и мейты через API

Рабочий рецепт: `ladder-panel.SLDASM` (проект Demo, 2026-09-12). Тулов для
сборок нет, только `execute_python`.

- **`AddComponent5(path, 0, "", False, "", X, Y, Z)`**: X,Y,Z — **центр
  bbox** детали, а не её начало. Проверка — `component.Transform2.ArrayData`
  (`[9:12]` сдвиг, `[0:9]` поворот).
- Фиксация: выбрать `COMPONENT` → `doc.FixComponent()`.
- Имя для `SelectByID2`: `"Фича@Компонент-1@Сборка"` — третий сегмент
  обязателен, без него тихо False. Сам компонент — `"Компонент-1@Сборка"`.
- **`AddMate5`**: 14 входных + out `ErrorStatus` =
  `VARIANT(VT_BYREF|VT_I4, 0)`. Обе стороны выбирать с `Mark=1`.
  `swMateCOINCIDENT=0`, `CONCENTRIC=1`, `DISTANCE=5`.
- Вторая копия детали параллельно первой без остаточных DOF:
  2× Coincident по одноимённым плоскостям поперёк сдвига + Distance по
  плоскости вдоль.
- **Выбор грани по координате в сборке неоднозначен**, если в точке
  совпадают элементы разных компонентов: берётся не та грань, `AddMate5` →
  None. Черновые компоненты ставить в заведомо пустое место.

## Типизированные обёртки (makepy) — ловушки

Объекты приходят то типизированными, то динамическими (`ActiveDoc` —
динамический), поэтому новый код оборачивает всё сам: `ext.T(obj, "IFace2")`,
члены — `ext.v(obj, "Name")`.

- `T()` надёжен только для «своего» интерфейса: `T(sketch, "IFeature")`
  проходит QI, но диспетчит по ISketch. Фичу эскиза искать через
  `ext.sketch_feature`.
- `ISldWorks.RevisionNumber`, `IFeature.GetTypeName2`, `IsSuppressed`,
  `IEquationMgr.Equation/Value/GlobalVariable` — в typed-обёртке методы.
- `SelectByID2` Callout: typed хочет `None`, динамический —
  `VARIANT(VT_DISPATCH)`. `ext.select_by_id` пробует оба.
- **`SelectByID2("", "FACE", x,y,z)` → False** даже на плоской грани;
  `ext.select_face_at` (`GetClosestPointOn` + `SelectByRay`).
- `InsertRefPlane` возвращает `IRefPlane`, не `IFeature` (новая фича —
  последняя в дереве). Double-параметры передавать float.
- `IMathUtility.CreatePoint` через typed даёт мусор (1e-311). Считать самому
  по `IMathTransform.ArrayData`: `p' = (p·R)·scale + t`, R = a[0:9],
  t = a[9:12], scale = a[12] (`ext.xform`).
- **Массив в свойство** (`IView.Position` и т.п.) — только
  `VARIANT(VT_ARRAY|VT_R8, [...])`; кортеж молча пишет мусор.
- Середина LINE-ребра — среднее концов (`Evaluate2` по середине
  параметра уезжает за ребро), остальных — `GetClosestPointOn`.
- `GetFirstDisplayDimension` у фичи без размеров → int 0 (makepy падает):
  `ext.display_dims()`.

## Мелкие исправления тулов (2026-09-19)

Уже в коде; записано, чтобы не «чинить» обратно:
`FeatureFillet3` — 14 параметров, рёбра с Mark=1; `InsertFeatureChamfer`
угол-дистанция = `swChamferAngleDistance` = 1; `swDelete_Absorbed` = **2**,
`swDelete_Children` = 1, `DeleteSelection2(0)` удаляет фичу, оставляя
эскиз; extrude/cut отказываются брать уже поглощённый эскиз;
`revolve_sketch(cut=True)` — `IsSolid` всегда True, режет `IsCut` (иначе
Surface-Revolve); `Save3` — два out-параметра, успех = True
(`ext.save_in_place`).

## HoleWizard5 — карта Value-слотов (проверено SW 2026)

Грань выбирать **с точкой** (`select_face_at`) — отверстие встаёт в неё.
Сигнатура: `(GenericHoleType, StandardIndex, FastenerTypeIndex, SSize,
EndType, Diameter, Depth, Length, Value1..Value12, ThreadClass, RevDir,
FeatureScope, AutoSelect, AssemblyFeatureScope, AutoSelectComponents,
PropagateFeatureToParts)`. `-1` — не универсальное «по умолчанию»:

| тип | что работает |
|---|---|
| `swWzdHole` (2), ISO drill sizes (143) | `SSize="Ø6.0"` (с `.0`, U+00D8), Value2 = угол сверла (рад), остальные 0. С -1 и through-all → None |
| `swWzdTap` (4), ISO tapped hole (147) | Diameter = сверло, Depth = глубина сверла, **Value1 = глубина резьбы**, остальные -1 (Value6=0 или Value1=0 → дюймовый профиль). Диаметр резьбы — из SSize |
| `swWzdCounterBore` (0), ISO 4762 (139) | Diameter=-1, все Value=-1 → ISO-посадки (M5: Ø5.5 / Ø10×5.4) |

Правка после создания — `IWizardHoleFeatureData2` (`AccessSelections` →
поля → `ModifyDefinition(data, md, None)`):
- косметическая резьба: `Type` = 46 (глухая) / 48 (сквозная),
  `CosmeticThreadType = 1` (по умолчанию «remove thread», Type 31);
- угол сверла у резьбы — `DrillAngle = radians(118)` **отдельным**
  `ModifyDefinition` (смена Type в том же вызове сбрасывает угол);
- первый подэскиз — точки размещения, второй — профиль.

## Уравнения (IEquationMgr)

- `Add3` → -1 на любом входе; **`Add2(-1, eq, True)` работает**.
- `SetEquationAndConfigurationOption` → -1; **`SetEquation(i, eq)`** работает.
- Размер: `md.Parameter("D1@Boss-Extrude1")` → `IDimension.SetSystemValue3(м,
  swAllConfiguration, None)`. Привязка — уравнение `"D1@Boss-Extrude1" = "h"`.

## Массивы и зеркало

- `FeatureCircularPattern5`: ось Mark 1, фичи Mark 4 (ось-фича или
  цилиндрическая грань). `InsertMirrorFeature2`: плоскость Mark 2, фичи Mark 1.
  **Оба → None без диагностики, если экземпляр ложится в существующий вырез.**
- `FeatureLinearPattern5`: направление Mark 1 (второе Mark 2), фичи Mark 4.
- `md.BlankRefGeom()` / `UnBlankRefGeom()` — скрыть/показать ось/плоскость.

## Транзакции

`Extension.StartRecordingUndoObject()` → операции →
`FinishRecordingUndoObject2(name, False)`, откат — `EditUndo2(1)`: один шаг
откатывает всё записанное. `transaction abort` после undo добивает
оставшиеся новые фичи.

## Определённость эскизов (sketch_entities / add_sketch_relation / add_sketch_dimension)

Эскиз, ведущий фичу, — **полностью определён**; `delta` ставит ⚠ на фичу по
недоопределённому эскизу.

- Статус `ISketch.GetConstrainedStatus()`: 2 under, 3 fully, 4 over,
  5 no solution. Сегменты — тип 0 line, 1 arc/circle. Начало координат —
  `SelectByID2("Point1@Origin", "EXTSKETCHPOINT")`.
- Связи: `Select4` сущностей → `md.SketchAddConstraints("sgCONCENTRIC")`.
  Горизонталь двух точек — `sgHORIZONTALPOINTS2D`, линии — `sgHORIZONTAL2D`.
- Размеры: `AddDiameterDimension2` / `AddHorizontalDimension2` /
  `AddVerticalDimension2` / `AddRadialDimension2` / `AddDimension2(x,y,z)`
  (координаты модельные). Сначала выключить диалог:
  `SetUserPreferenceToggle(swInputDimValOnCreate=10, False)`.
- Окружности, нарисованные в одну точку центра, делят центр (уже
  concentric).
- Привязка к переменной (уравнение + EvaluateAll) выкидывает из режима
  эскиза — тул возвращает его `EditSketch`.

## Revolve вокруг оси вне эскиза

`FeatureRevolve2` без осевой: эскиз Mark 0 + ось (фича или ребро) **Mark 4**.
Ось создавать ДО выбора эскиза (создание сбрасывает выбор). Ось из двух
плоскостей смотрит в -Z; сторона разворота меняется без пересоздания через
`IRevolveFeatureData2.ReverseDirection` + `ModifyDefinition`.

## Фаска на кромке тор/плоскость

`InsertFeatureChamfer` на торце гнутой трубы даёт грань BSURF, а не конус.
Сторона меряется вдоль тора. Проверять по площади торцевого кольца.

## Чертежи (IDrawingDoc): виды и аннотации

Проверено 2026-09-22/23 (SW 2026, ГОСТ, `kolco-flanec-1230.SLDDRW` +
черновики). Тулы: `add_drawing_view`, `add_section_view`,
`add_broken_out_section`, `add_detail_view`, `move_drawing_view`,
`delete_drawing_view`, `add_drawing_dimension`, `add_gtol`, `add_datum`,
`add_surface_finish`, `delete_annotation`. Находки по
`insert_model_dimensions` и `add_note` — в докстрингах `ext.py`.

### Координаты: лист ≠ эскиз вида

- После `ActivateView` все `SketchManager.Create*` рисуют в **эскизе вида**:
  масштаб модели, начало — центр вида (`IView.Position`). Центр вида ≠
  проекция начала модели у несимметричной детали.
- Лист(м) → эскиз вида(м): `ISketch.ModelToSketchTransform` у
  `IView.GetSketch()` (его «модель» — лист; `a[12] = 1/scale`):
  `ext.sheet_to_view_sketch` / `view_sketch_to_sheet` (офлайн-тест в
  `scripts/selftest.py`). Модель → лист: `IView.ModelToViewTransform`
  (`ext.model_to_sheet`).
- Публичный API тулов — мм листа. `add_section_view` сверяет концы линии
  обратным преобразованием (⚠ при > 0,01 мм).

### Сечение тела вращения — местный разрез, а не вид + разрез А-А

Так делает пользователь (View9 на `kolco-flanec-1230`, 2026-09-23): **один
вид сбоку + местный разрез вокруг всего вида, глубина = наружный радиус**.
Получается полный разрез со штриховкой: без вида с торца (пустые окружности
на пол-листа), без линии сечения и подписи «А–А».

- `add_broken_out_section(view)` — по умолчанию контур = габарит вида +2 мм,
  глубина = половина габарита детали вдоль взгляда (у кольца Ø1230 — 615).
  Свой контур — `points` (мм листа), своя глубина — `depth`.
- API: замкнутый контур в эскизе активного вида (4× `CreateLine` под
  `NoInference`) → выбрать все сегменты → `CreateBreakOutSection(depth_м)`
  → True. Глубина — **модельная** дистанция от ближней точки детали.
- Контур поглощается фичей (`DrBreakoutSectionLine` под фичей вида): после
  создания `IBrokenOutSectionFeatureData.SketchSegment` → None, число
  сегментов → -1, подэскизы пустые — **форму контура не прочитать**.
  Глубина читается (`Depth`, м).
- Размеры к рёбрам внутри местного разреза ставятся как обычно
  (`add_drawing_dimension`, проверено: Ø830 и поясок 5 мм).
- Вид с торца добавлять, только если на нём есть информация: отверстия,
  лыски, пазы.

### Удаление видов и сегментов эскиза вида

- Сегмент эскиза вида удаляется, **только пока его вид активен**; иначе
  `Select4` → True, а удаление молча ничего не делает.
- Удалённый разрез оставляет в родителе осиротевшую линию сечения; её
  удаление (`SelectByID2(имя, "SECTIONLINE")`, без `@вид`) возвращает в
  эскиз родителя секущую линию. `delete_drawing_view` снимает все три.
- **Имя вида ≠ имя его фичи**: разрез показан как `Section View A-A`, а
  строки выбора аннотаций — `RD1@Drawing View3`. Соответствие — подфичи
  `DrSheet` (`ext._view_feature_names`); тулы принимают оба имени.

### Выбор рёбер в виде

- Основной путь — `IView.SelectEntity(ребро_детали, append)`: ребро по
  индексу `list_edges` детали или ближайшее к точке модели `x,y,z`.
  **Точка должна лежать на ребре**: точка «внутри» профиля выберет
  ближайшую окружность другого диаметра (Ø1050 вместо Ø830 — сверять
  значение в ответе).
- Координатный `SelectByID2("", "EDGE", x, y)` ненадёжен: мажет по коротким
  рёбрам, может вернуть сам вид. Не используется. Цена: точку крепления
  выноски на ребре выбирает SW, положение символа задаётся явно.

### Размеры

- `AddDimension2(x, y, 0)` по выбранным рёбрам: 1 круговое → диаметр,
  2 → расстояние; точка текста решает горизонталь/вертикаль. Размер
  ведомый (`RD1`); ГОСТ ставит скобки → `ShowParenthesis = False`.
- `IDimension.Value` — в единицах документа, брать `SystemValue` (м).
- Допуск: `SetValues2` в чертеже → False, **`SetValues(min, max)`**
  работает. Тип: BILAT 2, SYMMETRIC 4, FIT 7, FITWITHTOL 8;
  `SetFitValues(hole, shaft)`. Отклонения FITWITHTOL SW считает сам по
  ISO 286.
- Префикс/суффикс: `IDisplayDimension.SetText(1 | 2, …)`, Ø = `<MOD-DIAM>`.
  Справочный по ГОСТ — суффикс `*` + ТТ «* Размеры для справок».
- Точность — настройки документа: `swDetailingLinearDimPrecision` 24,
  `…LinearTolPrecision` 25, нули `swDetailingDimTrailingZero` 15 /
  `…TrailingZeroTolerance` 582 (2 = убрать), вид «x» у фаски
  `swDetailingChamferDimXStyle` 156. Фаска — `IDrawingDoc.AddChamferDim`
  (кромка фаски + соседнее ребро).

### Допуски формы (IGtol)

- `SetFrameSymbols2(Frame, GCS, …)`: **GCS — строка** (`VT_BSTR`) — имя из
  `lang\english\gtol.sym`: `<IGTOL-SRUN>`, `<IGTOL-TRUN>`, `<IGTOL-CYL>`,
  `<IGTOL-PARA>`… (`#GGTOL` — ГОСТ-набор). int 25 рисуется литералом «25».
  Значения — `SetFrameValues2(1, tol, '', datum, '', '')`.
- `GetFormat()`: 1 = старый (свежий `InsertGtol`, рабочий),
  `ConvertFormat()` → 2. **Сначала символ и значения, потом конвертация**:
  обратный порядок теряет значение, хэндл после конвертации протухает.
  Перечитывать через `IView.GetFirstGTOL()` / `GetNextGTOL()`.

### Шероховатость (ISFSymbol)

`InsertSurfaceFinishSymbol3(… MaxRoughness …)` кладёт текст в слот 5, а
какие слоты рисуются, решает `swDetailingSFSymbolStandard` (pref 629):

| pref 629 | стандарт | значение в слоте |
|---|---|---|
| 0 | ISO 1302:1992 (дефолт `gost.drwdot`) | **5** |
| 1 | ISO 1302:2002 | **8** |
| 2 | ISO 21920-1 | **8** |

`SetText(8, 'Ra 3,2')`; текст в нерисуемом слоте хранится и читается —
видно только на картинке. `add_surface_finish` выбирает слот сам. Символ
**без выноски** игнорирует LocX/LocY и встаёт в (0,0) —
`IAnnotation.SetPosition2` после вставки.

### Удаление аннотаций

`IAnnotation.Select3` → False для размеров и баз в виде. Работает
`SelectByID2('<имя>@<имя ФИЧИ вида>', тип)` + `DeleteSelection2`; типы
`DIMENSION`, `DATUMTAG`, `GTOL`, `SFSYMBOL`, `NOTE` (`delete_annotation`).

### Лист и штамп

- Формат: `SetupSheet6(name, 9 (A2), 12 (custom), 1, 5, firstAngle,
  "<…>\lang\english\sheetformat\a2 - gost_sh1_land.slddrt", 0.594, 0.42,
  "Default", True, 0,0,0,0,0,0)`. Стандарт оформления ГОСТ —
  `swDetailingDimensionStandard` 13 = 6.
- Штамп ГОСТ берёт `$PRP` из свойств **чертежа**
  (`CustomPropertyManager('').Add3(name, 30, value, 2)`): `Обозначение`,
  `Код_документа` (в формате стоит «СБ» — у детали очистить),
  `Наименование`, `Материал`, `Масса`. Системные `SW- Масштаб листа`,
  `SW- Текущий лист`, `SW- Количество листов` в русском формате не
  резолвятся — задать теми же свойствами.
- Подпись разреза: `INote.SetText('<VLLABEL>')` у первой заметки вида
  (иначе «SECTION A-A SCALE 1:5»).

## Известные открытые вопросы

- `GetTypeName2` иногда отдаёт нестандартные имена (`"ICE"` и для
  Boss-, и для Cut-Extrude — общее «фича из эскиза»). Не расшифровывать.
- Паз на цилиндре — см. § Sketch на цилиндрической грани.
- `IMessageFilter` для SW, ставшего модальным во время вызова, — не сделан.
