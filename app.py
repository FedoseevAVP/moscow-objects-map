import html
import math
from pathlib import Path

import folium
import pandas as pd
import requests
import streamlit as st
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium


# =========================================================
# ОСНОВНЫЕ НАСТРОЙКИ
# =========================================================

st.set_page_config(
    page_title="Карта обслуживаемых объектов",
    page_icon="🗺️",
    layout="wide",
)

# Excel должен лежать рядом с app.py
EXCEL_FILE = Path("objects.xlsx")

# Центр Москвы на случай, если он понадобится
MOSCOW_CENTER = [55.751244, 37.618423]

# Масштаб карты
DEFAULT_ZOOM = 10

# Столбцы, которые обязательно должны быть в Excel
REQUIRED_COLUMNS = [
    "ID",
    "Объект",
    "Адрес",
    "Тип",
    "Ответственный",
    "Телефон",
    "Статус",
    "Комментарий",
    "Широта",
    "Долгота",
]


# =========================================================
# СЛУЖЕБНЫЕ ФУНКЦИИ
# =========================================================

def clean_text(value):
    """
    Превращает значение Excel в нормальную строку.
    Пустые ячейки становятся пустой строкой.
    """
    if pd.isna(value):
        return ""

    return str(value).strip()


def clean_phone(value):
    """
    Исправляет проблему Excel, когда телефон может
    превратиться из 79999999999 в 79999999999.0.
    """
    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    text = str(value).strip()

    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]

    return text


def to_number(series):
    """
    Преобразует координаты в числа.

    Поддерживает:
    55.752
    и
    55,752
    """
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )


def haversine_km(lat1, lon1, lat2, lon2):
    """
    Вычисляет расстояние между двумя координатами
    по поверхности Земли.

    Результат — километры.
    """

    earth_radius = 6371.0088

    lat1_rad = math.radians(float(lat1))
    lat2_rad = math.radians(float(lat2))

    delta_lat = math.radians(float(lat2) - float(lat1))
    delta_lon = math.radians(float(lon2) - float(lon1))

    a = (
        math.sin(delta_lat / 2) ** 2
        + math.cos(lat1_rad)
        * math.cos(lat2_rad)
        * math.sin(delta_lon / 2) ** 2
    )

    return 2 * earth_radius * math.asin(math.sqrt(a))


def marker_color(status):
    """
    Определяет цвет объекта по статусу.
    """

    status = clean_text(status).lower()

    # Красный
    if any(
        word in status
        for word in ["авар", "проблем", "сроч"]
    ):
        return "red"

    # Оранжевый
    if any(
        word in status
        for word in ["нов", "запуск"]
    ):
        return "orange"

    # Серый
    if any(
        word in status
        for word in ["приост", "закрыт", "не действ"]
    ):
        return "gray"

    # Зеленый
    if "действ" in status:
        return "green"

    # Если статус непонятен
    return "blue"


# =========================================================
# ЧТЕНИЕ EXCEL
# =========================================================

@st.cache_data
def load_objects(file_mtime):
    """
    Загружает Excel.

    file_mtime нужен для того, чтобы Streamlit понял,
    что Excel был заменен новым.
    """

    # Само значение внутри функции не используется,
    # оно нужно только для обновления кэша.
    del file_mtime

    df = pd.read_excel(
        EXCEL_FILE,

        # Берем первый лист Excel
        sheet_name=0,

        engine="openpyxl",

        dtype={
            "ID": str,
            "Объект": str,
            "Адрес": str,
            "Тип": str,
            "Ответственный": str,
            "Телефон": str,
            "Статус": str,
            "Комментарий": str,
        },
    )

    # Убираем случайные пробелы из названий столбцов
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    # Проверяем наличие всех нужных колонок
    missing_columns = [
        column
        for column in REQUIRED_COLUMNS
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "В Excel отсутствуют обязательные столбцы: "
            + ", ".join(missing_columns)
        )

    # Чистим текстовые столбцы
    text_columns = [
        "ID",
        "Объект",
        "Адрес",
        "Тип",
        "Ответственный",
        "Статус",
        "Комментарий",
    ]

    for column in text_columns:

        df[column] = (
            df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    df["Телефон"] = df["Телефон"].apply(clean_phone)

    # Координаты превращаем в числа
    df["Широта"] = to_number(df["Широта"])
    df["Долгота"] = to_number(df["Долгота"])

    return df


# =========================================================
# ПОЛУЧЕНИЕ СТАНЦИЙ МЕТРО
# =========================================================

@st.cache_data(
    ttl=86400,
    show_spinner=False
)
def load_metro_stations():
    """
    Получает станции Московского метро
    из OpenStreetMap через Overpass API.

    Данные сохраняются в кэше на 24 часа.
    """

    query = """
    [out:json][timeout:25];

    (
      nwr
      ["railway"="station"]
      ["station"="subway"]
      (55.30,36.70,56.10,38.20);

      nwr
      ["railway"="station"]
      ["subway"="yes"]
      (55.30,36.70,56.10,38.20);

      nwr
      ["public_transport"="station"]
      ["subway"="yes"]
      (55.30,36.70,56.10,38.20);
    );

    out center tags;
    """

    # Если первый сервер временно недоступен,
    # попробуем второй.
    endpoints = [
        "https://overpass-api.de/api/interpreter",
        "https://overpass.kumi.systems/api/interpreter",
    ]

    last_error = None

    for endpoint in endpoints:

        try:

            response = requests.post(
                endpoint,
                data={
                    "data": query
                },
                timeout=35,
                headers={
                    "User-Agent":
                    "CompanyObjectsMap/1.0"
                },
            )

            response.raise_for_status()

            data = response.json()

            stations = []

            for element in data.get("elements", []):

                tags = element.get("tags", {})

                # Сначала русское название,
                # если его нет — обычное
                name = (
                    tags.get("name:ru")
                    or tags.get("name")
                )

                if not name:
                    continue

                lat = element.get("lat")
                lon = element.get("lon")

                # У way/relation координаты могут находиться в center
                if lat is None or lon is None:

                    center = element.get(
                        "center",
                        {}
                    )

                    lat = center.get("lat")
                    lon = center.get("lon")

                if lat is None or lon is None:
                    continue

                stations.append(
                    {
                        "Станция": str(name).strip(),
                        "Широта метро": float(lat),
                        "Долгота метро": float(lon),
                    }
                )

            metro = pd.DataFrame(stations)

            if metro.empty:
                raise RuntimeError(
                    "Список метро оказался пустым"
                )

            # Одна станция в OpenStreetMap
            # иногда встречается несколько раз.
            # Объединяем такие записи.
            metro["key"] = (
                metro["Станция"]
                .str.lower()
                .str.replace(
                    "ё",
                    "е",
                    regex=False
                )
            )

            metro = (
                metro
                .groupby(
                    "key",
                    as_index=False
                )
                .agg(
                    {
                        "Станция": "first",
                        "Широта метро": "mean",
                        "Долгота метро": "mean",
                    }
                )
                .drop(
                    columns="key"
                )
            )

            return metro

        except Exception as error:

            last_error = error

    raise RuntimeError(
        "Не удалось загрузить метро: "
        + str(last_error)
    )


# =========================================================
# ПОИСК БЛИЖАЙШЕГО МЕТРО
# =========================================================

def nearest_metro(
    object_lat,
    object_lon,
    metro_df
):
    """
    Находит ближайшую станцию метро
    к конкретному объекту.
    """

    if metro_df.empty:
        return "", None, None, None

    nearest_name = ""

    nearest_distance = float("inf")

    nearest_lat = None
    nearest_lon = None

    for _, station in metro_df.iterrows():

        station_lat = station["Широта метро"]
        station_lon = station["Долгота метро"]

        distance = haversine_km(
            object_lat,
            object_lon,
            station_lat,
            station_lon,
        )

        if distance < nearest_distance:

            nearest_distance = distance

            nearest_name = station["Станция"]

            nearest_lat = station_lat
            nearest_lon = station_lon

    return (
        nearest_name,
        nearest_distance,
        nearest_lat,
        nearest_lon,
    )


def add_nearest_metro(
    objects_df,
    metro_df
):
    """
    Добавляет к каждому объекту:

    - ближайшее метро
    - расстояние до метро
    - координаты метро
    """

    result = objects_df.copy()

    metro_names = []

    distances = []

    metro_lats = []

    metro_lons = []

    for _, row in result.iterrows():

        lat = row["Широта"]
        lon = row["Долгота"]

        if (
            pd.isna(lat)
            or pd.isna(lon)
            or metro_df.empty
        ):

            metro_names.append("")

            distances.append(None)

            metro_lats.append(None)

            metro_lons.append(None)

            continue

        (
            name,
            distance_km,
            station_lat,
            station_lon,
        ) = nearest_metro(
            float(lat),
            float(lon),
            metro_df,
        )

        metro_names.append(name)

        # Округляем примерно до 10 метров
        if distance_km is not None:

            distance_m = round(
                distance_km * 100
            ) * 10

        else:

            distance_m = None

        distances.append(
            distance_m
        )

        metro_lats.append(
            station_lat
        )

        metro_lons.append(
            station_lon
        )

    result[
        "Ближайшее метро"
    ] = metro_names

    result[
        "До метро, м"
    ] = distances

    result[
        "_metro_lat"
    ] = metro_lats

    result[
        "_metro_lon"
    ] = metro_lons

    return result


# =========================================================
# ЗАГОЛОВОК САЙТА
# =========================================================

st.title(
    "🗺️ Карта обслуживаемых объектов"
)

st.caption(
    "Интерактивная карта объектов компании "
    "и ближайших станций метро"
)


# =========================================================
# ПРОВЕРЯЕМ НАЛИЧИЕ EXCEL
# =========================================================

if not EXCEL_FILE.exists():

    st.error(
        "Не найден файл objects.xlsx. "
        "Файл objects.xlsx должен находиться "
        "рядом с app.py."
    )

    st.stop()


# =========================================================
# ЗАГРУЖАЕМ EXCEL
# =========================================================

try:

    objects = load_objects(
        EXCEL_FILE.stat().st_mtime
    )

except Exception as error:

    st.error(
        "Ошибка чтения objects.xlsx:\n\n"
        + str(error)
    )

    st.stop()


if objects.empty:

    st.warning(
        "В objects.xlsx пока нет объектов."
    )

    st.stop()


# =========================================================
# ПРОВЕРЯЕМ КООРДИНАТЫ
# =========================================================

invalid_coordinates = objects[
    objects["Широта"].isna()
    |
    objects["Долгота"].isna()
    |
    ~objects["Широта"].between(
        -90,
        90
    )
    |
    ~objects["Долгота"].between(
        -180,
        180
    )
]


# Берем только объекты с нормальными координатами
objects_valid = objects.drop(
    index=invalid_coordinates.index
).copy()


if objects_valid.empty:

    st.error(
        "Ни у одного объекта нет "
        "корректных координат.\n\n"
        "Проверьте столбцы "
        "«Широта» и «Долгота»."
    )

    st.stop()


# =========================================================
# ЗАГРУЖАЕМ МЕТРО
# =========================================================

metro_error = None

try:

    with st.spinner(
        "Получаем данные о метро..."
    ):

        metro_df = load_metro_stations()

except Exception as error:

    metro_df = pd.DataFrame(
        columns=[
            "Станция",
            "Широта метро",
            "Долгота метро",
        ]
    )

    metro_error = str(error)


# Добавляем ближайшее метро
objects_valid = add_nearest_metro(
    objects_valid,
    metro_df
)


# =========================================================
# БОКОВАЯ ПАНЕЛЬ
# =========================================================

st.sidebar.header(
    "Фильтры"
)


# ---------------------------------------------------------
# ПОИСК
# ---------------------------------------------------------

search_text = st.sidebar.text_input(
    "🔎 Поиск",
    placeholder=(
        "Объект, адрес, ответственный..."
    ),
).strip()


# ---------------------------------------------------------
# ТИП ОБЪЕКТА
# ---------------------------------------------------------

all_types = sorted(
    [
        value
        for value
        in objects_valid["Тип"].dropna().unique()
        if value
    ]
)

selected_types = st.sidebar.multiselect(
    "Тип",
    options=all_types,
    default=all_types,
)


# ---------------------------------------------------------
# СТАТУС
# ---------------------------------------------------------

all_statuses = sorted(
    [
        value
        for value
        in objects_valid["Статус"].dropna().unique()
        if value
    ]
)

selected_statuses = st.sidebar.multiselect(
    "Статус",
    options=all_statuses,
    default=all_statuses,
)


# ---------------------------------------------------------
# ОТВЕТСТВЕННЫЙ
# ---------------------------------------------------------

all_people = sorted(
    [
        value
        for value
        in objects_valid[
            "Ответственный"
        ]
        .dropna()
        .unique()
        if value
    ]
)

selected_people = st.sidebar.multiselect(
    "Ответственный",
    options=all_people,
    default=all_people,
)


# ---------------------------------------------------------
# ДОПОЛНИТЕЛЬНЫЕ НАСТРОЙКИ
# ---------------------------------------------------------

st.sidebar.divider()

show_metro = st.sidebar.checkbox(
    "Ⓜ Показать ближайшее метро",
    value=True,
)

show_lines = st.sidebar.checkbox(
    "Показать линию до метро",
    value=True,
    disabled=not show_metro,
)

show_table = st.sidebar.checkbox(
    "Показать таблицу под картой",
    value=False,
)


# =========================================================
# ПРИМЕНЯЕМ ФИЛЬТРЫ
# =========================================================

filtered = objects_valid.copy()


if all_types:

    filtered = filtered[
        filtered["Тип"].isin(
            selected_types
        )
    ]


if all_statuses:

    filtered = filtered[
        filtered["Статус"].isin(
            selected_statuses
        )
    ]


if all_people:

    filtered = filtered[
        filtered["Ответственный"].isin(
            selected_people
        )
    ]


# =========================================================
# ПОИСК ПО ТЕКСТУ
# =========================================================

if search_text:

    search_lower = search_text.lower()

    search_columns = [
        "ID",
        "Объект",
        "Адрес",
        "Тип",
        "Ответственный",
        "Телефон",
        "Статус",
        "Комментарий",
    ]

    mask = pd.Series(
        False,
        index=filtered.index
    )

    for column in search_columns:

        mask = (
            mask
            |
            filtered[column]
            .fillna("")
            .astype(str)
            .str.lower()
            .str.contains(
                search_lower,
                regex=False,
            )
        )

    filtered = filtered[mask]


# =========================================================
# СТАТИСТИКА СВЕРХУ
# =========================================================

col1, col2, col3 = st.columns(3)


col1.metric(
    "Всего объектов",
    len(objects_valid)
)


col2.metric(
    "Показано на карте",
    len(filtered)
)


if metro_df.empty:

    col3.metric(
        "Станций метро",
        "—"
    )

else:

    col3.metric(
        "Станций метро",
        len(metro_df)
    )


# =========================================================
# ПРЕДУПРЕЖДЕНИЯ
# =========================================================

if not invalid_coordinates.empty:

    st.warning(
        f"⚠️ Не показано объектов: "
        f"{len(invalid_coordinates)}. "
        "Причина: отсутствуют или неправильно "
        "заполнены Широта/Долгота."
    )


if metro_error:

    st.warning(
        "⚠️ Сейчас не удалось получить "
        "данные метро. "
        "Карта объектов продолжает работать, "
        "но метро временно не показывается."
    )


# =========================================================
# ЕСЛИ ФИЛЬТР НИЧЕГО НЕ НАШЕЛ
# =========================================================

if filtered.empty:

    st.info(
        "По выбранным фильтрам "
        "объектов не найдено."
    )

    st.stop()


# =========================================================
# СОЗДАЕМ КАРТУ
# =========================================================

map_center = [
    float(
        filtered["Широта"].mean()
    ),
    float(
        filtered["Долгота"].mean()
    ),
]


m = folium.Map(
    location=map_center,
    zoom_start=DEFAULT_ZOOM,
    tiles="OpenStreetMap",
    control_scale=True,
)


# Группа объектов
objects_cluster = MarkerCluster(
    name="Объекты"
).add_to(m)


# Сюда записываем уже нанесенные станции метро,
# чтобы одна станция не рисовалась 10 раз.
metro_markers_added = set()


# =========================================================
# ДОБАВЛЯЕМ ОБЪЕКТЫ
# =========================================================

for _, row in filtered.iterrows():

    lat = float(
        row["Широта"]
    )

    lon = float(
        row["Долгота"]
    )


    # -----------------------------------------------------
    # ТЕКСТ ИЗ EXCEL
    # -----------------------------------------------------

    object_name = html.escape(
        clean_text(
            row["Объект"]
        )
        or
        "Без названия"
    )

    address = html.escape(
        clean_text(
            row["Адрес"]
        )
    )

    object_type = html.escape(
        clean_text(
            row["Тип"]
        )
    )

    responsible = html.escape(
        clean_text(
            row["Ответственный"]
        )
    )

    phone = html.escape(
        clean_text(
            row["Телефон"]
        )
    )

    status = html.escape(
        clean_text(
            row["Статус"]
        )
    )

    comment = html.escape(
        clean_text(
            row["Комментарий"]
        )
    )

    object_id = html.escape(
        clean_text(
            row["ID"]
        )
    )


    # -----------------------------------------------------
    # МЕТРО
    # -----------------------------------------------------

    metro_name = html.escape(
        clean_text(
            row.get(
                "Ближайшее метро",
                ""
            )
        )
    )

    metro_distance = row.get(
        "До метро, м"
    )


    if (
        metro_name
        and
        pd.notna(
            metro_distance
        )
    ):

        metro_text = (
            f"{metro_name} — "
            f"примерно "
            f"{int(metro_distance)} м"
        )

    else:

        metro_text = (
            "Не определено"
        )


    # -----------------------------------------------------
    # ССЫЛКИ НА КАРТЫ
    # -----------------------------------------------------

    google_route = (
        "https://www.google.com/maps/dir/"
        f"?api=1&destination={lat},{lon}"
    )

    yandex_map = (
        "https://yandex.ru/maps/"
        f"?pt={lon},{lat}"
        "&z=16&l=map"
    )


    # -----------------------------------------------------
    # ТЕЛЕФОН
    # -----------------------------------------------------

    if phone:

        phone_digits = "".join(
            symbol
            for symbol in phone
            if (
                symbol.isdigit()
                or symbol == "+"
            )
        )

        phone_html = (
            f'<a href="tel:{phone_digits}">'
            f'{phone}'
            f'</a>'
        )

    else:

        phone_html = "—"


    # -----------------------------------------------------
    # КАРТОЧКА ОБЪЕКТА
    # -----------------------------------------------------

    popup_html = f"""
    <div style="
        width: 320px;
        font-family: Arial, sans-serif;
        font-size: 14px;
        line-height: 1.5;
    ">

        <h3 style="
            margin-top: 0;
            margin-bottom: 10px;
        ">
            {object_name}
        </h3>

        <b>ID:</b>
        {object_id}
        <br>

        <b>Адрес:</b>
        {address or "—"}
        <br>

        <b>Тип:</b>
        {object_type or "—"}
        <br>

        <b>Ответственный:</b>
        {responsible or "—"}
        <br>

        <b>Телефон:</b>
        {phone_html}
        <br>

        <b>Статус:</b>
        {status or "—"}
        <br>

        <b>Комментарий:</b>
        {comment or "—"}
        <br>

        <hr>

        <b>Ⓜ Ближайшее метро:</b>
        <br>

        {metro_text}

        <br><br>

        <a
            href="{yandex_map}"
            target="_blank"
        >
            Открыть в Яндекс Картах
        </a>

        <br>

        <a
            href="{google_route}"
            target="_blank"
        >
            Построить маршрут
        </a>

    </div>
    """


    # -----------------------------------------------------
    # МАРКЕР ОБЪЕКТА
    # -----------------------------------------------------

    folium.Marker(
        location=[
            lat,
            lon
        ],

        tooltip=(
            clean_text(
                row["Объект"]
            )
            or
            "Объект"
        ),

        popup=folium.Popup(
            popup_html,
            max_width=380,
        ),

        icon=folium.Icon(
            color=marker_color(
                row["Статус"]
            ),

            icon="wrench",

            prefix="fa",
        ),

    ).add_to(
        objects_cluster
    )


    # =====================================================
    # МЕТРО
    # =====================================================

    if (
        show_metro
        and
        metro_name
        and
        pd.notna(
            row.get(
                "_metro_lat"
            )
        )
    ):

        metro_lat = float(
            row["_metro_lat"]
        )

        metro_lon = float(
            row["_metro_lon"]
        )


        metro_key = (
            clean_text(
                row[
                    "Ближайшее метро"
                ]
            ),
            round(
                metro_lat,
                5
            ),
            round(
                metro_lon,
                5
            ),
        )


        # Добавляем маркер станции
        # только один раз
        if (
            metro_key
            not in metro_markers_added
        ):

            folium.CircleMarker(

                location=[
                    metro_lat,
                    metro_lon
                ],

                radius=7,

                color="#1565C0",

                fill=True,

                fill_color="#1565C0",

                fill_opacity=0.9,

                tooltip=(
                    "Ⓜ "
                    +
                    clean_text(
                        row[
                            "Ближайшее метро"
                        ]
                    )
                ),

                popup=folium.Popup(
                    "<b>Ⓜ "
                    +
                    html.escape(
                        clean_text(
                            row[
                                "Ближайшее метро"
                            ]
                        )
                    )
                    +
                    "</b>",
                    max_width=260,
                ),

            ).add_to(m)


            metro_markers_added.add(
                metro_key
            )


        # -------------------------------------------------
        # ЛИНИЯ ОТ ОБЪЕКТА ДО МЕТРО
        # -------------------------------------------------

        if (
            show_lines
            and
            pd.notna(
                metro_distance
            )
        ):

            folium.PolyLine(

                locations=[
                    [
                        lat,
                        lon
                    ],

                    [
                        metro_lat,
                        metro_lon
                    ]
                ],

                weight=2,

                opacity=0.55,

                dash_array="5, 8",

                tooltip=(
                    "До метро примерно "
                    +
                    str(
                        int(
                            metro_distance
                        )
                    )
                    +
                    " м"
                ),

            ).add_to(m)


# =========================================================
# УПРАВЛЕНИЕ СЛОЯМИ
# =========================================================

folium.LayerControl(
    collapsed=True
).add_to(m)


# =========================================================
# ПОКАЗЫВАЕМ КАРТУ
# =========================================================

st_folium(
    m,
    height=680,
    use_container_width=True,

    # Не нужно перезапускать приложение
    # при каждом клике по карте.
    returned_objects=[],
)


# =========================================================
# ТАБЛИЦА
# =========================================================

if show_table:

    st.subheader(
        "Список объектов"
    )

    table_columns = [
        "ID",
        "Объект",
        "Адрес",
        "Тип",
        "Ответственный",
        "Телефон",
        "Статус",
        "Комментарий",
        "Ближайшее метро",
        "До метро, м",
    ]

    st.dataframe(
        filtered[
            table_columns
        ]
        .reset_index(
            drop=True
        ),

        use_container_width=True,

        hide_index=True,
    )


# =========================================================
# ПОКАЗЫВАЕМ ОБЪЕКТЫ С ОШИБКОЙ КООРДИНАТ
# =========================================================

if not invalid_coordinates.empty:

    with st.expander(
        "Объекты с ошибкой координат"
    ):

        st.dataframe(

            invalid_coordinates[
                [
                    "ID",
                    "Объект",
                    "Адрес",
                    "Широта",
                    "Долгота",
                ]
            ]
            .reset_index(
                drop=True
            ),

            use_container_width=True,

            hide_index=True,
        )
