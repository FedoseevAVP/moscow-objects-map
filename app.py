from __future__ import annotations

import html
import math
import re
from pathlib import Path
from typing import Any
from urllib.parse import quote

import folium
import pandas as pd
import requests
import streamlit as st
from branca.element import MacroElement, Template
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium


# =========================================================
# НАСТРОЙКИ
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "objects.xlsx"
LOGO_FILE = BASE_DIR / "assets" / "logo_linkor.png"

DEFAULT_CENTER = [55.751244, 37.618423]
DEFAULT_ZOOM = 10

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
DGIS_GEOCODER_URL = "https://catalog.api.2gis.com/3.0/items/geocode"
OSRM_TABLE_URL = "https://router.project-osrm.org/table/v1/driving"
OSRM_ROUTE_URL = "https://router.project-osrm.org/route/v1/driving"
DGIS_ROUTING_URL = "https://routing.api.2gis.com/routing/7.0.0/global"
DGIS_TRANSIT_URL = "https://routing.api.2gis.com/public_transport/2.0"
OVERPASS_URLS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
]
HH_METRO_URL = "https://api.hh.ru/metro/1"

USER_AGENT = "LinkorObjectsMap/2.0 (service@lin-cor.ru)"

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

BRAND_NAVY = "#171A63"
BRAND_BLUE = "#2E4F9A"
BRAND_RED = "#D64A4A"
BRAND_BG = "#F4F6FA"
BRAND_LINE = "#DDE2ED"


st.set_page_config(
    page_title="Линкор — карта объектов",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# ОФОРМЛЕНИЕ
# =========================================================

st.markdown(
    f"""
    <style>
        :root {{
            --linkor-navy: {BRAND_NAVY};
            --linkor-blue: {BRAND_BLUE};
            --linkor-red: {BRAND_RED};
            --linkor-bg: {BRAND_BG};
            --linkor-line: {BRAND_LINE};
        }}

        .stApp {{
            background: var(--linkor-bg);
            color: #171A2B;
        }}

        [data-testid="stHeader"] {{
            background: rgba(244, 246, 250, 0.92);
        }}

        [data-testid="stSidebar"] {{
            background: #FFFFFF;
            border-right: 1px solid var(--linkor-line);
        }}

        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3 {{
            color: var(--linkor-navy);
        }}

        .block-container {{
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            max-width: 1600px;
        }}

        .linkor-header {{
            background: #FFFFFF;
            border: 1px solid var(--linkor-line);
            border-left: 5px solid var(--linkor-red);
            padding: 18px 22px;
            margin-bottom: 18px;
            box-shadow: 0 8px 24px rgba(23, 26, 99, 0.06);
        }}

        .linkor-kicker {{
            color: var(--linkor-red);
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            margin-bottom: 4px;
        }}

        .linkor-title {{
            color: var(--linkor-navy);
            font-size: clamp(1.55rem, 2.6vw, 2.35rem);
            font-weight: 750;
            line-height: 1.1;
            margin: 0;
        }}

        .linkor-subtitle {{
            color: #62677D;
            margin-top: 7px;
            font-size: 0.96rem;
        }}

        .logistics-card {{
            background: #FFFFFF;
            border: 1px solid var(--linkor-line);
            border-top: 4px solid var(--linkor-blue);
            padding: 18px 20px;
            box-shadow: 0 8px 22px rgba(23, 26, 99, 0.06);
            margin: 8px 0 14px;
        }}

        .verdict-good {{ border-top-color: #208A61; }}
        .verdict-medium {{ border-top-color: #D08B24; }}
        .verdict-review {{ border-top-color: var(--linkor-red); }}

        .verdict-label {{
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #73798D;
        }}

        .verdict-title {{
            color: var(--linkor-navy);
            font-size: 1.35rem;
            font-weight: 750;
            margin: 4px 0 8px;
        }}

        .metric-note {{ color: #666D82; font-size: 0.9rem; }}

        div[data-testid="stMetric"] {{
            background: #FFFFFF;
            border: 1px solid var(--linkor-line);
            padding: 12px 14px;
        }}

        div[data-testid="stMetric"] label {{ color: #62677D; }}
        div[data-testid="stMetric"] [data-testid="stMetricValue"] {{ color: var(--linkor-navy); }}

        .stButton > button,
        .stFormSubmitButton > button {{
            background: var(--linkor-navy);
            color: #FFFFFF;
            border: 1px solid var(--linkor-navy);
            border-radius: 3px;
            font-weight: 700;
            min-height: 42px;
        }}

        .stButton > button:hover,
        .stFormSubmitButton > button:hover {{
            background: var(--linkor-blue);
            color: #FFFFFF;
            border-color: var(--linkor-blue);
        }}

        div[role="radiogroup"] {{
            background: #FFFFFF;
            border: 1px solid var(--linkor-line);
            padding: 6px 10px;
        }}

        a {{ color: var(--linkor-blue); }}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =========================================================


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def get_2gis_api_key() -> str:
    """Ключ хранится в Streamlit Secrets и никогда не записывается в Excel."""
    try:
        return clean_text(st.secrets.get("DGIS_API_KEY", ""))
    except Exception:
        return ""


class PlainLeafletAttribution(MacroElement):
    """Keep map and tile-provider attribution, without Leaflet's default flag."""

    _template = Template(
        """
        {% macro script(this, kwargs) %}
        if ({{ this._parent.get_name() }}.attributionControl) {
            {{ this._parent.get_name() }}.attributionControl.setPrefix(
                '<a href="https://leafletjs.com" target="_blank" rel="noopener noreferrer">Leaflet</a>'
            );
        }
        {% endmacro %}
        """
    )

    def __init__(self):
        super().__init__()
        self._name = "PlainLeafletAttribution"


def normalize_name(value: Any) -> str:
    return (
        clean_text(value)
        .lower()
        .replace("ё", "е")
        .replace("станция метро", "")
        .replace("метро", "")
        .strip(" «»\"'")
    )


def safe_color(value: str, fallback: str = "#7B8092") -> str:
    value = clean_text(value)
    if re.fullmatch(r"#[0-9A-Fa-f]{6}", value):
        return value
    if re.fullmatch(r"[0-9A-Fa-f]{6}", value):
        return "#" + value
    return fallback


def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius = 6371.0088
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)
    a = (
        math.sin(delta_phi / 2) ** 2
        + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2) ** 2
    )
    return radius * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def marker_color(status: Any) -> str:
    value = normalize_name(status)
    if any(word in value for word in ["авар", "проблем", "просроч"]):
        return "red"
    if any(word in value for word in ["нов", "соглас", "переговор"]):
        return "orange"
    if any(word in value for word in ["приост", "закрыт", "архив"]):
        return "gray"
    if any(word in value for word in ["действ", "обслуж", "актив"]):
        return "green"
    return "blue"


def service_object_mask(frame: pd.DataFrame) -> pd.Series:
    status = frame["Статус"].fillna("").astype(str).map(normalize_name)
    excluded = status.str.contains("закрыт|архив|отказ", regex=True)
    return ~excluded


# =========================================================
# EXCEL
# =========================================================


@st.cache_data(show_spinner=False)
def load_objects(file_mtime: float) -> pd.DataFrame:
    del file_mtime
    frame = pd.read_excel(EXCEL_FILE, engine="openpyxl")
    frame.columns = [str(column).strip() for column in frame.columns]

    missing = [column for column in REQUIRED_COLUMNS if column not in frame.columns]
    if missing:
        raise ValueError("В Excel не хватает столбцов: " + ", ".join(missing))

    frame = frame[REQUIRED_COLUMNS].copy()
    for column in ["Широта", "Долгота"]:
        frame[column] = pd.to_numeric(frame[column], errors="coerce")

    for column in [
        "ID",
        "Объект",
        "Адрес",
        "Тип",
        "Ответственный",
        "Телефон",
        "Статус",
        "Комментарий",
    ]:
        frame[column] = frame[column].fillna("").astype(str).str.strip()

    return frame


# =========================================================
# ГЕОКОДИРОВАНИЕ И МАРШРУТИЗАЦИЯ
# =========================================================


@st.cache_data(ttl=86400, show_spinner=False)
def geocode_address(address: str, dgis_api_key: str = "") -> dict[str, Any] | None:
    query = address.strip()
    if not query:
        return None

    # Основной источник — 2ГИС. Он ищет именно здания и возвращает
    # актуальную карточку объекта, назначение и точные координаты.
    if dgis_api_key:
        try:
            response = requests.get(
                DGIS_GEOCODER_URL,
                params={
                    "q": query,
                    "fields": "items.point,items.geometry.centroid,items.address",
                    "locale": "ru_RU",
                    "page_size": 5,
                    "key": dgis_api_key,
                },
                headers={"User-Agent": USER_AGENT},
                timeout=18,
            )
            response.raise_for_status()
            payload = response.json()
            items = payload.get("result", {}).get("items", [])
            with_coordinates = [item for item in items if item.get("point")]
            building = next(
                (item for item in with_coordinates if item.get("type") == "building"),
                with_coordinates[0] if with_coordinates else None,
            )
            if building:
                point = building["point"]
                return {
                    "lat": float(point["lat"]),
                    "lon": float(point["lon"]),
                    "display_name": building.get("full_name") or building.get("name") or address,
                    "provider": "2ГИС",
                    "building_id": clean_text(building.get("id")),
                    "building_type": clean_text(building.get("purpose_name")),
                }
        except Exception:
            # При временной ошибке 2ГИС сайт продолжит работу через резервный поиск.
            pass

    variants = [query]
    if "москв" not in query.lower():
        variants.append(query + ", Москва")

    for variant in variants:
        response = requests.get(
            NOMINATIM_URL,
            params={
                "q": variant,
                "format": "jsonv2",
                "limit": 1,
                "addressdetails": 1,
                "countrycodes": "ru",
            },
            headers={"User-Agent": USER_AGENT, "Accept-Language": "ru"},
            timeout=15,
        )
        response.raise_for_status()
        items = response.json()
        if items:
            item = items[0]
            return {
                "lat": float(item["lat"]),
                "lon": float(item["lon"]),
                "display_name": item.get("display_name", address),
                "provider": "OpenStreetMap — резерв",
                "building_id": "",
                "building_type": "",
            }
    return None


@st.cache_data(ttl=1800, show_spinner=False)
def road_table(
    origin_lat: float,
    origin_lon: float,
    destinations: tuple[tuple[float, float], ...],
) -> list[dict[str, float | None]]:
    if not destinations:
        return []

    coordinates = [f"{origin_lon:.6f},{origin_lat:.6f}"]
    coordinates.extend(f"{lon:.6f},{lat:.6f}" for lat, lon in destinations)
    destination_indexes = ";".join(str(i) for i in range(1, len(coordinates)))
    url = OSRM_TABLE_URL + "/" + ";".join(coordinates)

    response = requests.get(
        url,
        params={
            "sources": "0",
            "destinations": destination_indexes,
            "annotations": "distance,duration",
        },
        headers={"User-Agent": USER_AGENT},
        timeout=18,
    )
    response.raise_for_status()
    payload = response.json()
    distances = payload.get("distances", [[]])[0]
    durations = payload.get("durations", [[]])[0]

    result = []
    for distance, duration in zip(distances, durations):
        result.append(
            {
                "road_km": None if distance is None else float(distance) / 1000,
                "minutes": None if duration is None else float(duration) / 60,
            }
        )
    return result


@st.cache_data(ttl=1800, show_spinner=False)
def road_route(
    origin_lat: float,
    origin_lon: float,
    destination_lat: float,
    destination_lon: float,
) -> list[list[float]]:
    url = (
        OSRM_ROUTE_URL
        + f"/{origin_lon:.6f},{origin_lat:.6f};"
        + f"{destination_lon:.6f},{destination_lat:.6f}"
    )
    response = requests.get(
        url,
        params={"overview": "full", "geometries": "geojson"},
        headers={"User-Agent": USER_AGENT},
        timeout=18,
    )
    response.raise_for_status()
    coordinates = response.json()["routes"][0]["geometry"]["coordinates"]
    return [[float(lat), float(lon)] for lon, lat in coordinates]


def wkt_line(value: str) -> list[list[float]]:
    match = re.fullmatch(r"LINESTRING\s*\((.+)\)", value.strip(), re.IGNORECASE)
    if not match:
        return []
    try:
        points = []
        for pair in match.group(1).split(","):
            lon, lat = map(float, pair.strip().split()[:2])
            points.append([lat, lon])
        return points if len(points) >= 2 else []
    except (ValueError, IndexError):
        return []


@st.cache_data(ttl=900, show_spinner=False)
def dgis_route(
    origin_lat: float,
    origin_lon: float,
    destination_lat: float,
    destination_lon: float,
    travel_mode: str,
    api_key: str,
) -> dict[str, Any] | None:
    if not api_key:
        return None

    if travel_mode == "Общественный транспорт":
        url = DGIS_TRANSIT_URL
        body = {
            "source": {"point": {"lat": origin_lat, "lon": origin_lon}},
            "target": {"point": {"lat": destination_lat, "lon": destination_lon}},
            "transport": ["pedestrian", "metro", "light_metro", "mcc", "mcd",
                          "suburban_train", "tram", "bus", "trolleybus", "shuttle_bus"],
            "locale": "ru",
        }
    else:
        url = DGIS_ROUTING_URL
        point_type = "walking" if travel_mode == "Пешком" else "stop"
        body = {
            "points": [
                {"type": point_type, "lon": origin_lon, "lat": origin_lat},
                {"type": point_type, "lon": destination_lon, "lat": destination_lat},
            ],
            "transport": "walking" if travel_mode == "Пешком" else "driving",
            "route_mode": "fastest",
            "output": "detailed",
            "locale": "ru",
        }

    response = requests.post(url, params={"key": api_key}, json=body, timeout=18)
    response.raise_for_status()
    payload = response.json()
    variants = payload if isinstance(payload, list) else payload.get("result", [])
    if not isinstance(variants, list):
        return None
    variants = [
        item for item in variants
        if isinstance(item, dict) and item.get("total_duration") is not None
        and (travel_mode != "Общественный транспорт" or not item.get("pedestrian"))
    ]
    if not variants:
        return None
    route = min(variants, key=lambda item: item["total_duration"])
    selections = []
    if travel_mode == "Общественный транспорт":
        for movement in route.get("movements", []):
            alternatives = movement.get("alternatives") or []
            if alternatives:
                selections.extend(part.get("selection", "") for part in alternatives[0].get("geometry", []))
    else:
        for end in ("begin_pedestrian_path", "end_pedestrian_path"):
            selections.append((route.get(end) or {}).get("geometry", {}).get("selection", ""))
        for maneuver in route.get("maneuvers", []):
            selections.extend(part.get("selection", "") for part in
                              (maneuver.get("outcoming_path") or {}).get("geometry", []))
    return {
        "minutes": float(route["total_duration"]) / 60,
        "road_km": float(route["total_distance"]) / 1000,
        "segments": [line for selection in selections if (line := wkt_line(selection))],
        "provider": "2ГИС",
        "transfers": route.get("transfer_count") if travel_mode == "Общественный транспорт" else None,
    }


def evaluate_logistics(
    lat: float,
    lon: float,
    objects: pd.DataFrame,
    good_minutes: int,
    acceptable_minutes: int,
    density_radius_km: int,
    travel_mode: str,
    dgis_api_key: str,
) -> dict[str, Any]:
    candidates = objects[service_object_mask(objects)].copy()
    if candidates.empty:
        candidates = objects.copy()

    candidates["air_km"] = candidates.apply(
        lambda row: haversine_km(lat, lon, float(row["Широта"]), float(row["Долгота"])),
        axis=1,
    )
    candidates = candidates.sort_values("air_km")
    nearby_count = int((candidates["air_km"] <= density_radius_km).sum())
    route_candidates = candidates.head(5 if dgis_api_key else 8).copy()
    nearest = route_candidates.iloc[0]
    selected_route = None
    if dgis_api_key:
        routes = []
        for _, candidate in route_candidates.iterrows():
            try:
                route = dgis_route(
                    lat, lon, float(candidate["Широта"]), float(candidate["Долгота"]),
                    travel_mode, dgis_api_key,
                )
                if route:
                    routes.append((candidate, route))
            except (requests.RequestException, ValueError, KeyError, TypeError):
                continue
        if routes:
            nearest, selected_route = min(routes, key=lambda item: item[1]["minutes"])

    if selected_route is None and travel_mode == "На машине":
        try:
            destinations = tuple(
                (float(row["Широта"]), float(row["Долгота"]))
                for _, row in route_candidates.iterrows()
            )
            routes = road_table(lat, lon, destinations)
            available = [(candidate, route) for (_, candidate), route in zip(
                route_candidates.iterrows(), routes
            ) if route["minutes"] is not None]
            if available:
                nearest, selected_route = min(available, key=lambda item: item[1]["minutes"])
                selected_route = {**selected_route, "provider": "OSRM", "segments": []}
                try:
                    selected_route["segments"] = [road_route(
                        lat, lon, float(nearest["Широта"]), float(nearest["Долгота"])
                    )]
                except (requests.RequestException, ValueError, KeyError, IndexError):
                    pass
        except (requests.RequestException, ValueError, KeyError, IndexError):
            pass

    routing_ok = selected_route is not None
    minutes = selected_route["minutes"] if routing_ok else None
    road_km = selected_route["road_km"] if routing_ok else None
    air_km = float(nearest["air_km"])

    if routing_ok:
        if minutes <= good_minutes or nearby_count >= 3:
            level = "good"
            title = "Логистически удобно"
            reason = "Адрес близко к действующему маршруту или находится в плотном кластере объектов."
        elif minutes <= acceptable_minutes or nearby_count >= 2:
            level = "medium"
            title = "Можно рассматривать"
            reason = "Выезд находится в допустимом диапазоне, но график инженера стоит проверить."
        else:
            level = "review"
            title = "Нужна отдельная оценка"
            reason = "До ближайшего объекта далеко; желательно объединить выезд с другими работами."
    else:
        level, title = "review", "Маршрут недоступен"
        reason = (
            "Для выбранного способа передвижения нет маршрута. Проверьте доступ к Routing API 2ГИС "
            "для этого ключа либо выберите другой способ. Показано лишь расстояние по прямой."
        )

    return {
        "level": level,
        "title": title,
        "reason": reason,
        "nearest_name": clean_text(nearest["Объект"]) or "Без названия",
        "nearest_address": clean_text(nearest["Адрес"]),
        "nearest_lat": float(nearest["Широта"]),
        "nearest_lon": float(nearest["Долгота"]),
        "air_km": air_km,
        "road_km": road_km,
        "minutes": minutes,
        "nearby_count": nearby_count,
        "routing_ok": routing_ok,
        "travel_mode": travel_mode,
        "route_provider": selected_route["provider"] if routing_ok else "",
        "route_segments": selected_route["segments"] if routing_ok else [],
        "transfers": selected_route.get("transfers") if routing_ok else None,
    }


# =========================================================
# МЕТРО И ЛИНИИ МЕТРО
# =========================================================


@st.cache_data(ttl=86400, show_spinner=False)
def load_metro_network() -> dict[str, Any]:
    # Основной источник: открытый справочник метро HeadHunter.
    # Он возвращает станции в порядке следования по каждой линии,
    # поэтому из него можно быстро построить понятную схему.
    try:
        response = requests.get(
            HH_METRO_URL,
            headers={
                "User-Agent": USER_AGENT,
                "HH-User-Agent": USER_AGENT,
                "Accept-Language": "ru",
            },
            timeout=18,
        )
        response.raise_for_status()
        payload = response.json()
        station_rows: list[dict[str, Any]] = []
        line_segments: list[dict[str, Any]] = []

        for line in payload.get("lines", []):
            line_name = clean_text(line.get("name")) or "Линия метро"
            line_color = safe_color(clean_text(line.get("hex_color")))
            points: list[list[float]] = []

            for station in line.get("stations", []):
                lat = station.get("lat")
                lon = station.get("lng")
                name = clean_text(station.get("name"))
                if not name or lat is None or lon is None:
                    continue
                lat = float(lat)
                lon = float(lon)
                points.append([lat, lon])
                station_rows.append(
                    {
                        "Станция": name,
                        "Широта метро": lat,
                        "Долгота метро": lon,
                        "Линия": line_name,
                        "Цвет": line_color,
                    }
                )

            if len(points) >= 2:
                line_segments.append(
                    {"line": line_name, "color": line_color, "points": points}
                )

        if station_rows:
            stations = pd.DataFrame(station_rows)
            stations["key"] = stations["Станция"].map(normalize_name)
            stations = (
                stations.groupby("key", as_index=False)
                .agg(
                    {
                        "Станция": "first",
                        "Широта метро": "mean",
                        "Долгота метро": "mean",
                        "Линия": lambda values: " / ".join(sorted(set(values))),
                        "Цвет": "first",
                    }
                )
                .drop(columns="key")
            )
            return {"stations": stations, "segments": line_segments}
    except Exception:
        # Если основной справочник недоступен, ниже используется OpenStreetMap.
        pass

    query = """
    [out:json][timeout:45];
    area["name"="Москва"]["boundary"="administrative"]->.moscow;
    (
      relation["route"="subway"](area.moscow);
      relation["route"="light_rail"](area.moscow);
    );
    out body;
    >;
    out body geom;
    """

    last_error: Exception | None = None
    for endpoint in OVERPASS_URLS:
        try:
            response = requests.post(
                endpoint,
                data={"data": query},
                headers={"User-Agent": USER_AGENT},
                timeout=24,
            )
            response.raise_for_status()
            elements = response.json().get("elements", [])
            nodes = {item["id"]: item for item in elements if item.get("type") == "node"}
            ways = {item["id"]: item for item in elements if item.get("type") == "way"}
            relations = [item for item in elements if item.get("type") == "relation"]

            station_rows: list[dict[str, Any]] = []
            line_segments: list[dict[str, Any]] = []

            for relation in relations:
                tags = relation.get("tags", {})
                line_name = tags.get("name") or tags.get("ref") or "Линия метро"
                line_color = safe_color(tags.get("colour", ""))

                for member in relation.get("members", []):
                    role = clean_text(member.get("role")).lower()
                    ref = member.get("ref")

                    if member.get("type") == "node" and ("stop" in role or "platform" in role):
                        node = nodes.get(ref, {})
                        node_tags = node.get("tags", {})
                        name = node_tags.get("name") or node_tags.get("name:ru")
                        if name and node.get("lat") is not None and node.get("lon") is not None:
                            station_rows.append(
                                {
                                    "Станция": name,
                                    "Широта метро": float(node["lat"]),
                                    "Долгота метро": float(node["lon"]),
                                    "Линия": line_name,
                                    "Цвет": line_color,
                                }
                            )

                    if member.get("type") == "way" and not role.startswith("platform"):
                        way = ways.get(ref, {})
                        geometry = way.get("geometry") or member.get("geometry") or []
                        points = [
                            [float(point["lat"]), float(point["lon"])]
                            for point in geometry
                            if point.get("lat") is not None and point.get("lon") is not None
                        ]
                        if len(points) >= 2:
                            line_segments.append(
                                {"line": line_name, "color": line_color, "points": points}
                            )

            if not station_rows:
                raise RuntimeError("Станции метро не найдены")

            stations = pd.DataFrame(station_rows)
            stations["key"] = stations["Станция"].map(normalize_name)
            stations = (
                stations.groupby("key", as_index=False)
                .agg(
                    {
                        "Станция": "first",
                        "Широта метро": "mean",
                        "Долгота метро": "mean",
                        "Линия": lambda values: " / ".join(sorted(set(values))),
                        "Цвет": "first",
                    }
                )
                .drop(columns="key")
            )
            return {"stations": stations, "segments": line_segments}
        except Exception as error:
            last_error = error

    raise RuntimeError("Не удалось загрузить схему метро: " + str(last_error))


def nearest_metro(lat: float, lon: float, metro_df: pd.DataFrame) -> tuple[str, float, float, float]:
    best_name = ""
    best_distance = float("inf")
    best_lat = 0.0
    best_lon = 0.0

    for _, station in metro_df.iterrows():
        distance = haversine_km(
            lat,
            lon,
            float(station["Широта метро"]),
            float(station["Долгота метро"]),
        )
        if distance < best_distance:
            best_name = clean_text(station["Станция"])
            best_distance = distance
            best_lat = float(station["Широта метро"])
            best_lon = float(station["Долгота метро"])

    return best_name, best_distance, best_lat, best_lon


def is_mcd_3_or_4(line_name: Any) -> bool:
    """Определяет две линии, которые сильнее всего растягивают схему."""
    compact = re.sub(r"[\s–—_]", "", normalize_name(line_name))
    return "мцд-3" in compact or "мцд3" in compact or "мцд-4" in compact or "мцд4" in compact


def filter_metro_network(
    metro_network: dict[str, Any],
    include_mcd_3_4: bool,
) -> dict[str, Any]:
    """Скрывает МЦД-3/4, сохраняя пересадочные станции обычного метро."""
    if include_mcd_3_4:
        return metro_network

    stations = metro_network["stations"].copy()

    def station_has_regular_line(value: Any) -> bool:
        line_names = [part.strip() for part in clean_text(value).split(" / ") if part.strip()]
        return any(not is_mcd_3_or_4(line_name) for line_name in line_names)

    stations = stations[stations["Линия"].map(station_has_regular_line)].copy()
    segments = [
        segment
        for segment in metro_network.get("segments", [])
        if not is_mcd_3_or_4(segment.get("line", ""))
    ]
    return {"stations": stations, "segments": segments}


@st.cache_data(show_spinner=False)
def add_nearest_metro(objects: pd.DataFrame, metro_df: pd.DataFrame) -> pd.DataFrame:
    result = objects.copy()
    names: list[str] = []
    distances: list[int | None] = []
    lats: list[float | None] = []
    lons: list[float | None] = []

    for _, row in result.iterrows():
        name, distance_km, metro_lat, metro_lon = nearest_metro(
            float(row["Широта"]), float(row["Долгота"]), metro_df
        )
        names.append(name)
        distances.append(round(distance_km * 100) * 10)
        lats.append(metro_lat)
        lons.append(metro_lon)

    result["Ближайшее метро"] = names
    result["До метро, м"] = distances
    result["_metro_lat"] = lats
    result["_metro_lon"] = lons
    return result


# =========================================================
# КАРТЫ
# =========================================================


def object_popup(row: pd.Series) -> str:
    object_name = html.escape(clean_text(row["Объект"]) or "Без названия")
    fields = [
        ("ID", row["ID"]),
        ("Адрес", row["Адрес"]),
        ("Тип", row["Тип"]),
        ("Ответственный", row["Ответственный"]),
        ("Статус", row["Статус"]),
        ("Комментарий", row["Комментарий"]),
    ]
    field_html = "".join(
        f"<b>{label}:</b> {html.escape(clean_text(value)) or '—'}<br>"
        for label, value in fields
    )

    phone = clean_text(row["Телефон"])
    if phone:
        phone_digits = "".join(symbol for symbol in phone if symbol.isdigit() or symbol == "+")
        phone_html = f'<a href="tel:{phone_digits}">{html.escape(phone)}</a>'
    else:
        phone_html = "—"

    metro_name = html.escape(clean_text(row.get("Ближайшее метро", "")))
    metro_distance = row.get("До метро, м")
    metro_html = "Не определено"
    if metro_name and pd.notna(metro_distance):
        metro_html = f"{metro_name} — около {int(metro_distance)} м"

    lat = float(row["Широта"])
    lon = float(row["Долгота"])
    yandex_link = f"https://yandex.ru/maps/?pt={lon},{lat}&z=16&l=map"
    google_link = f"https://www.google.com/maps/dir/?api=1&destination={lat},{lon}"

    return f"""
    <div style="width:320px;font-family:Arial,sans-serif;font-size:14px;line-height:1.5">
      <h3 style="color:{BRAND_NAVY};margin:0 0 10px">{object_name}</h3>
      {field_html}
      <b>Телефон:</b> {phone_html}<br>
      <hr style="border:none;border-top:1px solid #DDE2ED">
      <b>Ⓜ Ближайшее метро:</b><br>{metro_html}<br><br>
      <a href="{yandex_link}" target="_blank">Открыть в Яндекс Картах</a><br>
      <a href="{google_link}" target="_blank">Построить маршрут</a>
    </div>
    """


def build_city_map(
    filtered: pd.DataFrame,
    show_metro: bool,
    show_lines: bool,
    address_result: dict[str, Any] | None,
    dgis_api_key: str,
) -> folium.Map:
    if address_result:
        map_center = [address_result["lat"], address_result["lon"]]
        zoom = 11
    else:
        map_center = [float(filtered["Широта"].mean()), float(filtered["Долгота"].mean())]
        zoom = DEFAULT_ZOOM

    if dgis_api_key:
        map_object = folium.Map(
            location=map_center,
            zoom_start=zoom,
            tiles=None,
            control_scale=True,
        )
        encoded_key = quote(dgis_api_key, safe="")
        folium.TileLayer(
            tiles=(
                "https://tile0.maps.2gis.com/v2/tiles/online_hd/"
                "{z}/{x}/{y}.png?key=" + encoded_key
            ),
            attr="© 2ГИС",
            name="2ГИС — актуальные здания",
            overlay=False,
            control=False,
            max_native_zoom=19,
            max_zoom=20,
        ).add_to(map_object)
    else:
        map_object = folium.Map(
            location=map_center,
            zoom_start=zoom,
            tiles="OpenStreetMap",
            control_scale=True,
        )
    cluster = MarkerCluster(name="Объекты").add_to(map_object)
    metro_added: set[tuple[str, float, float]] = set()

    for _, row in filtered.iterrows():
        lat = float(row["Широта"])
        lon = float(row["Долгота"])
        folium.Marker(
            location=[lat, lon],
            tooltip=clean_text(row["Объект"]) or "Объект",
            popup=folium.Popup(object_popup(row), max_width=390),
            icon=folium.Icon(color=marker_color(row["Статус"]), icon="wrench", prefix="fa"),
        ).add_to(cluster)

        metro_name = clean_text(row.get("Ближайшее метро", ""))
        metro_lat = row.get("_metro_lat")
        metro_lon = row.get("_metro_lon")
        if show_metro and metro_name and pd.notna(metro_lat) and pd.notna(metro_lon):
            key = (metro_name, round(float(metro_lat), 5), round(float(metro_lon), 5))
            if key not in metro_added:
                folium.CircleMarker(
                    location=[float(metro_lat), float(metro_lon)],
                    radius=7,
                    color=BRAND_BLUE,
                    fill=True,
                    fill_color=BRAND_BLUE,
                    fill_opacity=0.95,
                    tooltip="Ⓜ " + metro_name,
                ).add_to(map_object)
                metro_added.add(key)

            if show_lines:
                folium.PolyLine(
                    locations=[[lat, lon], [float(metro_lat), float(metro_lon)]],
                    color=BRAND_BLUE,
                    weight=2,
                    opacity=0.38,
                    dash_array="5,8",
                    tooltip=f"До метро около {int(row['До метро, м'])} м",
                ).add_to(map_object)

    if address_result:
        lat = address_result["lat"]
        lon = address_result["lon"]
        logistics = address_result["logistics"]
        popup = (
            f"<b>Проверяемый адрес</b><br>{html.escape(address_result['display_name'])}"
            f"<br><br><b>{html.escape(logistics['title'])}</b>"
        )
        folium.Marker(
            [lat, lon],
            tooltip="Проверяемый адрес",
            popup=folium.Popup(popup, max_width=360),
            icon=folium.Icon(color="cadetblue", icon="search", prefix="fa"),
        ).add_to(map_object)

        for route in logistics["route_segments"]:
            folium.PolyLine(
                route,
                color=BRAND_RED,
                weight=5,
                opacity=0.82,
                tooltip=f"{logistics['travel_mode']} · маршрут до ближайшего объекта",
            ).add_to(map_object)
        if not logistics["route_segments"]:
            folium.PolyLine(
                [[lat, lon], [logistics["nearest_lat"], logistics["nearest_lon"]]],
                color=BRAND_RED,
                weight=3,
                opacity=0.72,
                dash_array="7,8",
                tooltip="Направление до ближайшего объекта — маршрут не получен",
            ).add_to(map_object)

    folium.LayerControl(collapsed=True).add_to(map_object)
    PlainLeafletAttribution().add_to(map_object)
    return map_object


def build_metro_map(
    filtered: pd.DataFrame,
    metro_network: dict[str, Any],
    address_result: dict[str, Any] | None,
) -> folium.Map:
    stations = metro_network["stations"]
    map_object = folium.Map(
        location=DEFAULT_CENTER,
        zoom_start=10,
        tiles=None,
        control_scale=False,
        zoom_control=True,
        attribution_control=True,
    )

    for segment in metro_network.get("segments", []):
        folium.PolyLine(
            segment["points"],
            color=segment["color"],
            weight=5,
            opacity=0.82,
            tooltip=segment["line"],
        ).add_to(map_object)

    grouped: dict[str, list[pd.Series]] = {}
    for _, row in filtered.iterrows():
        grouped.setdefault(normalize_name(row["Ближайшее метро"]), []).append(row)

    for _, station in stations.iterrows():
        station_name = clean_text(station["Станция"])
        station_key = normalize_name(station_name)
        linked_objects = grouped.get(station_key, [])
        lat = float(station["Широта метро"])
        lon = float(station["Долгота метро"])
        color = safe_color(clean_text(station.get("Цвет", "")), BRAND_BLUE)

        if linked_objects:
            object_lines = "".join(
                "<li><b>"
                + html.escape(clean_text(row["Объект"]) or "Без названия")
                + "</b><br>"
                + html.escape(clean_text(row["Адрес"]))
                + "</li>"
                for row in linked_objects
            )
            popup = (
                f"<div style='width:300px;font-family:Arial'>"
                f"<h3 style='color:{BRAND_NAVY}'>Ⓜ {html.escape(station_name)}</h3>"
                f"<b>Объектов: {len(linked_objects)}</b><ul>{object_lines}</ul></div>"
            )
            icon_html = f"""
                <div style="
                    width:30px;height:30px;border-radius:50%;
                    background:{BRAND_RED};color:white;border:3px solid white;
                    box-shadow:0 2px 8px rgba(0,0,0,.28);
                    display:flex;align-items:center;justify-content:center;
                    font:bold 13px Arial;transform:translate(-8px,-8px)">
                    {len(linked_objects)}
                </div>
            """
            folium.Marker(
                [lat, lon],
                tooltip=f"{station_name}: {len(linked_objects)} объект(а)",
                popup=folium.Popup(popup, max_width=360),
                icon=folium.DivIcon(html=icon_html),
            ).add_to(map_object)
        else:
            folium.CircleMarker(
                [lat, lon],
                radius=4,
                color=color,
                weight=2,
                fill=True,
                fill_color="#FFFFFF",
                fill_opacity=1,
                tooltip=station_name,
            ).add_to(map_object)

    if address_result:
        address_lat = address_result["lat"]
        address_lon = address_result["lon"]
        station_name, distance_km, metro_lat, metro_lon = nearest_metro(
            address_lat, address_lon, stations
        )
        folium.Marker(
            [metro_lat, metro_lon],
            tooltip=f"Новый адрес → {station_name}",
            popup=folium.Popup(
                f"<b>Новый адрес</b><br>Ближайшее метро: {html.escape(station_name)}"
                f"<br>По прямой: около {distance_km:.1f} км",
                max_width=320,
            ),
            icon=folium.Icon(color="cadetblue", icon="search", prefix="fa"),
        ).add_to(map_object)

    if not stations.empty:
        map_object.fit_bounds(
            [
                [float(stations["Широта метро"].min()), float(stations["Долгота метро"].min())],
                [float(stations["Широта метро"].max()), float(stations["Долгота метро"].max())],
            ],
            padding=(18, 18),
        )
    PlainLeafletAttribution().add_to(map_object)
    return map_object


# =========================================================
# ШАПКА
# =========================================================


header_logo, header_text = st.columns([1.2, 3.8], vertical_alignment="center")
with header_logo:
    if LOGO_FILE.exists():
        st.image(str(LOGO_FILE), width="stretch")
with header_text:
    st.markdown(
        """
        <div class="linkor-header">
            <div class="linkor-kicker">Внутренний рабочий сервис</div>
            <div class="linkor-title">Карта обслуживаемых объектов</div>
            <div class="linkor-subtitle">Объекты, метро и предварительная оценка нового адреса</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

dgis_api_key = get_2gis_api_key()


# =========================================================
# ЗАГРУЗКА ДАННЫХ
# =========================================================


if not EXCEL_FILE.exists():
    st.error("Не найден файл objects.xlsx. Он должен находиться рядом с app.py.")
    st.stop()

try:
    objects = load_objects(EXCEL_FILE.stat().st_mtime)
except Exception as error:
    st.error("Не удалось прочитать objects.xlsx: " + str(error))
    st.stop()

if objects.empty:
    st.warning("В objects.xlsx пока нет объектов.")
    st.stop()

invalid_coordinates = objects[
    objects["Широта"].isna()
    | objects["Долгота"].isna()
    | ~objects["Широта"].between(-90, 90)
    | ~objects["Долгота"].between(-180, 180)
]
objects_valid = objects.drop(index=invalid_coordinates.index).copy()

if objects_valid.empty:
    st.error("Ни у одного объекта нет корректных координат.")
    st.stop()

metro_error: str | None = None
show_mcd_3_4 = bool(st.session_state.get("show_mcd_3_4", False))
try:
    with st.spinner("Обновляем схему метро…"):
        full_metro_network = load_metro_network()
    metro_network = filter_metro_network(full_metro_network, show_mcd_3_4)
    metro_df = metro_network["stations"]
    objects_valid = add_nearest_metro(objects_valid, metro_df)
except Exception as error:
    metro_error = str(error)
    metro_network = {"stations": pd.DataFrame(), "segments": []}
    metro_df = metro_network["stations"]
    for column, value in [
        ("Ближайшее метро", ""),
        ("До метро, м", None),
        ("_metro_lat", None),
        ("_metro_lon", None),
    ]:
        objects_valid[column] = value


# =========================================================
# БОКОВАЯ ПАНЕЛЬ И ФИЛЬТРЫ
# =========================================================


st.sidebar.markdown("### Фильтры")
st.sidebar.caption(
    "Карта: 2ГИС · актуальные здания"
    if dgis_api_key
    else "Карта: OpenStreetMap · резервный режим"
)
search_text = st.sidebar.text_input(
    "Поиск по базе", placeholder="Объект, адрес, ответственный…"
).strip()

all_types = sorted(value for value in objects_valid["Тип"].unique() if value)
all_statuses = sorted(value for value in objects_valid["Статус"].unique() if value)
all_people = sorted(value for value in objects_valid["Ответственный"].unique() if value)

selected_types = st.sidebar.multiselect("Тип", all_types, default=all_types)
selected_statuses = st.sidebar.multiselect("Статус", all_statuses, default=all_statuses)
selected_people = st.sidebar.multiselect("Ответственный", all_people, default=all_people)

st.sidebar.divider()
show_metro = st.sidebar.checkbox("Показывать ближайшее метро", value=True)
show_lines = st.sidebar.checkbox(
    "Линии от объектов до метро", value=False, disabled=not show_metro
)
show_mcd_3_4 = st.sidebar.checkbox(
    "Показывать МЦД-3 и МЦД-4",
    value=False,
    key="show_mcd_3_4",
    help="По умолчанию скрыты, чтобы длинные диаметры не уменьшали масштаб схемы метро.",
)
show_table = st.sidebar.checkbox("Таблица под картой", value=False)

with st.sidebar.expander("Правила логистики"):
    good_minutes = st.number_input(
        "Удобно, до минут", min_value=5, max_value=120, value=25, step=5
    )
    acceptable_minutes = st.number_input(
        "Допустимо, до минут", min_value=int(good_minutes) + 5, max_value=180, value=45, step=5
    )
    density_radius_km = st.number_input(
        "Радиус кластера, км", min_value=2, max_value=30, value=8, step=1
    )
    st.caption("Плотный кластер: не менее трёх действующих объектов в указанном радиусе.")

filtered = objects_valid.copy()
if all_types:
    filtered = filtered[filtered["Тип"].isin(selected_types)]
if all_statuses:
    filtered = filtered[filtered["Статус"].isin(selected_statuses)]
if all_people:
    filtered = filtered[filtered["Ответственный"].isin(selected_people)]

if search_text:
    search_lower = search_text.lower()
    mask = pd.Series(False, index=filtered.index)
    for column in [
        "ID",
        "Объект",
        "Адрес",
        "Тип",
        "Ответственный",
        "Телефон",
        "Статус",
        "Комментарий",
    ]:
        mask |= filtered[column].str.lower().str.contains(search_lower, regex=False)
    filtered = filtered[mask]


# =========================================================
# НОВЫЙ АДРЕС
# =========================================================


st.subheader("Проверить новый адрес")
st.caption(
    "Введите адрес потенциального объекта. Сервис сравнит его с действующей базой и покажет предварительную логистическую оценку."
)

with st.form("address_check", clear_on_submit=False):
    travel_mode = st.radio(
        "Способ передвижения",
        ["На машине", "Пешком", "Общественный транспорт"],
        horizontal=True,
    )
    address_col, button_col = st.columns([5, 1], vertical_alignment="bottom")
    with address_col:
        new_address = st.text_input(
            "Адрес",
            placeholder="Например: Москва, ул. Веерная, д. 24",
            label_visibility="collapsed",
        )
    with button_col:
        check_address = st.form_submit_button("Проверить", use_container_width=True)

if check_address:
    if not new_address.strip():
        st.warning("Введите адрес объекта.")
    else:
        try:
            with st.spinner("Ищем адрес и рассчитываем логистику…"):
                geocoded = geocode_address(new_address, dgis_api_key)
                if not geocoded:
                    raise ValueError("Адрес не найден. Уточните город, улицу и номер дома.")
                logistics = evaluate_logistics(
                    geocoded["lat"],
                    geocoded["lon"],
                    objects_valid,
                    int(good_minutes),
                    int(acceptable_minutes),
                    int(density_radius_km),
                    travel_mode,
                    dgis_api_key,
                )
                st.session_state["address_result"] = {**geocoded, "logistics": logistics}
        except Exception as error:
            st.error("Не удалось проверить адрес: " + str(error))

address_result = st.session_state.get("address_result")

if address_result:
    logistics = address_result["logistics"]
    minutes_text = (
        f"{logistics['minutes']:.0f} мин" if logistics["minutes"] is not None else "нет данных"
    )
    distance_text = (
        f"{logistics['road_km']:.1f} км"
        if logistics["road_km"] is not None
        else f"{logistics['air_km']:.1f} км по прямой"
    )

    st.markdown(
        f"""
        <div class="logistics-card verdict-{logistics['level']}">
            <div class="verdict-label">Предварительная оценка</div>
            <div class="verdict-title">{html.escape(logistics['title'])}</div>
            <div>{html.escape(logistics['reason'])}</div>
            <div class="metric-note" style="margin-top:8px">
                Найдено: {html.escape(address_result['display_name'])}
                · источник: {html.escape(address_result.get('provider', ''))}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    result_col1, result_col2, result_col3, result_col4 = st.columns(4)
    result_col1.metric("Ближайший объект", logistics["nearest_name"])
    result_col2.metric("Время · " + logistics["travel_mode"].lower(), minutes_text)
    result_col3.metric("Расстояние", distance_text)
    result_col4.metric(
        f"Объектов в радиусе {int(density_radius_km)} км", logistics["nearby_count"]
    )
    st.caption(
        "Ближайший объект: "
        + logistics["nearest_address"]
        + (". Маршрут: " + logistics["route_provider"] if logistics["routing_ok"] else "")
        + (f" · пересадок: {logistics['transfers']}" if logistics["transfers"] is not None else "")
        + ". Оценка не учитывает график инженеров и состав работ."
    )


# =========================================================
# СТАТИСТИКА И ПЕРЕКЛЮЧАТЕЛЬ КАРТ
# =========================================================


if filtered.empty:
    st.info("По выбранным фильтрам объектов не найдено.")
    st.stop()

metric1, metric2, metric3 = st.columns(3)
metric1.metric("Всего объектов", len(objects_valid))
metric2.metric("Показано", len(filtered))
metric3.metric("Станций метро", len(metro_df) if not metro_df.empty else "—")

if not invalid_coordinates.empty:
    st.warning(
        f"Не показано объектов: {len(invalid_coordinates)}. Проверьте широту и долготу в Excel."
    )
if metro_error:
    st.warning(
        "Сервис метро временно недоступен. Обычная карта работает, режим схемы метро появится после восстановления соединения."
    )

view_mode = st.radio(
    "Режим карты",
    ["Карта Москвы", "Схема метро"],
    horizontal=True,
    label_visibility="collapsed",
)

if view_mode == "Карта Москвы":
    if not dgis_api_key:
        st.warning(
            "На этом экране работает резервная карта OpenStreetMap: приложение не видит ключ 2ГИС. "
            "Проверьте в Streamlit Community Cloud → Settings → Secrets запись "
            "DGIS_API_KEY = \"...\", сохраните настройки и перезапустите приложение. "
            "Ключ в GitHub или в файле secrets.toml.example не подключает карту."
        )
    city_map = build_city_map(
        filtered, show_metro, show_lines, address_result, dgis_api_key
    )
    st_folium(city_map, height=690, use_container_width=True, returned_objects=[])
else:
    if metro_df.empty:
        st.info("Схема метро сейчас недоступна. Переключитесь на карту Москвы.")
    else:
        st.caption(
            "Красные маркеры показывают количество объектов у станции. Нажмите на маркер, чтобы увидеть список. "
            + (
                "МЦД-3 и МЦД-4 включены."
                if show_mcd_3_4
                else "МЦД-3 и МЦД-4 скрыты для компактного масштаба."
            )
        )
        metro_map = build_metro_map(filtered, metro_network, address_result)
        st_folium(metro_map, height=720, use_container_width=True, returned_objects=[])


# =========================================================
# ТАБЛИЦЫ
# =========================================================


if show_table:
    st.subheader("Список объектов")
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
        filtered[table_columns].reset_index(drop=True),
        use_container_width=True,
        hide_index=True,
    )

if not invalid_coordinates.empty:
    with st.expander("Объекты с ошибкой координат"):
        st.dataframe(
            invalid_coordinates[["ID", "Объект", "Адрес", "Широта", "Долгота"]].reset_index(
                drop=True
            ),
            use_container_width=True,
            hide_index=True,
        )
