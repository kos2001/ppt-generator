"""Bundled icons drawn as native PowerPoint shapes.

Icons come from Phosphor Icons (fill weight, MIT License — see
assets/icons/LICENSE-phosphor.txt), pre-converted by build_icon_set.py into
move/line/cubic path data. Each icon becomes one a:custGeom freeform shape,
so it stays sharp at any size and can be recolored or resized in PowerPoint
like any other shape — no images, no fonts.

Icons are referenced by Phosphor name ("lightbulb", "shield-check") or by a
concept alias in English or Korean ("idea", "보안"). See references/icons.md.
"""
import difflib
import json
import os

_ICON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                          "assets", "icons", "phosphor-fill.json")
_CACHE = {}

# Concept -> icon. Lets a spec say what the icon *means*; the renderer picks
# the drawable object.
ALIASES = {
    "idea": "lightbulb", "insight": "lightbulb", "아이디어": "lightbulb",
    "security": "shield-check", "secure": "shield-check", "보안": "shield-check",
    "privacy": "lock", "개인정보": "lock",
    "cost": "coins", "money": "coins", "revenue": "currency-dollar", "비용": "coins",
    "매출": "chart-line-up", "수익": "hand-coins", "krw": "currency-krw", "원화": "currency-krw",
    "growth": "trend-up", "성장": "trend-up", "trend": "chart-line-up",
    "team": "users-three", "people": "users", "팀": "users-three", "조직": "users-three",
    "customer": "user-circle", "user": "user", "고객": "user-circle", "사용자": "user",
    "ai": "robot", "automation": "robot", "자동화": "robot", "인공지능": "robot",
    "ml": "brain", "intelligence": "brain", "학습": "brain",
    "data": "database", "데이터": "database", "storage": "hard-drives", "저장소": "hard-drives",
    "cloud": "cloud", "클라우드": "cloud", "upload": "cloud-arrow-up",
    "speed": "lightning", "fast": "lightning", "속도": "lightning", "performance": "gauge",
    "성능": "gauge",
    "time": "clock", "시간": "clock", "schedule": "calendar", "일정": "calendar",
    "deadline": "calendar-check",
    "goal": "target", "목표": "target", "focus": "target", "strategy": "strategy",
    "전략": "strategy",
    "risk": "warning", "위험": "warning", "리스크": "warning", "issue": "warning-circle",
    "search": "magnifying-glass", "검색": "magnifying-glass", "research": "binoculars",
    "settings": "gear", "설정": "gear", "operations": "gear", "운영": "gear",
    "launch": "rocket-launch", "출시": "rocket-launch", "startup": "rocket",
    "quality": "seal-check", "품질": "seal-check", "done": "check-circle", "완료": "check-circle",
    "checklist": "list-checks", "체크리스트": "list-checks",
    "communication": "chats", "chat": "chat-circle", "소통": "chats", "상담": "headset",
    "support": "headset", "지원": "headset", "voice": "microphone", "음성": "microphone",
    "email": "envelope", "메일": "envelope", "notification": "bell", "알림": "bell",
    "global": "globe", "글로벌": "globe", "language": "translate", "다국어": "translate",
    "번역": "translate",
    "knowledge": "book-open", "지식": "book-open", "docs": "file-text", "문서": "file-text",
    "education": "graduation-cap", "교육": "graduation-cap",
    "code": "code", "코드": "code", "development": "code", "개발": "code",
    "infrastructure": "hard-drives", "인프라": "hard-drives", "server": "hard-drives",
    "network": "network", "네트워크": "network", "integration": "plugs-connected",
    "연동": "plugs-connected", "api": "plugs-connected",
    "process": "flow-arrow", "프로세스": "flow-arrow", "workflow": "flow-arrow",
    "analytics": "chart-bar", "분석": "chart-bar", "report": "presentation-chart",
    "보고": "presentation-chart", "dashboard": "squares-four",
    "partnership": "handshake", "협력": "handshake", "파트너": "handshake",
    "award": "trophy", "성과": "trophy", "win": "trophy",
    "location": "map-pin", "위치": "map-pin", "factory": "factory", "공장": "factory",
    "company": "buildings", "회사": "buildings", "store": "storefront", "매장": "storefront",
    "shipping": "truck", "배송": "truck", "물류": "truck", "commerce": "shopping-cart",
    "payment": "credit-card", "결제": "credit-card", "finance": "bank", "금융": "bank",
    "legal": "scales", "법무": "scales", "compliance": "scales", "규정": "scales",
    "health": "heartbeat", "건강": "heartbeat", "sustainability": "plant", "친환경": "plant",
    "experiment": "flask", "실험": "flask", "test": "flask", "테스트": "flask",
    "bug": "bug", "버그": "bug", "mobile": "device-mobile", "모바일": "device-mobile",
    "desktop": "desktop", "video": "video-camera", "영상": "video-camera",
    "star": "star", "favorite": "heart", "feedback": "thumbs-up", "만족": "thumbs-up",
    "key": "key", "access": "key", "권한": "key", "identity": "fingerprint", "인증": "fingerprint",
    "marketing": "megaphone", "마케팅": "megaphone", "news": "newspaper",
    "package": "package", "product": "cube", "제품": "cube", "platform": "stack",
    "플랫폼": "stack", "layers": "stack", "puzzle": "puzzle-piece", "solution": "puzzle-piece",
    "솔루션": "puzzle-piece", "hierarchy": "tree-structure", "구조": "tree-structure",
}


def _icons():
    if "icons" not in _CACHE:
        try:
            with open(_ICON_FILE, encoding="utf-8") as f:
                _CACHE["icons"] = json.load(f)["icons"]
        except (OSError, ValueError, KeyError):
            _CACHE["icons"] = {}
    return _CACHE["icons"]


def names():
    return sorted(_icons())


def resolve(name):
    """Phosphor name or alias -> bundled icon name, or None."""
    if not name:
        return None
    key = str(name).strip().lower().replace(" ", "-").replace("_", "-")
    if key.endswith("-fill"):
        key = key[:-5]
    icons = _icons()
    if key in icons:
        return key
    alias = ALIASES.get(key) or ALIASES.get(str(name).strip())
    return alias if alias in icons else None


def suggest(name, n=3):
    pool = list(_icons()) + list(ALIASES)
    return difflib.get_close_matches(str(name).lower(), pool, n=n, cutoff=0.5)


def add_icon(slide, name, left, top, size, color_hex):
    """Draw icon `name` as a freeform shape in a size x size box. Returns the
    shape, or None when the icon is unknown (callers draw a fallback)."""
    # Imported here so name lookup (used by the validator) works without
    # python-pptx installed.
    from lxml import etree
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml.ns import qn

    key = resolve(name)
    if key is None:
        return None
    icon = _icons()[key]
    scale = 100  # path units: 1/100 of the SVG unit, kept integral
    w = h = icon["size"] * scale
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, int(left), int(top),
                                int(size), int(size))
    sp.name = "Icon: %s" % key
    spPr = sp._element.spPr
    prst = spPr.find(qn("a:prstGeom"))
    cust = etree.Element(qn("a:custGeom"))
    for tag in ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst"):
        etree.SubElement(cust, qn(tag))
    rect = etree.SubElement(cust, qn("a:rect"))
    for k, v in (("l", "l"), ("t", "t"), ("r", "r"), ("b", "b")):
        rect.set(k, v)
    path_lst = etree.SubElement(cust, qn("a:pathLst"))
    path = etree.SubElement(path_lst, qn("a:path"))
    path.set("w", str(w))
    path.set("h", str(h))

    def pt(parent, x, y):
        p = etree.SubElement(parent, qn("a:pt"))
        p.set("x", str(int(round(x * scale))))
        p.set("y", str(int(round(y * scale))))

    for op in icon["ops"]:
        kind = op[0]
        if kind == "M":
            pt(etree.SubElement(path, qn("a:moveTo")), op[1], op[2])
        elif kind == "L":
            pt(etree.SubElement(path, qn("a:lnTo")), op[1], op[2])
        elif kind == "C":
            c = etree.SubElement(path, qn("a:cubicBezTo"))
            pt(c, op[1], op[2])
            pt(c, op[3], op[4])
            pt(c, op[5], op[6])
        elif kind == "Z":
            etree.SubElement(path, qn("a:close"))
    prst.addprevious(cust)
    spPr.remove(prst)
    sp.fill.solid()
    sp.fill.fore_color.rgb = RGBColor.from_string(color_hex)
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp
