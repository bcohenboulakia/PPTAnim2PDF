import argparse
import copy
import posixpath
import re
import zipfile
from pathlib import Path

from lxml import etree


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

P = NS["p"]
A = NS["a"]
R = NS["r"]
CT = NS["ct"]
REL = NS["rel"]

MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

SLIDE_RE = re.compile(r"^ppt/slides/slide\d+\.xml$")
SLIDE_RELS_RE = re.compile(r"^ppt/slides/_rels/slide\d+\.xml\.rels$")

FULL_TURN = 360 * 60000

TEXT_STYLE_KINDS = {"bold", "italic", "underline", "strike"}

def qn(ns, tag):
    return f"{{{ns}}}{tag}"

COLOR_TAGS = {
    qn(A, "scrgbClr"),
    qn(A, "srgbClr"),
    qn(A, "hslClr"),
    qn(A, "sysClr"),
    qn(A, "schemeClr"),
    qn(A, "prstClr"),
}

FILL_TAGS = {
    qn(A, "noFill"),
    qn(A, "solidFill"),
    qn(A, "gradFill"),
    qn(A, "blipFill"),
    qn(A, "pattFill"),
    qn(A, "grpFill"),
}


def parse_xml(data):
    return etree.fromstring(data)


def xml_bytes(root):
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=False,
    )

def rel_target_to_part(source_part, target):
    base = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(base, target))

def remove_timing_and_transition(slide_root):
    for el in list(slide_root):
        if el.tag in {qn(P, "timing"), qn(P, "transition")}:
            slide_root.remove(el)


def is_inside_mc_fallback(el):
    cur = el.getparent()

    while cur is not None:
        q = etree.QName(cur)

        if q.namespace == MC_NS and q.localname == "Fallback":
            return True

        cur = cur.getparent()

    return False


def all_shape_elements(slide_root):
    result = {}

    candidates = slide_root.xpath(
        ".//p:sp | .//p:pic | .//p:graphicFrame | .//p:cxnSp | .//p:grpSp",
        namespaces=NS,
    )

    paths = {
        qn(P, "sp"): "./p:nvSpPr/p:cNvPr",
        qn(P, "pic"): "./p:nvPicPr/p:cNvPr",
        qn(P, "graphicFrame"): "./p:nvGraphicFramePr/p:cNvPr",
        qn(P, "cxnSp"): "./p:nvCxnSpPr/p:cNvPr",
        qn(P, "grpSp"): "./p:nvGrpSpPr/p:cNvPr",
    }

    for el in candidates:
        if is_inside_mc_fallback(el):
            continue

        path = paths.get(el.tag)
        if not path:
            continue

        c_nv_pr = el.find(path, namespaces=NS)

        if c_nv_pr is not None and c_nv_pr.get("id"):
            result[c_nv_pr.get("id")] = el

    return result


def paragraphs_of_shape(shape_el):
    return shape_el.xpath(".//p:txBody/a:p", namespaces=NS)


def target_from_behavior(cbhvr):
    targets = []

    sp_tgt = cbhvr.find(".//p:tgtEl/p:spTgt", namespaces=NS)
    if sp_tgt is None:
        return targets

    sid = sp_tgt.get("spid")
    if not sid:
        return targets

    p_rg = sp_tgt.find("./p:txEl/p:pRg", namespaces=NS)

    if p_rg is None:
        return [("shape", sid)]

    try:
        start = int(p_rg.get("st", "0"))
        end = int(p_rg.get("end", str(start)))
    except ValueError:
        return [("shape", sid)]

    if end < start:
        start, end = end, start

    return [("paragraph", sid, i) for i in range(start, end + 1)]


def dedup_targets(targets):
    seen = set()
    out = []

    for target in targets:
        if target not in seen:
            seen.add(target)
            out.append(target)

    return out

def parse_motion_path_numbers(path):
    """
    Extrait commandes et nombres depuis un path PowerPoint.

    Exemple typique :
        M 0 0 L 0.25 0 E
    """
    if not path:
        return []

    token_re = re.compile(
        r"[MLCZmlczEez]|[-+]?(?:\d+\.\d+|\d+|\.\d+)(?:[eE][-+]?\d+)?"
    )

    return token_re.findall(path)


def final_point_from_motion_path(path):
    """
    Retourne le déplacement final normalisé (dx, dy) d'un motion path.

    Les coordonnées PowerPoint de motion path sont généralement exprimées
    comme fractions de la taille de la slide. Par exemple, x=0.25 signifie
    environ un quart de largeur de slide.

    On ne conserve que le point final :
    - L : dernier point du segment ;
    - C : dernier point de la courbe ;
    - Z : retour au point initial du chemin ;
    - E : fin.
    """
    tokens = parse_motion_path_numbers(path)

    if not tokens:
        return None

    i = 0
    cmd = None

    start = None
    current = (0.0, 0.0)
    last_point = current

    def read_float():
        nonlocal i
        if i >= len(tokens):
            raise ValueError("motion path incomplet")

        value = float(tokens[i])
        i += 1
        return value

    while i < len(tokens):
        token = tokens[i]

        if re.fullmatch(r"[MLCZmlczEez]", token):
            cmd = token
            i += 1

            if cmd in {"E", "e"}:
                break

            if cmd in {"Z", "z"}:
                if start is not None:
                    current = start
                    last_point = current
                continue

        if cmd is None:
            break

        absolute = cmd.isupper()
        cmd_lower = cmd.lower()

        if cmd_lower in {"m", "l"}:
            x = read_float()
            y = read_float()

            if absolute:
                current = (x, y)
            else:
                current = (current[0] + x, current[1] + y)

            if start is None:
                start = current

            last_point = current

        elif cmd_lower == "c":
            # Courbe de Bézier : trois points, le dernier est l'arrivée.
            coords = []

            for _ in range(3):
                x = read_float()
                y = read_float()
                coords.append((x, y))

            end_x, end_y = coords[-1]

            if absolute:
                current = (end_x, end_y)
            else:
                current = (current[0] + end_x, current[1] + end_y)

            if start is None:
                start = current

            last_point = current

        else:
            break

    if start is None:
        start = (0.0, 0.0)

    return last_point[0] - start[0], last_point[1] - start[1]

def motion_delta_from_anim_motion(anim_motion, slide_width, slide_height):
    """
    Convertit un p:animMotion en déplacement EMU.

    Retourne :
        ((dx, dy), None)
    ou :
        (None, "closed_or_zero_motion_paths")
        (None, "unsupported_motion_paths")

    Un motion path dont le point final est identique au point initial
    est une animation purement dynamique : elle ne peut pas être rendue
    par un changement de slide statique.
    """
    path = anim_motion.get("path")

    try:
        point = final_point_from_motion_path(path)
    except Exception:
        return None, "unsupported_motion_paths"

    if point is None:
        return None, "unsupported_motion_paths"

    dx_norm, dy_norm = point

    dx = int(round(dx_norm * slide_width))
    dy = int(round(dy_norm * slide_height))

    if dx == 0 and dy == 0:
        return None, "closed_or_zero_motion_paths"

    return (dx, dy), None

def scale_value_to_factor(value):
    """
    Convertit une valeur OOXML de scale en facteur Python.

    Dans les animations PowerPoint, 100000 correspond à 100%.
    Donc :
      150000 -> 1.5
       50000 -> 0.5
    """
    if value is None:
        return None

    raw = str(value).strip()

    if not raw:
        return None

    if raw.endswith("%"):
        return float(raw[:-1]) / 100.0

    return float(raw) / 100000.0


def scale_factor_from_anim_scale(anim_scale):
    """
    Convertit un p:animScale en facteur de scale final.

    Retourne :
        ((sx, sy), None)
    ou :
        (None, "unsupported_scale_animations")
        (None, "neutral_scale_animations")

    Pour Grow/Shrink, PowerPoint encode généralement le facteur dans
    p:by ou p:to avec x/y en cent-millièmes de pourcentage.
    """
    scale_node = None

    for child_name in ("by", "to"):
        candidate = anim_scale.find(f"./p:{child_name}", namespaces=NS)

        if candidate is not None:
            scale_node = candidate
            break

    if scale_node is None:
        return None, "unsupported_scale_animations"

    sx = scale_value_to_factor(scale_node.get("x"))
    sy = scale_value_to_factor(scale_node.get("y"))

    if sx is None and sy is None:
        return None, "unsupported_scale_animations"

    if sx is None:
        sx = sy

    if sy is None:
        sy = sx

    if sx is None or sy is None:
        return None, "unsupported_scale_animations"

    if sx <= 0 or sy <= 0:
        return None, "unsupported_scale_animations"

    if abs(sx - 1.0) < 1e-9 and abs(sy - 1.0) < 1e-9:
        return None, "neutral_scale_animations"

    return (sx, sy), None

def clamp_alpha(alpha):
    return max(0, min(100000, int(round(alpha))))


def opacity_value_to_alpha(value, value_is_transparency=False):
    """
    Convertit une valeur d'animation en alpha OOXML.

    alpha OOXML :
      100000 -> opaque
       50000 -> 50% opaque
           0 -> invisible

    Si value_is_transparency=True, la valeur est interprétée comme un
    pourcentage de transparence et convertie en opacité.
    """
    if value is None:
        return None

    raw = str(value).strip()

    if not raw:
        return None

    try:
        if raw.endswith("%"):
            fixed = float(raw[:-1]) * 1000.0
        else:
            numeric = float(raw)

            if 0.0 <= numeric <= 1.0:
                fixed = numeric * 100000.0
            elif 0.0 <= numeric <= 100.0:
                fixed = numeric * 1000.0
            else:
                fixed = numeric
    except ValueError:
        return None

    fixed = clamp_alpha(fixed)

    if value_is_transparency:
        fixed = 100000 - fixed

    return clamp_alpha(fixed)


def final_value_from_anim(anim):
    """
    Récupère une valeur finale dans un p:anim.

    PowerPoint peut encoder certains effets via :
    - p:to ;
    - p:tavLst / p:tav, avec la dernière valeur de la séquence.
    """
    value_paths = [
        "./p:to/p:strVal/@val",
        "./p:to/p:fltVal/@val",
        "./p:to/p:intVal/@val",
        "./p:to/*/@val",
    ]

    for path in value_paths:
        values = anim.xpath(path, namespaces=NS)
        if values:
            return values[-1]

    tavs = anim.xpath("./p:tavLst/p:tav", namespaces=NS)

    if tavs:
        def tav_time(tav):
            try:
                return int(tav.get("tm", "0"))
            except ValueError:
                return 0

        last_tav = sorted(tavs, key=tav_time)[-1]

        values = last_tav.xpath("./p:val/*/@val", namespaces=NS)
        if values:
            return values[-1]

    return None


def opacity_alpha_from_anim(anim):
    """
    Détecte une animation de transparence/opacité et renvoie l'alpha final.

    Cas attendu pour l'effet Transparency :
    - attrName contenant opacity ;
    - valeur finale dans p:to ou p:tavLst.

    Retourne :
        (alpha, None)
    ou :
        (None, None) si ce n'est pas une animation d'opacité
    ou :
        (None, "unsupported_opacity_animations")
    """
    attr_names = [
        t.strip().lower()
        for t in anim.xpath(".//p:attrName/text()", namespaces=NS)
        if t and t.strip()
    ]

    if not attr_names:
        return None, None

    has_opacity = any("opacity" in name for name in attr_names)
    has_transparency = any("transpar" in name for name in attr_names)

    if not has_opacity and not has_transparency:
        return None, None

    final_value = final_value_from_anim(anim)

    if final_value is None:
        return None, "unsupported_opacity_animations"

    alpha = opacity_value_to_alpha(
        final_value,
        value_is_transparency=has_transparency and not has_opacity,
    )

    if alpha is None:
        return None, "unsupported_opacity_animations"

    return alpha, None

def rotation_value_to_ooxml(value):
    """
    Convertit une valeur de rotation OOXML en entier.

    Les rotations PowerPoint sont exprimées en 1/60000e de degré :
      90°  ->  5400000
      360° -> 21600000
    """
    if value is None:
        return None

    raw = str(value).strip()

    if not raw:
        return None

    # Cas pratique si un export produit une valeur lisible en degrés.
    if raw.endswith("deg"):
        return int(round(float(raw[:-3]) * 60000))

    return int(round(float(raw)))


def normalize_rotation(rot):
    """
    Normalise une rotation pour éviter les valeurs énormes.

    Visuellement, 0°, 360° et 720° sont équivalents.
    """
    return rot % FULL_TURN


def rotation_transform_from_anim_rot(anim_rot):
    """
    Convertit un p:animRot en transformation de rotation.

    Retourne :
        ({"rotation_mode": "delta", "rot": value}, None)
        ({"rotation_mode": "absolute", "rot": value}, None)

    ou :
        (None, "unsupported_rotation_animations")
        (None, "neutral_rotation_animations")

    Règles :
    - by : rotation relative à l'état courant ;
    - to : rotation finale absolue ;
    - from + to : on garde l'état final absolu to.
    """
    by_value = rotation_value_to_ooxml(anim_rot.get("by"))
    to_value = rotation_value_to_ooxml(anim_rot.get("to"))
    from_value = rotation_value_to_ooxml(anim_rot.get("from"))

    if by_value is not None:
        if normalize_rotation(by_value) == 0:
            return None, "neutral_rotation_animations"

        return {
            "rotation_mode": "delta",
            "rot": by_value,
        }, None

    if to_value is not None:
        if from_value is not None and normalize_rotation(to_value - from_value) == 0:
            return None, "neutral_rotation_animations"

        return {
            "rotation_mode": "absolute",
            "rot": to_value,
        }, None

    return None, "unsupported_rotation_animations"

def visibility_action_from_effect_container(ctn):
    """
    Réduit une animation PowerPoint à son effet statique final.
    On garde uniquement la sémantique apparition/disparition.
    """
    preset = ctn.get("presetClass")

    if preset == "entr":
        return "show"

    if preset == "exit":
        return "hide"

    attr_names = [
        t.strip().lower()
        for t in ctn.xpath(".//p:attrName/text()", namespaces=NS)
        if t and t.strip()
    ]

    values = []

    for xp in [
        ".//p:to/p:strVal/@val",
        ".//p:to/p:boolVal/@val",
        ".//p:to/p:intVal/@val",
        ".//p:to/*/@val",
    ]:
        values.extend(ctn.xpath(xp, namespaces=NS))

    values = [
        str(v).strip().lower()
        for v in values
        if str(v).strip()
    ]

    if any("visibility" in name for name in attr_names):
        if any(v in {"visible", "true", "1"} for v in values):
            return "show"

        if any(v in {"hidden", "false", "0"} for v in values):
            return "hide"

    return None

def has_nonzero_delay(ctn):
    """
    Détecte un délai explicite sur l'effet.

    La plupart du temps, le délai est directement sur le cTn de l'effet :
        p:cTn / p:stCondLst / p:cond delay="..."

    On garde un fallback descendant pour certains encodages PowerPoint.
    """
    delays = ctn.xpath("./p:stCondLst/p:cond/@delay", namespaces=NS)

    if not delays:
        delays = ctn.xpath(".//p:cBhvr/p:cTn/p:stCondLst/p:cond/@delay", namespaces=NS)

    for raw_delay in delays:
        delay = str(raw_delay).strip().lower()

        if delay in {"", "0", "indefinite"}:
            continue

        return True

    return False

def step_for_effect(node_type, ctn, current_step):
    """
    Convertit les déclenchements PowerPoint en étapes statiques.

    Règles :
    - clickEffect : nouvelle étape.
    - afterEffect après une étape existante : nouvelle étape.
    - withEffect sans délai : même étape.
    - withEffect avec délai : nouvelle étape.
    - effet automatique à délai 0 avant toute étape : étape 0.

    L'étape 0 correspond à l'état immédiatement affiché de la slide.
    """
    has_delay = has_nonzero_delay(ctn)

    if node_type == "clickEffect":
        return current_step + 1

    if node_type == "afterEffect":
        if current_step == 0 and not has_delay:
            return 0

        return current_step + 1

    if node_type == "withEffect":
        if current_step == 0:
            return 1 if has_delay else 0

        return current_step + 1 if has_delay else current_step

    return current_step + 1

def text_style_kind_from_attr_name(attr_name):
    name = attr_name.strip().lower()

    if name == "style.fontweight":
        return "bold"

    if name == "style.fontstyle":
        return "italic"

    if name == "style.textdecorationunderline":
        return "underline"

    if name == "style.textdecorationlinethrough":
        return "strike"

    return None


def container_has_supported_text_style(ctn):
    attr_names = [
        t.strip()
        for t in ctn.xpath(".//p:attrName/text()", namespaces=NS)
        if t and t.strip()
    ]

    return any(
        text_style_kind_from_attr_name(attr_name) is not None
        for attr_name in attr_names
    )

def unsupported_text_animation_kind(ctn):
    """
    Détecte les animations de texte plus fines que le paragraphe.

    Les effets de style/couleur texte que l'on sait aplatir sont acceptés,
    même si PowerPoint les encode avec p:iterate type="lt" ou "wd".
    """
    if container_has_supported_text_or_color_effect(ctn):
        return None

    iterate_nodes = ctn.xpath(
        "./p:iterate | ancestor::p:cTn/p:iterate",
        namespaces=NS,
    )

    for iterate in iterate_nodes:
        iterate_type = iterate.get("type", "el").strip().lower()

        if iterate_type == "wd":
            return "word"

        if iterate_type == "lt":
            return "letter"

    if ctn.xpath(".//p:tgtEl/p:spTgt/p:txEl/p:charRg", namespaces=NS):
        return "character"

    return None

def ordered_transform_nodes_from_container(ctn):
    """
    Retourne les nœuds de transformation dans l'ordre XML réel.

    Les nœuds dans p:subTnLst sont exclus : ils correspondent à des
    effets post-animation, comme Dim after animation.
    """
    preset = ctn.get("presetClass")

    if preset == "path":
        return ctn.xpath(
            ".//p:animMotion[not(ancestor::p:subTnLst)]",
            namespaces=NS,
        )

    if preset == "emph":
        return ctn.xpath(
            (
                ".//*["
                "not(ancestor::p:subTnLst) and "
                "("
                "self::p:animScale or "
                "self::p:animRot or "
                "self::p:animClr or "
                "self::p:anim or "
                "self::p:set"
                ")"
                "]"
            ),
            namespaces=NS,
        )

    return []

def color_key(color_el):
    """
    Clé normalisée pour comparer deux couleurs.

    Exemples :
      <a:srgbClr val="FF0000"/>    -> ("srgbClr", "FF0000")
      <a:schemeClr val="accent1"/> -> ("schemeClr", "accent1")
    """
    if color_el is None:
        return ("unknown", None)

    if color_el == "none":
        return ("none", None)

    q = etree.QName(color_el)
    return (q.localname, color_el.get("val"))


def color_from_color_container(container_el):
    """
    Récupère la première couleur DrawingML contenue dans un élément
    p:from / p:to / p:by.
    """
    if container_el is None:
        return None

    for el in container_el.iter():
        if el.tag in COLOR_TAGS:
            return copy.deepcopy(el)

    return None


def color_from_anim_clr_child(anim_clr, child_name):
    child = anim_clr.find(f"./p:{child_name}", namespaces=NS)
    return color_from_color_container(child)


def color_kind_from_anim_clr(anim_clr):
    """
    Détermine la cible d'un p:animClr.

    Encodages attendus :
    - fillcolor    -> remplissage
    - stroke.color -> contour
    - style.color  -> couleur de texte
    - ppt_c        -> couleur de texte, utilisé notamment par Dim after animation
    """
    attr_names = [
        t.strip().lower()
        for t in anim_clr.xpath(".//p:attrName/text()", namespaces=NS)
        if t and t.strip()
    ]

    for name in attr_names:
        if name == "fillcolor" or "fill.color" in name:
            return "fill"

        if name == "stroke.color" or name == "line.color":
            return "line"

        if name in {"style.color", "font.color", "ppt_c"}:
            return "text"

    return None

def color_transition_from_anim_clr(anim_clr):
    """
    Convertit un p:animClr en transition de couleur.

    Retourne :
        ({
            "color_kind": "fill" | "line" | "text",
            "from_color": color_or_None,
            "to_color": color,
        }, None)

    ou :
        (None, "unsupported_color_animations")
    """
    color_kind = color_kind_from_anim_clr(anim_clr)

    if color_kind is None:
        return None, "unsupported_color_animations"

    from_color = color_from_anim_clr_child(anim_clr, "from")
    to_color = color_from_anim_clr_child(anim_clr, "to")

    if to_color is None:
        return None, "unsupported_color_animations"

    return {
        "color_kind": color_kind,
        "from_color": from_color,
        "to_color": to_color,
    }, None

def boolean_text_style_value(raw_value, style_kind):
    if raw_value is None:
        return None

    value = str(raw_value).strip().lower()

    style_values = {
        "bold": (
            {"bold"},
            {"normal", "none", "false", "0"},
        ),
        "italic": (
            {"italic"},
            {"normal", "none", "false", "0"},
        ),
        "underline": (
            {"true", "t", "1"},
            {"false", "f", "0", "none", "normal"},
        ),
        "strike": (
            {"true", "t", "1"},
            {"false", "f", "0", "none", "normal"},
        ),
    }

    spec = style_values.get(style_kind)

    if spec is None:
        return None

    true_values, false_values = spec

    if value in true_values:
        return True

    if value in false_values:
        return False

    return None

def text_style_change_from_anim(anim):
    """
    Détecte les effets texte discrets :
    - Bold
    - Italic
    - Underline
    - Strikethrough

    Retourne :
        ({"text_style": kind, "value": bool}, None)
    ou :
        (None, None) si ce n'est pas un effet texte reconnu
    ou :
        (None, "unsupported_text_style_animations")
    """
    attr_names = [
        t.strip()
        for t in anim.xpath(".//p:attrName/text()", namespaces=NS)
        if t and t.strip()
    ]

    for attr_name in attr_names:
        text_style = text_style_kind_from_attr_name(attr_name)

        if text_style is None:
            continue

        final_value = final_value_from_anim(anim)

        if final_value is None:
            return None, "unsupported_text_style_animations"

        value = boolean_text_style_value(final_value, text_style)

        if value is None:
            return None, "unsupported_text_style_animations"

        return {
            "text_style": text_style,
            "value": value,
        }, None

    return None, None

def attr_names_from_animation_node(node):
    return [
        t.strip().lower()
        for t in node.xpath(".//p:attrName/text()", namespaces=NS)
        if t and t.strip()
    ]


def container_has_supported_text_or_color_effect(ctn):
    """
    Permet d'éviter qu'un effet texte supporté soit rejeté uniquement parce
    qu'il est encodé avec p:iterate type="lt" ou type="wd".

    PowerPoint encode par exemple Underline avec type="lt", même si nous
    savons l'aplatir en état final.
    """
    attr_names = attr_names_from_animation_node(ctn)

    for name in attr_names:
        if text_style_kind_from_attr_name(name) is not None:
            return True

        if name in {"style.color", "font.color"}:
            return True

    return False

def after_animation_events_from_node(animation_node):
    """
    Extrait les effets post-animation connus depuis un nœud de p:subTnLst.

    Supporté :
    - Dim after animation via p:animClr, souvent attrName=ppt_c ;
    - Dim after animation via p:set/p:anim sur couleur texte ;
    - Hide after animation via style.visibility -> hidden.
    """
    tag = animation_node.tag

    # Cas principal observé : p:animClr / ppt_c / to=...
    if tag == qn(P, "animClr"):
        color_transition, _ = color_transition_from_anim_clr(animation_node)

        if color_transition is None:
            return []

        targets = []

        for cbhvr in animation_node.xpath("./p:cBhvr", namespaces=NS):
            targets.extend(target_from_behavior(cbhvr))

        targets = dedup_targets(targets)

        if not targets:
            return []

        return [
            {
                "target": target,
                "action": "color",
                "color_kind": color_transition["color_kind"],
                "color": color_transition["to_color"],
            }
            for target in targets
        ]

    # Cas génériques p:set / p:anim.
    if tag in {qn(P, "set"), qn(P, "anim")}:
        attr_names = attr_names_from_animation_node(animation_node)

        targets = []

        for cbhvr in animation_node.xpath("./p:cBhvr", namespaces=NS):
            targets.extend(target_from_behavior(cbhvr))

        targets = dedup_targets(targets)

        if not targets:
            return []

        # Dim after animation : changement de couleur texte.
        if any(name in {"style.color", "font.color", "ppt_c"} for name in attr_names):
            to_node = animation_node.find("./p:to", namespaces=NS)
            color = color_from_color_container(to_node)

            if color is None:
                return []

            return [
                {
                    "target": target,
                    "action": "color",
                    "color_kind": "text",
                    "color": color,
                }
                for target in targets
            ]

        # Hide after animation.
        if any("visibility" in name for name in attr_names):
            final_value = final_value_from_anim(animation_node)

            if final_value is None:
                return []

            value = str(final_value).strip().lower()

            if value in {"hidden", "false", "0"}:
                return [
                    {
                        "target": target,
                        "action": "hide",
                    }
                    for target in targets
                ]

    return []

def after_animation_events_from_container(ctn):
    """
    Retourne les événements post-animation rattachés à un cTn.

    Chaque élément retourné est :
        (master_relation, [event_template, ...])

    master_relation est typiquement "nextClick" pour Dim after animation.
    """
    result = []

    for animation_node in ctn.xpath("./p:subTnLst/*", namespaces=NS):
        sub_ctn = animation_node.find("./p:cBhvr/p:cTn", namespaces=NS)

        if sub_ctn is None:
            continue

        after_effect = str(sub_ctn.get("afterEffect", "")).strip().lower()

        if after_effect not in {"1", "true", "t"}:
            continue

        event_templates = after_animation_events_from_node(animation_node)

        if not event_templates:
            continue

        result.append(
            (
                sub_ctn.get("masterRel", ""),
                event_templates,
            )
        )

    return result

def is_synthetic_all_at_once_visibility_setup(ctn):
    """
    Ignore certains nœuds techniques générés par PowerPoint pour les listes.

    Dans le fichier testé, PowerPoint 2019 ajoute des withEffect d'entrée
    avec grpId="0" qui rendent tous les paragraphes visibles au step 0.
    Ces nœuds ne correspondent pas aux clics réels de l'utilisateur.
    """
    if ctn.get("nodeType") != "withEffect":
        return False

    if ctn.get("presetClass") != "entr":
        return False

    if ctn.get("grpId") is None:
        return False

    if ctn.find("./p:subTnLst", namespaces=NS) is not None:
        return False

    attr_names = attr_names_from_animation_node(ctn)

    if set(attr_names) != {"style.visibility"}:
        return False

    # On limite l'exclusion aux nœuds placés dans une séquence onBegin,
    # ce qui correspond au setup initial observé.
    return bool(
        ctn.xpath(
            "ancestor::p:cTn[p:stCondLst/p:cond[@evt='onBegin']]",
            namespaces=NS,
        )
    )

def all_at_once_build_shape_ids(slide_root):
    """
    Retourne les IDs de formes dont les paragraphes sont construits
    all-at-once.

    Dans ce cas, les paragraphes sont visibles dès l'état initial, même si
    PowerPoint ajoute des nœuds internes style.visibility=visible dans la
    timeline.
    """
    return {
        bld.get("spid")
        for bld in slide_root.xpath(
            ".//p:timing/p:bldLst/p:bldP[@build='allAtOnce']",
            namespaces=NS,
        )
        if bld.get("spid")
    }

def extract_timeline_events(slide_root, slide_width, slide_height, report=None):
    """
    Extrait les événements statiques à partir de la timeline.

    Les effets dans p:subTnLst, comme Dim after animation, sont traités
    comme des événements post-animation. Quand masterRel="nextClick", ils
    sont appliqués à l'étape du clic suivant.
    """
    events = []
    ignored = 0
    step = 0
    pending_next_click_events = []
    all_at_once_shapes = all_at_once_build_shape_ids(slide_root)

    effect_ctns = slide_root.xpath(
        (
            ".//p:timing//p:cTn"
            "[@nodeType='clickEffect' or @nodeType='withEffect' or @nodeType='afterEffect']"
        ),
        namespaces=NS,
    )

    for ctn in effect_ctns:
        node_type = ctn.get("nodeType")
        
        if is_synthetic_all_at_once_visibility_setup(ctn):
            continue

        unsupported_text_kind = unsupported_text_animation_kind(ctn)

        if unsupported_text_kind is not None:
            key = f"unsupported_text_by_{unsupported_text_kind}"

            if report is not None and key in report:
                report[key] += 1

            ignored += 1
            continue

        visibility_action = visibility_action_from_effect_container(ctn)

        has_supported_event = False
        step_for_this_ctn = None

        def ensure_step_for_ctn():
            nonlocal step
            nonlocal step_for_this_ctn
            nonlocal pending_next_click_events
            nonlocal has_supported_event

            if step_for_this_ctn is None:
                step = step_for_effect(node_type, ctn, step)
                step_for_this_ctn = step

                if node_type == "clickEffect" and pending_next_click_events:
                    for pending_event in pending_next_click_events:
                        event = dict(pending_event)
                        event["step"] = step_for_this_ctn
                        events.append(event)

                    pending_next_click_events = []
                    has_supported_event = True

            return step_for_this_ctn

        # Apparitions / disparitions
        if visibility_action is not None:
            targets = []

            for cbhvr in ctn.xpath(
                ".//p:cBhvr[not(ancestor::p:subTnLst)]",
                namespaces=NS,
            ):
                targets.extend(target_from_behavior(cbhvr))

            targets = dedup_targets(targets)

            # Dans les listes build="allAtOnce", PowerPoint ajoute des
            # style.visibility=visible par paragraphe, mais ces paragraphes sont
            # déjà visibles dès l'état initial. Ces shows ne doivent donc pas
            # déclencher build_initial_visibility().
            visible_targets = []

            for target in targets:
                # Si un événement de visibilité cible un paragraphe appartenant
                # à une forme construite all-at-once.
                target_is_all_at_once_paragraph = (
                    target[0] == "paragraph"
                    and target[1] in all_at_once_shapes
                )

                if visibility_action == "show" and target_is_all_at_once_paragraph:
                    continue

                visible_targets.append(target)

            if visible_targets:
                current_step = ensure_step_for_ctn()

                for target in visible_targets:
                    events.append(
                        {
                            "step": current_step,
                            "target": target,
                            "action": visibility_action,
                        }
                    )

                has_supported_event = True

        # Transformations normales dans l'ordre XML réel.
        for anim_node in ordered_transform_nodes_from_container(ctn):
            tag = anim_node.tag

            if tag == qn(P, "animMotion"):
                delta, skip_reason = motion_delta_from_anim_motion(
                    anim_node,
                    slide_width,
                    slide_height,
                )

                if delta is None:
                    if report is not None and skip_reason in report:
                        report[skip_reason] += 1
                    continue

                targets = []

                for cbhvr in anim_node.xpath(".//p:cBhvr", namespaces=NS):
                    targets.extend(target_from_behavior(cbhvr))

                targets = dedup_targets(targets)

                if not targets:
                    continue

                current_step = ensure_step_for_ctn()
                dx, dy = delta

                for target in targets:
                    if target[0] == "paragraph":
                        target = ("shape", target[1])

                    events.append(
                        {
                            "step": current_step,
                            "target": target,
                            "action": "move",
                            "dx": dx,
                            "dy": dy,
                        }
                    )

                has_supported_event = True
                continue

            if tag == qn(P, "animScale"):
                scale_factor, skip_reason = scale_factor_from_anim_scale(anim_node)

                if scale_factor is None:
                    if report is not None and skip_reason in report:
                        report[skip_reason] += 1
                    continue

                targets = []

                for cbhvr in anim_node.xpath(".//p:cBhvr", namespaces=NS):
                    targets.extend(target_from_behavior(cbhvr))

                targets = dedup_targets(targets)

                if not targets:
                    continue

                current_step = ensure_step_for_ctn()
                sx, sy = scale_factor

                for target in targets:
                    if target[0] == "paragraph":
                        target = ("shape", target[1])

                    events.append(
                        {
                            "step": current_step,
                            "target": target,
                            "action": "scale",
                            "sx": sx,
                            "sy": sy,
                        }
                    )

                has_supported_event = True
                continue

            if tag == qn(P, "animRot"):
                rotation_transform, skip_reason = rotation_transform_from_anim_rot(anim_node)

                if rotation_transform is None:
                    if report is not None and skip_reason in report:
                        report[skip_reason] += 1
                    continue

                targets = []

                for cbhvr in anim_node.xpath(".//p:cBhvr", namespaces=NS):
                    targets.extend(target_from_behavior(cbhvr))

                targets = dedup_targets(targets)

                if not targets:
                    continue

                current_step = ensure_step_for_ctn()

                for target in targets:
                    if target[0] == "paragraph":
                        target = ("shape", target[1])

                    events.append(
                        {
                            "step": current_step,
                            "target": target,
                            "action": "rotate",
                            "rotation_mode": rotation_transform["rotation_mode"],
                            "rot": rotation_transform["rot"],
                        }
                    )

                has_supported_event = True
                continue

            if tag == qn(P, "animClr"):
                color_transition, skip_reason = color_transition_from_anim_clr(anim_node)

                if color_transition is None:
                    if report is not None and skip_reason in report:
                        report[skip_reason] += 1
                    continue

                targets = []

                for cbhvr in anim_node.xpath(".//p:cBhvr", namespaces=NS):
                    targets.extend(target_from_behavior(cbhvr))

                targets = dedup_targets(targets)

                if not targets:
                    continue

                current_step = ensure_step_for_ctn()

                for target in targets:
                    events.append(
                        {
                            "step": current_step,
                            "target": target,
                            "action": "color_transition",
                            "color_kind": color_transition["color_kind"],
                            "from_color": color_transition["from_color"],
                            "to_color": color_transition["to_color"],
                        }
                    )

                has_supported_event = True
                continue

            if tag in {qn(P, "anim"), qn(P, "set")}:
                alpha, skip_reason = opacity_alpha_from_anim(anim_node)

                if alpha is not None:
                    targets = []

                    for cbhvr in anim_node.xpath(".//p:cBhvr", namespaces=NS):
                        targets.extend(target_from_behavior(cbhvr))

                    targets = dedup_targets(targets)

                    if not targets:
                        continue

                    current_step = ensure_step_for_ctn()

                    for target in targets:
                        if target[0] == "paragraph":
                            target = ("shape", target[1])

                        events.append(
                            {
                                "step": current_step,
                                "target": target,
                                "action": "opacity",
                                "alpha": alpha,
                            }
                        )

                    has_supported_event = True
                    continue

                if skip_reason is not None:
                    if report is not None and skip_reason in report:
                        report[skip_reason] += 1
                    continue

                text_style_change, skip_reason = text_style_change_from_anim(anim_node)

                if text_style_change is None:
                    if report is not None and skip_reason in report:
                        report[skip_reason] += 1
                    continue

                targets = []

                for cbhvr in anim_node.xpath(".//p:cBhvr", namespaces=NS):
                    targets.extend(target_from_behavior(cbhvr))

                targets = dedup_targets(targets)

                if not targets:
                    continue

                current_step = ensure_step_for_ctn()

                for target in targets:
                    events.append(
                        {
                            "step": current_step,
                            "target": target,
                            "action": "text_style",
                            "text_style": text_style_change["text_style"],
                            "value": text_style_change["value"],
                        }
                    )

                has_supported_event = True
                continue

        # Effets post-animation : Dim after animation, etc
        after_animation_groups = after_animation_events_from_container(ctn)

        if after_animation_groups:
            # Même si l'événement principal est un faux show ignoré à cause de
            # build="allAtOnce", le conteneur clickEffect représente quand même
            # un clic utilisateur. Il doit donc appliquer les dims en attente.
            if node_type == "clickEffect" and pending_next_click_events:
                ensure_step_for_ctn()

            for master_relation, after_events in after_animation_groups:
                if master_relation == "nextClick":
                    pending_next_click_events.extend(after_events)
                    has_supported_event = True
                    continue

                if step_for_this_ctn is None:
                    ensure_step_for_ctn()

                post_step = step_for_this_ctn + 1

                for after_event in after_events:
                    event = dict(after_event)
                    event["step"] = post_step
                    events.append(event)

                step = max(step, post_step)
                has_supported_event = True

        if not has_supported_event:
            ignored += 1

    return events, ignored

def set_transparent_text_fill(rpr):
    fill_tags = {
        qn(A, "noFill"),
        qn(A, "solidFill"),
        qn(A, "gradFill"),
        qn(A, "blipFill"),
        qn(A, "pattFill"),
        qn(A, "grpFill"),
    }

    for child in list(rpr):
        if child.tag in fill_tags:
            rpr.remove(child)

    solid = etree.Element(qn(A, "solidFill"))

    color = etree.SubElement(solid, qn(A, "srgbClr"))
    color.set("val", "FFFFFF")

    alpha = etree.SubElement(color, qn(A, "alpha"))
    alpha.set("val", "0")

    insert_before = {
        qn(A, "effectLst"),
        qn(A, "effectDag"),
        qn(A, "highlight"),
        qn(A, "uLnTx"),
        qn(A, "uLn"),
        qn(A, "uFillTx"),
        qn(A, "uFill"),
        qn(A, "latin"),
        qn(A, "ea"),
        qn(A, "cs"),
        qn(A, "sym"),
        qn(A, "hlinkClick"),
        qn(A, "hlinkMouseOver"),
        qn(A, "rtl"),
        qn(A, "extLst"),
    }

    for i, child in enumerate(list(rpr)):
        if child.tag in insert_before:
            rpr.insert(i, solid)
            return

    rpr.append(solid)


def ensure_run_properties(run_el):
    rpr = run_el.find("./a:rPr", namespaces=NS)

    if rpr is None:
        rpr = etree.Element(qn(A, "rPr"))
        run_el.insert(0, rpr)

    return rpr


def hide_bullet_for_paragraph(p_el):
    ppr = p_el.find("./a:pPr", namespaces=NS)

    if ppr is None:
        ppr = etree.Element(qn(A, "pPr"))
        p_el.insert(0, ppr)

    for child in list(ppr):
        if child.tag in {qn(A, "buClr"), qn(A, "buClrTx")}:
            ppr.remove(child)

    bu_clr = etree.Element(qn(A, "buClr"))

    color = etree.SubElement(bu_clr, qn(A, "srgbClr"))
    color.set("val", "FFFFFF")

    alpha = etree.SubElement(color, qn(A, "alpha"))
    alpha.set("val", "0")

    insert_before = {
        qn(A, "buSzTx"),
        qn(A, "buSzPct"),
        qn(A, "buSzPts"),
        qn(A, "buFontTx"),
        qn(A, "buFont"),
        qn(A, "buNone"),
        qn(A, "buAutoNum"),
        qn(A, "buChar"),
        qn(A, "buBlip"),
        qn(A, "tabLst"),
        qn(A, "defRPr"),
        qn(A, "extLst"),
    }

    for i, child in enumerate(list(ppr)):
        if child.tag in insert_before:
            ppr.insert(i, bu_clr)
            return

    ppr.append(bu_clr)


def hide_paragraph_but_keep_layout(p_el):
    hide_bullet_for_paragraph(p_el)

    for run in p_el.xpath(".//a:r | .//a:fld", namespaces=NS):
        ensure_run_properties(run)

    for rpr in p_el.xpath(".//a:rPr", namespaces=NS):
        set_transparent_text_fill(rpr)

    end = p_el.find("./a:endParaRPr", namespaces=NS)
    if end is not None:
        set_transparent_text_fill(end)


def build_initial_visibility(slide_root, events):
    shapes = all_shape_elements(slide_root)

    shape_visible = {sid: True for sid in shapes}
    paragraph_visible = {}
    paragraph_animated_shapes = set()

    for event in events:
        target = event["target"]

        if target[0] == "paragraph":
            paragraph_animated_shapes.add(target[1])

    for event in events:
        action = event["action"]
        target = event["target"]

        if action != "show":
            continue

        if target[0] == "shape":
            _, sid = target
            shape_visible[sid] = False

        elif target[0] == "paragraph":
            _, sid, pidx = target
            paragraph_visible[(sid, pidx)] = False
            shape_visible.setdefault(sid, True)

    return shape_visible, paragraph_visible, paragraph_animated_shapes

def apply_visibility_to_slide(
    slide_root,
    shape_visible,
    paragraph_visible,
    paragraph_animated_shapes=None,
):
    if paragraph_animated_shapes is None:
        paragraph_animated_shapes = set()

    shapes = all_shape_elements(slide_root)

    for sid, shape in list(shapes.items()):
        has_para_anim = sid in paragraph_animated_shapes

        if not shape_visible.get(sid, True):
            if has_para_anim:
                for p in paragraphs_of_shape(shape):
                    hide_paragraph_but_keep_layout(p)

                continue

            parent = shape.getparent()
            if parent is not None:
                parent.remove(shape)

            continue

        for idx, p in enumerate(paragraphs_of_shape(shape)):
            if not paragraph_visible.get((sid, idx), True):
                hide_paragraph_but_keep_layout(p)

def shape_transform_element(shape):
    """
    Retourne l'élément de transformation direct d'une forme.

    Cas couverts :
    - p:sp / p:spPr / a:xfrm
    - p:pic / p:spPr / a:xfrm
    - p:cxnSp / p:spPr / a:xfrm
    - p:grpSp / p:grpSpPr / a:xfrm
    - p:graphicFrame / p:xfrm
    """
    for path in [
        "./p:spPr/a:xfrm",
        "./p:grpSpPr/a:xfrm",
        "./p:xfrm",
    ]:
        xfrm = shape.find(path, namespaces=NS)
        if xfrm is not None:
            return xfrm

    return None


def apply_shape_offset(shape, dx, dy):
    xfrm = shape_transform_element(shape)

    if xfrm is None:
        return False

    off = xfrm.find("./a:off", namespaces=NS)

    if off is None:
        off = etree.Element(qn(A, "off"))
        off.set("x", "0")
        off.set("y", "0")
        xfrm.insert(0, off)

    x = int(off.get("x", "0"))
    y = int(off.get("y", "0"))

    off.set("x", str(x + dx))
    off.set("y", str(y + dy))

    return True

def ensure_transform_parts(shape):
    """
    Retourne (xfrm, off, ext) pour une forme.

    Crée a:off si absent.
    Ne crée pas a:ext : sans taille existante, on ne sait pas scaler proprement.
    """
    xfrm = shape_transform_element(shape)

    if xfrm is None:
        return None, None, None

    off = xfrm.find("./a:off", namespaces=NS)

    if off is None:
        off = etree.Element(qn(A, "off"))
        off.set("x", "0")
        off.set("y", "0")
        xfrm.insert(0, off)

    ext = xfrm.find("./a:ext", namespaces=NS)

    if ext is None:
        return xfrm, off, None

    return xfrm, off, ext


def apply_shape_scale(shape, sx, sy):
    """
    Applique un Grow/Shrink en conservant le centre de l'objet.

    Les copies statiques repartent toujours du XML original de la slide ;
    sx/sy doivent donc être des facteurs cumulés depuis l'état initial.
    """
    _, off, ext = ensure_transform_parts(shape)

    if off is None or ext is None:
        return False

    x = int(off.get("x", "0"))
    y = int(off.get("y", "0"))
    cx = int(ext.get("cx", "0"))
    cy = int(ext.get("cy", "0"))

    if cx <= 0 or cy <= 0:
        return False

    new_cx = int(round(cx * sx))
    new_cy = int(round(cy * sy))

    new_x = int(round(x - (new_cx - cx) / 2))
    new_y = int(round(y - (new_cy - cy) / 2))

    off.set("x", str(new_x))
    off.set("y", str(new_y))
    ext.set("cx", str(new_cx))
    ext.set("cy", str(new_cy))

    return True

def apply_shape_rotation(shape, rot, rotation_mode="delta"):
    """
    Applique une rotation à une forme.

    rotation_mode="delta" :
        ajoute rot à la rotation existante.

    rotation_mode="absolute" :
        fixe la rotation finale à rot.

    La rotation DrawingML se fait autour du centre de la boîte englobante.
    """
    xfrm = shape_transform_element(shape)

    if xfrm is None:
        return False

    current_rot = int(xfrm.get("rot", "0"))

    if rotation_mode == "delta":
        new_rot = current_rot + rot
    elif rotation_mode == "absolute":
        new_rot = rot
    else:
        return False

    new_rot = normalize_rotation(new_rot)

    if normalize_rotation(current_rot) == new_rot:
        return False

    xfrm.set("rot", str(new_rot))

    return True

def set_alpha_on_color(color_el, alpha):
    """
    Force l'opacité d'une couleur DrawingML.
    """
    for child in list(color_el):
        if child.tag in {
            qn(A, "alpha"),
            qn(A, "alphaMod"),
            qn(A, "alphaOff"),
        }:
            color_el.remove(child)

    alpha_el = etree.Element(qn(A, "alpha"))
    alpha_el.set("val", str(clamp_alpha(alpha)))
    color_el.append(alpha_el)


def apply_alpha_to_solid_fill(solid_fill, alpha):
    for child in solid_fill:
        if child.tag in COLOR_TAGS:
            set_alpha_on_color(child, alpha)
            return True
    
    return False

def apply_alpha_to_style_refs(shape, alpha):
    """
    Applique l'opacité aux couleurs héritées via p:style.

    Couvre notamment les formes dont le remplissage ou le contour vient
    d'un fillRef / lnRef au lieu d'un a:solidFill explicite.
    """
    applied = False

    for color in shape.xpath(
        ".//p:style/a:fillRef/* | "
        ".//p:style/a:lnRef/* | "
        ".//p:style/a:fontRef/*",
        namespaces=NS,
    ):
        if color.tag in COLOR_TAGS:
            set_alpha_on_color(color, alpha)
            applied = True

    return applied

def apply_alpha_to_blip(shape, alpha):
    """
    Applique l'opacité aux images via a:alphaModFix.

    Couvre les p:pic et les remplissages image simples.
    """
    applied = False

    for blip in shape.xpath(".//a:blip", namespaces=NS):
        for child in list(blip):
            if child.tag in {qn(A, "alphaModFix"), qn(A, "alphaMod")}:
                blip.remove(child)

        alpha_mod = etree.Element(qn(A, "alphaModFix"))
        alpha_mod.set("amt", str(clamp_alpha(alpha)))
        blip.append(alpha_mod)

        applied = True

    return applied

def apply_shape_opacity(shape, alpha):
    """
    Applique une opacité finale à une forme.

    On cible les cas statiques les plus utiles :
    - remplissage de forme explicite ;
    - contour explicite ;
    - texte avec couleur explicite ;
    - images via a:blip ;
    - couleurs héritées via p:style / fillRef / lnRef / fontRef.

    Pour les groupes, les recherches .// parcourent aussi les formes enfants.
    """
    alpha = clamp_alpha(alpha)
    applied = False

    solid_fill_paths = [
        ".//p:spPr/a:solidFill | .//p:grpSpPr/a:solidFill",
        ".//p:spPr/a:ln/a:solidFill | .//p:grpSpPr/a:ln/a:solidFill",
        ".//a:rPr/a:solidFill",
    ]

    for path in solid_fill_paths:
        for solid_fill in shape.xpath(path, namespaces=NS):
            if apply_alpha_to_solid_fill(solid_fill, alpha):
                applied = True

    if apply_alpha_to_style_refs(shape, alpha):
        applied = True

    if apply_alpha_to_blip(shape, alpha):
        applied = True

    return applied

def set_text_style_on_rpr(rpr, text_style, value):
    style_attrs = {
        "bold": ("b", "1", "0"),
        "italic": ("i", "1", "0"),
        "underline": ("u", "sng", "none"),
        "strike": ("strike", "sngStrike", "noStrike"),
    }

    spec = style_attrs.get(text_style)

    if spec is None:
        return False

    attr_name, true_value, false_value = spec
    rpr.set(attr_name, true_value if value else false_value)

    return True

def apply_text_style_to_paragraph(p_el, text_style, value):
    applied = False

    for run in p_el.xpath(".//a:r | .//a:fld", namespaces=NS):
        ensure_run_properties(run)

    for rpr in p_el.xpath(".//a:rPr", namespaces=NS):
        if set_text_style_on_rpr(rpr, text_style, value):
            applied = True

    end = p_el.find("./a:endParaRPr", namespaces=NS)

    if end is not None:
        if set_text_style_on_rpr(end, text_style, value):
            applied = True

    return applied


def apply_shape_text_style(shape, text_style, value, paragraph_index=None):
    paragraphs = paragraphs_of_shape(shape)

    if paragraph_index is not None:
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            return False

        return apply_text_style_to_paragraph(
            paragraphs[paragraph_index],
            text_style,
            value,
        )

    applied = False

    for paragraph in paragraphs:
        if apply_text_style_to_paragraph(paragraph, text_style, value):
            applied = True

    return applied

def color_key(color_el):
    """
    Clé normalisée pour comparer deux couleurs.

    On compare le type de couleur DrawingML et sa valeur principale.
    Exemple :
      <a:srgbClr val="FF0000"/>    -> ("srgbClr", "FF0000")
      <a:schemeClr val="accent1"/> -> ("schemeClr", "accent1")

    La valeur spéciale "none" représente l'absence de couleur.
    """
    if color_el is None:
        return ("unknown", None)

    if color_el == "none":
        return ("none", None)

    q = etree.QName(color_el)
    return (q.localname, color_el.get("val"))

def replace_fill_child(parent, color_el, insert_before_tags=None):
    if insert_before_tags is None:
        insert_before_tags = set()

    for child in list(parent):
        if child.tag in FILL_TAGS:
            parent.remove(child)

    solid_fill = etree.Element(qn(A, "solidFill"))
    solid_fill.append(copy.deepcopy(color_el))


    for i, child in enumerate(list(parent)):
        if child.tag in insert_before_tags:
            parent.insert(i, solid_fill)
            return True

    parent.append(solid_fill)
    return True


def apply_fill_color_to_shape(shape, color_el):
    applied = False

    for sp_pr in shape.xpath(".//p:spPr | .//p:grpSpPr", namespaces=NS):
        if replace_fill_child(
            sp_pr,
            color_el,
            insert_before_tags={
                qn(A, "ln"),
                qn(A, "effectLst"),
                qn(A, "effectDag"),
                qn(A, "scene3d"),
                qn(A, "sp3d"),
                qn(A, "extLst"),
            },
        ):
            applied = True

    return applied


def ensure_line_properties(sp_pr):
    ln = sp_pr.find("./a:ln", namespaces=NS)

    if ln is not None:
        return ln

    ln = etree.Element(qn(A, "ln"))

    insert_before_tags = {
        qn(A, "effectLst"),
        qn(A, "effectDag"),
        qn(A, "scene3d"),
        qn(A, "sp3d"),
        qn(A, "extLst"),
    }

    for i, child in enumerate(list(sp_pr)):
        if child.tag in insert_before_tags:
            sp_pr.insert(i, ln)
            return ln

    sp_pr.append(ln)
    return ln

def apply_line_color_to_shape(shape, color_el):
    applied = False

    for sp_pr in shape.xpath(".//p:spPr | .//p:grpSpPr", namespaces=NS):
        ln = ensure_line_properties(sp_pr)

        if replace_fill_child(
            ln,
            color_el,
            insert_before_tags={
                qn(A, "prstDash"),
                qn(A, "custDash"),
                qn(A, "round"),
                qn(A, "bevel"),
                qn(A, "miter"),
                qn(A, "headEnd"),
                qn(A, "tailEnd"),
                qn(A, "extLst"),
            },
        ):
            applied = True

    return applied

def apply_text_color_to_paragraph(p_el, color_el):
    applied = False

    for run in p_el.xpath(".//a:r | .//a:fld", namespaces=NS):
        rpr = ensure_run_properties(run)

        if replace_fill_child(
            rpr,
            color_el,
            insert_before_tags={
                qn(A, "effectLst"),
                qn(A, "effectDag"),
                qn(A, "highlight"),
                qn(A, "uLnTx"),
                qn(A, "uLn"),
                qn(A, "uFillTx"),
                qn(A, "uFill"),
                qn(A, "latin"),
                qn(A, "ea"),
                qn(A, "cs"),
                qn(A, "sym"),
                qn(A, "hlinkClick"),
                qn(A, "hlinkMouseOver"),
                qn(A, "rtl"),
                qn(A, "extLst"),
            },
        ):
            applied = True

    end = p_el.find("./a:endParaRPr", namespaces=NS)

    if end is not None:
        if replace_fill_child(
            end,
            color_el,
            insert_before_tags={
                qn(A, "effectLst"),
                qn(A, "effectDag"),
                qn(A, "highlight"),
                qn(A, "latin"),
                qn(A, "ea"),
                qn(A, "cs"),
                qn(A, "sym"),
                qn(A, "extLst"),
            },
        ):
            applied = True

    return applied


def apply_text_color_to_shape(shape, color_el, paragraph_index=None):
    paragraphs = paragraphs_of_shape(shape)

    if paragraph_index is not None:
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            return False

        return apply_text_color_to_paragraph(
            paragraphs[paragraph_index],
            color_el,
        )

    applied = False

    for paragraph in paragraphs:
        if apply_text_color_to_paragraph(paragraph, color_el):
            applied = True

    return applied


def apply_shape_color(shape, color_kind, color_el, paragraph_index=None):
    if color_el is None:
        return False

    if color_kind == "fill":
        return apply_fill_color_to_shape(shape, color_el)

    if color_kind == "line":
        return apply_line_color_to_shape(shape, color_el)

    if color_kind == "text":
        return apply_text_color_to_shape(
            shape,
            color_el,
            paragraph_index,
        )

    return False

def direct_color_key_from_path(shape, path):
    colors = shape.xpath(path, namespaces=NS)

    if not colors:
        return ("unknown", None)

    return color_key(colors[0])


def current_fill_color_key(shape):
    if shape.xpath("./p:spPr/a:noFill | ./p:grpSpPr/a:noFill", namespaces=NS):
        return ("none", None)

    key = direct_color_key_from_path(
        shape,
        "./p:spPr/a:solidFill/* | ./p:grpSpPr/a:solidFill/*",
    )

    if key[0] != "unknown":
        return key

    key = direct_color_key_from_path(
        shape,
        "./p:style/a:fillRef/*",
    )

    return key


def current_line_color_key(shape):
    if shape.xpath("./p:spPr/a:ln/a:noFill | ./p:grpSpPr/a:ln/a:noFill", namespaces=NS):
        return ("none", None)

    key = direct_color_key_from_path(
        shape,
        "./p:spPr/a:ln/a:solidFill/* | ./p:grpSpPr/a:ln/a:solidFill/*",
    )

    if key[0] != "unknown":
        return key

    key = direct_color_key_from_path(
        shape,
        "./p:style/a:lnRef/*",
    )

    return key


def current_text_color_key(shape):
    key = direct_color_key_from_path(
        shape,
        ".//a:rPr/a:solidFill/*",
    )

    if key[0] != "unknown":
        return key

    key = direct_color_key_from_path(
        shape,
        "./p:style/a:fontRef/*",
    )

    return key

def text_style_value_from_rpr(rpr, text_style):
    if rpr is None:
        return False

    if text_style == "bold":
        return str(rpr.get("b", "0")).lower() in {"1", "true", "t"}

    if text_style == "italic":
        return str(rpr.get("i", "0")).lower() in {"1", "true", "t"}

    if text_style == "underline":
        underline = rpr.get("u")
        return underline is not None and underline != "none"

    if text_style == "strike":
        strike = rpr.get("strike")
        return strike is not None and strike != "noStrike"

    return "unknown"


def paragraph_text_style_state(p_el, text_style):
    runs = p_el.xpath(".//a:r | .//a:fld", namespaces=NS)

    if not runs:
        return False

    values = []

    for run in runs:
        rpr = run.find("./a:rPr", namespaces=NS)
        values.append(text_style_value_from_rpr(rpr, text_style))

    first = values[0]

    if all(value == first for value in values):
        return first

    return "unknown"


def shape_text_style_state(shape, text_style):
    paragraphs = paragraphs_of_shape(shape)

    if not paragraphs:
        return False

    values = [
        paragraph_text_style_state(p, text_style)
        for p in paragraphs
    ]

    first = values[0]

    if all(value == first for value in values):
        return first

    return "unknown"


def build_text_style_state_for_shape(shape):
    paragraphs = paragraphs_of_shape(shape)

    paragraph_states = {}

    for idx, paragraph in enumerate(paragraphs):
        paragraph_states[idx] = {
            text_style: paragraph_text_style_state(paragraph, text_style)
            for text_style in TEXT_STYLE_KINDS
        }

    shape_state = {
        text_style: shape_text_style_state(shape, text_style)
        for text_style in TEXT_STYLE_KINDS
    }

    return {
        "shape": shape_state,
        "paragraphs": paragraph_states,
    }


def build_initial_style_state(slide_root):
    """
    État courant des couleurs et des styles texte par forme.

    Les couleurs servent aux animations Fill/Line/Font Color.
    Les styles texte servent à éviter des slides redondantes pour
    Bold / Italic / Underline / Strikethrough.
    """
    state = {}
    shapes = all_shape_elements(slide_root)

    for sid, shape in shapes.items():
        paragraphs = paragraphs_of_shape(shape)

        state[sid] = {
            "fill": current_fill_color_key(shape),
            "line": current_line_color_key(shape),
            "text": current_text_color_key(shape),
            "paragraph_text_colors": {
                idx: direct_color_key_from_path(
                    paragraph,
                    ".//a:rPr/a:solidFill/*",
                )
                for idx, paragraph in enumerate(paragraphs)
            },
            "text_styles": build_text_style_state_for_shape(shape),
        }

    return state

def style_state_key(shape_style_state, sid, color_kind):
    return shape_style_state.get(sid, {}).get(color_kind, ("unknown", None))

def set_style_state_key(shape_style_state, sid, color_kind, key):
    shape_style_state.setdefault(sid, {})[color_kind] = key

def text_style_state_value(shape_style_state, target, text_style):
    if target[0] == "shape":
        _, sid = target

        return (
            shape_style_state
            .get(sid, {})
            .get("text_styles", {})
            .get("shape", {})
            .get(text_style, "unknown")
        )

    if target[0] == "paragraph":
        _, sid, pidx = target

        return (
            shape_style_state
            .get(sid, {})
            .get("text_styles", {})
            .get("paragraphs", {})
            .get(pidx, {})
            .get(text_style, "unknown")
        )

    return "unknown"

def set_text_style_state_value(shape_style_state, target, text_style, value):
    if target[0] == "shape":
        _, sid = target

        text_styles = shape_style_state.setdefault(sid, {}).setdefault(
            "text_styles",
            {"shape": {}, "paragraphs": {}},
        )

        text_styles.setdefault("shape", {})[text_style] = value

        for paragraph_state in text_styles.setdefault("paragraphs", {}).values():
            paragraph_state[text_style] = value

        return

    if target[0] == "paragraph":
        _, sid, pidx = target

        text_styles = shape_style_state.setdefault(sid, {}).setdefault(
            "text_styles",
            {"shape": {}, "paragraphs": {}},
        )

        paragraphs = text_styles.setdefault("paragraphs", {})
        paragraphs.setdefault(pidx, {})[text_style] = value

        values = [
            paragraph_state.get(text_style, "unknown")
            for paragraph_state in paragraphs.values()
        ]

        if values and all(v == values[0] for v in values):
            text_styles.setdefault("shape", {})[text_style] = values[0]
        else:
            text_styles.setdefault("shape", {})[text_style] = "unknown"

def color_state_key_for_target(shape_style_state, target, color_kind):
    if target[0] == "paragraph" and color_kind == "text":
        _, sid, pidx = target

        return (
            shape_style_state
            .get(sid, {})
            .get("paragraph_text_colors", {})
            .get(pidx, shape_style_state.get(sid, {}).get("text", ("unknown", None)))
        )

    _, sid = target[:2]

    return shape_style_state.get(sid, {}).get(color_kind, ("unknown", None))


def set_color_state_key_for_target(shape_style_state, target, color_kind, key):
    if target[0] == "paragraph" and color_kind == "text":
        _, sid, pidx = target

        paragraph_colors = shape_style_state.setdefault(sid, {}).setdefault(
            "paragraph_text_colors",
            {},
        )

        paragraph_colors[pidx] = key

        values = list(paragraph_colors.values())

        if values and all(value == values[0] for value in values):
            shape_style_state[sid]["text"] = values[0]
        else:
            shape_style_state[sid]["text"] = ("unknown", None)

        return

    _, sid = target[:2]
    shape_style_state.setdefault(sid, {})[color_kind] = key

def append_shape_transform(shape_transforms, sid, transform):
    """
    Ajoute une transformation à la liste ordonnée d'une forme.

    On conserve l'ordre chronologique des effets PowerPoint.
    Les actions actuellement utilisées sont move, scale, rotate et opacity.
    """
    shape_transforms.setdefault(sid, []).append(transform)


def apply_shape_transform(shape, transform):
    """
    Applique une transformation élémentaire à une forme.

    Cette fonction centralise l'application des transformations et permet
    d'ajouter de nouveaux effets sans modifier la boucle principale.
    """
    action = transform["action"]

    if action == "move":
        return apply_shape_offset(
            shape,
            transform["dx"],
            transform["dy"],
        )

    if action == "scale":
        return apply_shape_scale(
            shape,
            transform["sx"],
            transform["sy"],
        )

    if action == "rotate":
        return apply_shape_rotation(
            shape,
            transform["rot"],
            transform.get("rotation_mode", "delta"),
        )

    if action == "opacity":
        return apply_shape_opacity(
            shape,
            transform["alpha"],
        )

    if action == "color":
        return apply_shape_color(
            shape,
            transform["color_kind"],
            transform["color"],
            transform.get("paragraph_index"),
        )

    if action == "text_style":
        return apply_shape_text_style(
            shape,
            transform["text_style"],
            transform["value"],
            transform.get("paragraph_index"),
        )

    return False

def apply_transforms_to_slide(slide_root, shape_transforms):
    """
    Applique les transformations géométriques cumulées à une copie de slide.

    Contrairement à l'ancien modèle shape_offsets / shape_scales, on garde
    ici l'ordre des transformations. Cela ne change pas le comportement actuel
    pour move + scale centré, mais prépare les transformations non commutatives.
    """
    if not shape_transforms:
        return

    shapes = all_shape_elements(slide_root)

    for sid, transforms in shape_transforms.items():
        shape = shapes.get(sid)

        if shape is None:
            continue
        #Applique les transformations d'une forme dans l'ordre de la timeline
        for transform in transforms:
            apply_shape_transform(shape, transform)

def color_transition_needs_from_slide(event, shape_style_state):
    from_color = event.get("from_color")

    if from_color is None:
        return False

    _, sid = event["target"]
    color_kind = event["color_kind"]

    current_key = style_state_key(shape_style_state, sid, color_kind)
    from_key = color_key(from_color)

    # Si l'état courant est inconnu, on est conservateur :
    # on garde l'état from pour ne pas rater une étape visible.
    if current_key[0] == "unknown":
        return True

    return current_key != from_key


def event_changes_visible_state(
    event,
    shape_visible,
    paragraph_visible,
    shape_style_state=None,
    color_post_mode=False,
):
    """
    Indique si un événement justifie d'émettre une nouvelle slide.

    Un move / scale / rotate / opacity appliqué à un objet invisible est
    mémorisé, mais ne produit pas de slide visible immédiatement.
    """
    action = event["action"]
    target = event["target"]

    if action == "move":
        _, sid = target

        if event.get("dx", 0) == 0 and event.get("dy", 0) == 0:
            return False

        return shape_visible.get(sid, True)

    if action == "scale":
        _, sid = target

        sx = event.get("sx", 1.0)
        sy = event.get("sy", 1.0)

        if abs(sx - 1.0) < 1e-9 and abs(sy - 1.0) < 1e-9:
            return False

        return shape_visible.get(sid, True)

    if action == "rotate":
        _, sid = target

        rot = event.get("rot", 0)

        if normalize_rotation(rot) == 0 and event.get("rotation_mode", "delta") == "delta":
            return False

        return shape_visible.get(sid, True)

    if action == "opacity":
        _, sid = target

        alpha = event.get("alpha")

        if alpha is None:
            return False

        return shape_visible.get(sid, True)

    if action == "color":
        color_kind = event["color_kind"]

        if target[0] == "shape":
            _, sid = target

            if not shape_visible.get(sid, True):
                return False

        elif target[0] == "paragraph":
            _, sid, pidx = target

            if not shape_visible.get(sid, True):
                return False

            if not paragraph_visible.get((sid, pidx), True):
                return False

        if shape_style_state is None:
            return True

        current_key = color_state_key_for_target(
            shape_style_state,
            target,
            color_kind,
        )

        new_key = color_key(event["color"])

        return current_key != new_key

    if action == "color_transition":
        _, sid = target

        if shape_style_state is None:
            return shape_visible.get(sid, True)

        if color_post_mode:
            return (
                shape_visible.get(sid, True)
                and color_transition_needs_from_slide(event, shape_style_state)
            )

        current_key = style_state_key(
            shape_style_state,
            sid,
            event["color_kind"],
        )

        to_key = color_key(event["to_color"])

        return shape_visible.get(sid, True) and current_key != to_key

    if action == "text_style":
        text_style = event["text_style"]
        new_value = event["value"]

        if shape_style_state is None:
            if target[0] == "shape":
                _, sid = target
                return shape_visible.get(sid, True)

            if target[0] == "paragraph":
                _, sid, pidx = target
                return (
                    shape_visible.get(sid, True)
                    and paragraph_visible.get((sid, pidx), True)
                )

        if target[0] == "shape":
            _, sid = target

            if not shape_visible.get(sid, True):
                return False

        elif target[0] == "paragraph":
            _, sid, pidx = target

            if not shape_visible.get(sid, True):
                return False

            if not paragraph_visible.get((sid, pidx), True):
                return False

        current_value = text_style_state_value(
            shape_style_state,
            target,
            text_style,
        )

        if current_value == "unknown":
            return True

        return current_value != new_value

    if action in {"show", "hide"}:
        value = action == "show"

        if target[0] == "shape":
            _, sid = target
            return shape_visible.get(sid, True) != value

        if target[0] == "paragraph":
            _, sid, pidx = target

            if not shape_visible.get(sid, True):
                return False

            return paragraph_visible.get((sid, pidx), True) != value

    return False

def clean_slide_rels(rels_data):
    if rels_data is None:
        root = etree.Element(qn(REL, "Relationships"))
        return xml_bytes(root)

    root = parse_xml(rels_data)

    for rel in list(root):
        rel_type = rel.get("Type", "")

        if rel_type.endswith(("/notesSlide", "/comments", "/commentAuthors")):
            root.remove(rel)

    return xml_bytes(root)


def next_free_rid(used):
    i = 1

    while True:
        rid = f"rId{i}"

        if rid not in used:
            used.add(rid)
            return rid

        i += 1

def presentation_slide_size(pres_root):
    sld_sz = pres_root.find("./p:sldSz", namespaces=NS)

    if sld_sz is None:
        # Taille widescreen PowerPoint par défaut : 13.333 x 7.5 pouces.
        return 12192000, 6858000

    return int(sld_sz.get("cx")), int(sld_sz.get("cy"))

def new_conversion_report():
    return {
        "closed_or_zero_motion_paths": 0,
        "unsupported_motion_paths": 0,
        "skipped_redundant_events": 0,
        "unsupported_text_by_word": 0,
        "unsupported_text_by_letter": 0,
        "unsupported_text_by_character": 0,
        "unsupported_scale_animations": 0,
        "neutral_scale_animations": 0,
        "unsupported_rotation_animations": 0,
        "neutral_rotation_animations": 0,
        "unsupported_opacity_animations": 0,
        "unsupported_color_animations": 0,
        "unsupported_text_style_animations": 0,
    }

def append_color_transform(
    shape_transforms,
    sid,
    color_kind,
    color_el,
    paragraph_index=None,
):
    append_shape_transform(
        shape_transforms,
        sid,
        {
            "action": "color",
            "color_kind": color_kind,
            "color": color_el,
            "paragraph_index": paragraph_index,
        },
    )


def apply_timeline_event(
    event,
    shape_visible,
    paragraph_visible,
    shape_transforms,
    shape_style_state,
    post_step_events=None,
    color_post_mode=False,
):
    """
    Applique un événement à l'état courant de la slide.

    post_step_events est utilisé pour les effets de couleur from -> to :
    l'état from est appliqué dans l'étape courante si nécessaire, puis
    l'état to est appliqué dans une étape post-animation commune.
    """
    action = event["action"]
    target = event["target"]

    if action in {"show", "hide"}:
        value = action == "show"

        if target[0] == "shape":
            _, sid = target
            shape_visible[sid] = value

        elif target[0] == "paragraph":
            _, sid, pidx = target
            paragraph_visible[(sid, pidx)] = value

            if value:
                shape_visible[sid] = True

        return
        
    simple_transform_fields = {
        "move": [
            ("dx", "dx", None),
            ("dy", "dy", None),
        ],
        "scale": [
            ("sx", "sx", None),
            ("sy", "sy", None),
        ],
        "rotate": [
            ("rotation_mode", "rotation_mode", "delta"),
            ("rot", "rot", None),
        ],
        "opacity": [
            ("alpha", "alpha", None),
        ],
    }

    if action in simple_transform_fields:
        _, sid = target

        transform = {"action": action}

        for dst_key, src_key, default in simple_transform_fields[action]:
            if default is None:
                transform[dst_key] = event[src_key]
            else:
                transform[dst_key] = event.get(src_key, default)

        append_shape_transform(shape_transforms, sid, transform)
        return

    if action == "color":
        color_kind = event["color_kind"]
        color_el = event["color"]

        if target[0] == "shape":
            _, sid = target
            paragraph_index = None

        elif target[0] == "paragraph":
            _, sid, paragraph_index = target

            if color_kind != "text":
                paragraph_index = None

        else:
            return

        append_color_transform(
            shape_transforms,
            sid,
            color_kind,
            color_el,
            paragraph_index,
        )

        set_color_state_key_for_target(
            shape_style_state,
            target,
            color_kind,
            color_key(color_el),
        )

        return

    if action == "color_transition":
        _, sid = target
        color_kind = event["color_kind"]
        from_color = event.get("from_color")
        to_color = event["to_color"]

        if color_post_mode:
            if from_color is not None and color_transition_needs_from_slide(
                event,
                shape_style_state,
            ):
                append_color_transform(
                    shape_transforms,
                    sid,
                    color_kind,
                    from_color,
                )

                set_style_state_key(
                    shape_style_state,
                    sid,
                    color_kind,
                    color_key(from_color),
                )

            if post_step_events is not None:
                post_step_events.append(
                    {
                        "action": "color",
                        "target": target,
                        "color_kind": color_kind,
                        "color": to_color,
                    }
                )

            return

        append_color_transform(
            shape_transforms,
            sid,
            color_kind,
            to_color,
        )

        set_style_state_key(
            shape_style_state,
            sid,
            color_kind,
            color_key(to_color),
        )

    if action == "text_style":
        text_style = event["text_style"]
        value = event["value"]

        if target[0] == "shape":
            _, sid = target
            paragraph_index = None

        elif target[0] == "paragraph":
            _, sid, paragraph_index = target

        else:
            return

        append_shape_transform(
            shape_transforms,
            sid,
            {
                "action": "text_style",
                "text_style": text_style,
                "value": value,
                "paragraph_index": paragraph_index,
            },
        )

        set_text_style_state_value(
            shape_style_state,
            target,
            text_style,
            value,
        )

        return

def step_requires_color_post_state(step_events, shape_style_state):
    """
    Une étape a besoin d'un état post-animation si au moins une transition
    de couleur possède un from différent de l'état courant.
    """
    for event in step_events:
        if event["action"] != "color_transition":
            continue

        if color_transition_needs_from_slide(event, shape_style_state):
            return True

    return False


def apply_timeline_events(
    step_events,
    shape_visible,
    paragraph_visible,
    shape_transforms,
    shape_style_state,
    force_emit=False,
):
    """
    Applique les événements d'une étape dans leur ordre.

    Retourne :
      - emit_slide : faut-il émettre une slide après application ?
      - post_step_events : événements à appliquer après l'émission courante
    """
    emit_slide = force_emit
    post_step_events = []

    color_post_mode = step_requires_color_post_state(
        step_events,
        shape_style_state,
    )

    for event in step_events:
        if event_changes_visible_state(
            event,
            shape_visible,
            paragraph_visible,
            shape_style_state,
            color_post_mode,
        ):
            emit_slide = True

        apply_timeline_event(
            event,
            shape_visible,
            paragraph_visible,
            shape_transforms,
            shape_style_state,
            post_step_events,
            color_post_mode,
        )

    return emit_slide, post_step_events

def write_static_slide(
    slide_root,
    original_rels_data,
    new_entries,
    ct_root,
    pres_rels_root,
    sld_id_lst,
    used_rids,
    new_slide_index,
    new_slide_id,
    shape_visible,
    paragraph_visible,
    paragraph_animated_shapes,
    shape_transforms,
):
    """
    Écrit une copie statique de la slide courante dans le package PPTX.
    """
    new_slide_root = copy.deepcopy(slide_root)

    apply_visibility_to_slide(
        new_slide_root,
        shape_visible,
        paragraph_visible,
        paragraph_animated_shapes,
    )

    apply_transforms_to_slide(new_slide_root, shape_transforms)

    remove_timing_and_transition(new_slide_root)

    new_slide_path = f"ppt/slides/slide{new_slide_index}.xml"
    new_rels_path = f"ppt/slides/_rels/slide{new_slide_index}.xml.rels"

    new_entries[new_slide_path] = xml_bytes(new_slide_root)
    new_entries[new_rels_path] = clean_slide_rels(original_rels_data)

    override = etree.Element(qn(CT, "Override"))
    override.set("PartName", f"/{new_slide_path}")
    override.set("ContentType", SLIDE_CONTENT_TYPE)
    ct_root.append(override)

    rid = next_free_rid(used_rids)

    rel = etree.Element(qn(REL, "Relationship"))
    rel.set("Id", rid)
    rel.set("Type", SLIDE_REL_TYPE)
    rel.set("Target", f"slides/slide{new_slide_index}.xml")
    pres_rels_root.append(rel)

    sld_id = etree.Element(qn(P, "sldId"))
    sld_id.set("id", str(new_slide_id))
    sld_id.set(qn(R, "id"), rid)
    sld_id_lst.append(sld_id)

    return new_slide_index + 1, new_slide_id + 1

def split_pptx_static(input_path, output_path):
    input_path = Path(input_path)
    output_path = Path(output_path)

    with zipfile.ZipFile(input_path, "r") as zin:
        entries = {
            info.filename: zin.read(info.filename)
            for info in zin.infolist()
            if not info.is_dir()
        }

    required = [
        "[Content_Types].xml",
        "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels",
    ]

    for name in required:
        if name not in entries:
            raise ValueError(f"Fichier PPTX incomplet : {name} absent")

    pres_root = parse_xml(entries["ppt/presentation.xml"])
    pres_rels_root = parse_xml(entries["ppt/_rels/presentation.xml.rels"])
    ct_root = parse_xml(entries["[Content_Types].xml"])
    slide_width, slide_height = presentation_slide_size(pres_root)

    rel_by_id = {
        rel.get("Id"): rel
        for rel in pres_rels_root.findall(qn(REL, "Relationship"))
    }

    sld_id_lst = pres_root.find(".//p:sldIdLst", namespaces=NS)
    if sld_id_lst is None:
        raise ValueError("Aucune liste de slides p:sldIdLst trouvée")

    original_slide_paths = []

    for sld_id in sld_id_lst.findall("./p:sldId", namespaces=NS):
        rid = sld_id.get(qn(R, "id"))
        rel = rel_by_id.get(rid)

        if rel is None or rel.get("Type") != SLIDE_REL_TYPE:
            continue

        target = rel.get("Target")
        slide_path = rel_target_to_part("ppt/presentation.xml", target)

        if slide_path in entries:
            original_slide_paths.append(slide_path)

    if not original_slide_paths:
        raise ValueError("Aucune slide trouvée dans la présentation")

    new_entries = {
        name: data
        for name, data in entries.items()
        if not SLIDE_RE.match(name) and not SLIDE_RELS_RE.match(name)
    }

    for rel in list(pres_rels_root):
        if rel.get("Type") == SLIDE_REL_TYPE:
            pres_rels_root.remove(rel)

    used_rids = {
        rel.get("Id")
        for rel in pres_rels_root.findall(qn(REL, "Relationship"))
        if rel.get("Id")
    }

    for child in list(sld_id_lst):
        sld_id_lst.remove(child)

    for child in list(ct_root):
        if child.tag != qn(CT, "Override"):
            continue

        part_name = child.get("PartName", "")

        if re.match(r"^/ppt/slides/slide\d+\.xml$", part_name):
            ct_root.remove(child)

    new_slide_index = 1
    new_slide_id = 256

    total_original = 0
    total_generated = 0
    total_events = 0
    total_ignored = 0
    
    report = new_conversion_report()

    for original_slide_path in original_slide_paths:
        total_original += 1

        slide_root = parse_xml(entries[original_slide_path])
        original_rels_path = f"ppt/slides/_rels/{posixpath.basename(original_slide_path)}.rels"
        original_rels_data = entries.get(original_rels_path)

        events, ignored = extract_timeline_events(
            slide_root,
            slide_width,
            slide_height,
            report,
        )

        total_events += len(events)
        total_ignored += ignored

        (
            shape_visible,
            paragraph_visible,
            paragraph_animated_shapes,
        ) = build_initial_visibility(slide_root, events)
        shape_style_state = build_initial_style_state(slide_root)

        events_by_step = {}

        for event in events:
            events_by_step.setdefault(event["step"], []).append(event)

        max_step = max(events_by_step.keys(), default=0)

        # Transformations cumulées pour la slide originale courante.
        # On les réinitialise à chaque nouvelle slide originale,
        # mais on les conserve entre les étapes générées de cette slide.
        shape_transforms = {}

        for step in range(0, max_step + 1):
            step_events = events_by_step.get(step, [])

            emit_slide, post_step_events = apply_timeline_events(
                step_events,
                shape_visible,
                paragraph_visible,
                shape_transforms,
                shape_style_state,
                force_emit=(step == 0),
            )

            if not emit_slide:
                if step_events:
                    report["skipped_redundant_events"] += len(step_events)
                continue

            new_slide_index, new_slide_id = write_static_slide(
                slide_root,
                original_rels_data,
                new_entries,
                ct_root,
                pres_rels_root,
                sld_id_lst,
                used_rids,
                new_slide_index,
                new_slide_id,
                shape_visible,
                paragraph_visible,
                paragraph_animated_shapes,
                shape_transforms,
            )

            total_generated += 1

            while post_step_events:
                emit_post_slide, next_post_step_events = apply_timeline_events(
                    post_step_events,
                    shape_visible,
                    paragraph_visible,
                    shape_transforms,
                    shape_style_state,
                    force_emit=False,
                )

                if emit_post_slide:
                    new_slide_index, new_slide_id = write_static_slide(
                        slide_root,
                        original_rels_data,
                        new_entries,
                        ct_root,
                        pres_rels_root,
                        sld_id_lst,
                        used_rids,
                        new_slide_index,
                        new_slide_id,
                        shape_visible,
                        paragraph_visible,
                        paragraph_animated_shapes,
                        shape_transforms,
                    )

                    total_generated += 1

                elif post_step_events:
                    report["skipped_redundant_events"] += len(post_step_events)

                post_step_events = next_post_step_events

    new_entries["ppt/presentation.xml"] = xml_bytes(pres_root)
    new_entries["ppt/_rels/presentation.xml.rels"] = xml_bytes(pres_rels_root)
    new_entries["[Content_Types].xml"] = xml_bytes(ct_root)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in new_entries.items():
            zout.writestr(name, data)

    print("Terminé.")
    print(f"Slides originales       : {total_original}")
    print(f"Slides générées         : {total_generated}")
    print(f"Événements supportés    : {total_events}")
    print(f"Événements ignorés      : {total_ignored}")
    if report["closed_or_zero_motion_paths"]:
        print(
            "Motion paths fermés/nuls : "
            f"{report['closed_or_zero_motion_paths']} "
            "(animation dynamique sans état final statique distinct)"
        )

    if report["unsupported_motion_paths"]:
        print(
            "Motion paths non supportés : "
            f"{report['unsupported_motion_paths']} "
            "(chemin non interprétable par le convertisseur)"
        )

    if report["skipped_redundant_events"]:
        print(
            "Étapes sans changement visible : "
            f"{report['skipped_redundant_events']} "
            "(non exportées comme slides séparées)"
        )

    unsupported_text_total = (
        report["unsupported_text_by_word"]
        + report["unsupported_text_by_letter"]
        + report["unsupported_text_by_character"]
    )

    if unsupported_text_total:
        print(
            "Animations texte non supportées : "
            f"{unsupported_text_total} "
            "(mot par mot / lettre par lettre / caractère par caractère ; "
            "non converties en V1)"
        )

        if report["unsupported_text_by_word"]:
            print(f"  - mot par mot       : {report['unsupported_text_by_word']}")

        if report["unsupported_text_by_letter"]:
            print(f"  - lettre par lettre : {report['unsupported_text_by_letter']}")

        if report["unsupported_text_by_character"]:
            print(f"  - plage caractères  : {report['unsupported_text_by_character']}")
    
    if report["unsupported_scale_animations"]:
        print(
            "Animations Grow/Shrink non supportées : "
            f"{report['unsupported_scale_animations']} "
            "(facteur de taille non interprétable)"
        )

    if report["neutral_scale_animations"]:
        print(
            "Animations Grow/Shrink neutres : "
            f"{report['neutral_scale_animations']} "
            "(aucun changement de taille final)"
        )
    
    if report["unsupported_rotation_animations"]:
        print(
            "Animations de rotation non supportées : "
            f"{report['unsupported_rotation_animations']} "
            "(angle de rotation non interprétable)"
        )

    if report["neutral_rotation_animations"]:
        print(
            "Animations de rotation neutres : "
            f"{report['neutral_rotation_animations']} "
            "(aucun changement d'angle final)"
        )

    if report["unsupported_opacity_animations"]:
        print(
            "Animations de transparence non supportées : "
            f"{report['unsupported_opacity_animations']} "
            "(opacité finale non interprétable)"
        )

    if report["unsupported_color_animations"]:
        print(
            "Animations de couleur non supportées : "
            f"{report['unsupported_color_animations']} "
            "(cible ou couleur finale non interprétable)"
        )
    
    if report["unsupported_text_style_animations"]:
        print(
            "Animations de style texte non supportées : "
            f"{report['unsupported_text_style_animations']} "
            "(gras / italique / souligné / barré non interprétable)"
        )

    print(f"Sortie                  : {output_path}")


def default_output_path(input_path):
    input_path = Path(input_path)

    if input_path.suffix:
        return input_path.with_name(f"{input_path.stem}_split{input_path.suffix}")

    return input_path.with_name(f"{input_path.name}_split")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convertit un PPTX animé en PPTX statique par états successifs."
    )

    parser.add_argument(
        "input",
        help="PPTX source",
    )

    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="PPTX de sortie. Si omis, ajoute _split au nom du fichier d'entrée.",
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else default_output_path(input_path)

    if input_path.resolve() == output_path.resolve():
        raise ValueError("Le fichier de sortie ne doit pas écraser le fichier d'entrée.")

    split_pptx_static(input_path, output_path)