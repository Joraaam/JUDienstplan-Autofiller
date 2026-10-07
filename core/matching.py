"""Zuordnung Excel-Datei <-> Zeile im Dienstplan (HC: 'Nachname V.', FACL/LUG: ausgeschriebene Namen)."""
import re
import unicodedata


def norm(s):
    s = s.lower().replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    s = unicodedata.normalize("NFKD", s)
    return re.sub(r"[^a-z]", "", s)


def tokens(s):
    return [norm(t) for t in re.split(r"[\s._\-]+", s) if norm(t)]


def tok_match(plan_tok, file_tok):
    if plan_tok == file_tok:
        return 2
    if len(plan_tok) >= 2 and (file_tok.startswith(plan_tok) or plan_tok.startswith(file_tok) and len(file_tok) >= 2):
        return 1
    return 0


def match_person(nachname, vorname, people):
    """Liefert (Index oder None, Kandidatenliste). HC: 'Nachname V.'; FACL: ausgeschriebener Name."""
    ftoks = tokens(nachname) + tokens(vorname)
    if not ftoks:
        return None, []
    scored = []
    for i, p in enumerate(people):
        ptoks = tokens(p["name"])
        if not ptoks:
            continue
        initial_fmt = len(ptoks) >= 2 and len(ptoks[-1]) == 1
        surname = ptoks[:-1] if initial_fmt else ptoks
        used, score, ok = set(), 0, True
        for st in surname:
            best = max(((tok_match(st, ft), j) for j, ft in enumerate(ftoks) if j not in used), default=(0, -1))
            if best[0] == 0:
                ok = False
                break
            used.add(best[1])
            score += best[0]
        if not ok:
            continue
        rest = [ft for j, ft in enumerate(ftoks) if j not in used]
        if initial_fmt:
            if not any(ft.startswith(ptoks[-1]) for ft in rest):
                continue
            if ftoks[0] == surname[0]:
                score += 1
        else:
            score += 3 * len(used) / len(ftoks) + (0.5 if len(used) == len(ftoks) else 0)
            if len(surname) < 2 and len(ftoks) > 2:
                continue  # ein einzelnes Namensteil reicht bei langen Namen nicht
        scored.append((score, i))
    if not scored:
        return None, []
    scored.sort(reverse=True)
    top = [i for s_, i in scored if abs(s_ - scored[0][0]) < 1e-9]
    return (top[0] if len(top) == 1 else None), [i for _, i in scored]
