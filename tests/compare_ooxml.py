"""Compare two .xlsx packages part-by-part, for requirement 15 (template formats
must survive the mapping run).

Usage:
    py tests/compare_ooxml.py <template.xlsx> <output.xlsx>

Two traps this script exists to avoid, both hit during earlier verification:

1. Sheet file numbering SHIFTS between the two packages (sheet2.xml.rels becomes
   sheet3.xml.rels). Comparing xl/worksheets/sheet3.xml across packages compares
   two different worksheets. Sheets are resolved by NAME here.

2. Conditional-format rule COUNT is not the criterion. Excel merges adjacent
   rules with identical criteria (observed 48 -> 43), losing nothing. Judge by
   sqref COVERAGE, which is what this script reports.
"""
import re
import sys
import zipfile
from collections import defaultdict


def parts(p):
    with zipfile.ZipFile(p) as z:
        return {n: z.read(n) for n in z.namelist() if not n.endswith('/')}


def sheet_map(P):
    """sheet name -> part path, via workbook.xml + its rels."""
    wb = P['xl/workbook.xml'].decode('utf-8', 'replace')
    rels = P['xl/_rels/workbook.xml.rels'].decode('utf-8', 'replace')
    rid2t = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    m = {}
    for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        t = rid2t.get(rid, '').lstrip('/')
        if t and not t.startswith('xl/'):
            t = 'xl/' + t
        m[name] = t
    return m


def column_number(letters):
    value = 0
    for char in letters.upper():
        value = value * 26 + ord(char) - 64
    return value


def sqref_coverage(refs):
    """Canonical row intervals, so A1:A3 B1:B3 equals A1:B3."""
    rows = defaultdict(list)
    for token in refs:
        ends = token.replace('$', '').split(':')
        if len(ends) == 1:
            ends.append(ends[0])
        parsed = []
        for end in ends:
            match = re.fullmatch(r'([A-Z]+)(\d+)', end, re.I)
            if not match:
                raise ValueError('Unsupported sqref: ' + token)
            parsed.append((column_number(match.group(1)), int(match.group(2))))
        (c1, r1), (c2, r2) = parsed
        c1, c2 = sorted((c1, c2))
        r1, r2 = sorted((r1, r2))
        for row in range(r1, r2 + 1):
            rows[row].append((c1, c2))
    canonical = []
    for row, intervals in sorted(rows.items()):
        merged = []
        for start, end in sorted(intervals):
            if merged and start <= merged[-1][1] + 1:
                merged[-1] = (merged[-1][0], max(merged[-1][1], end))
            else:
                merged.append((start, end))
        canonical.append((row, tuple(merged)))
    return tuple(canonical)


def element_sqrefs(x, element):
    refs = []
    for match in re.findall(r'<%s\b[^>]*\bsqref="([^"]+)"' % element, x):
        refs.extend(match.split())
    return refs


def formulas(P, sheets):
    """address -> formula text, for every sheet part."""
    out = {}
    for sheet_name, part_path in sheets.items():
        if part_path not in P:
            continue
        text = P[part_path].decode('utf-8', 'replace')
        # Drop self-closing cell tags first. Otherwise `<c r="BS3" s="1"/>` matches
        # `<c r="..."[^>]*>(.*?)</c>` by swallowing the NEXT cell's body, and every
        # formula after an empty cell gets reported under the wrong address.
        text = re.sub(r'<c\b[^>]*/>', '', text)
        for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*>(.*?)</c>', text, re.S):
            addr, body = m.group(1), m.group(2)
            fm = re.search(r'<f([^>]*)>(.*?)</f>', body, re.S)
            if fm:
                out[sheet_name + '!' + addr] = (fm.group(1).strip(), fm.group(2))
    return out


def main():
    tpl, out = parts(sys.argv[1]), parts(sys.argv[2])
    failed = False

    print("== part counts ==")
    print("  template: %d   output: %d" % (len(tpl), len(out)))

    missing = sorted(set(tpl) - set(out))
    added = sorted(set(out) - set(tpl))
    failed = failed or bool(missing)
    print("\n== parts MISSING from output ==")
    print("  (none)" if not missing else "\n".join("  - " + m for m in missing))
    print("\n== parts ADDED in output ==")
    print("  (none)" if not added else "\n".join("  + " + a for a in added))

    tm, om = sheet_map(tpl), sheet_map(out)
    if set(tm) != set(om):
        failed = True

    print("\n== sheet -> part mapping (aligned BY NAME) ==")
    for n in tm:
        print("  %-36s tpl=%-24s out=%s" % (n[:36], tm[n], om.get(n, '** MISSING **')))

    print("\n== dataValidation counts ==")
    for n, tp in tm.items():
        op = om.get(n)
        template_xml = tpl[tp].decode('utf-8', 'replace')
        output_xml = out[op].decode('utf-8', 'replace') if op else ''
        template_refs = element_sqrefs(template_xml, 'dataValidation')
        output_refs = element_sqrefs(output_xml, 'dataValidation')
        ok = (template_xml.count('<dataValidation ') == output_xml.count('<dataValidation ') and
              sqref_coverage(template_refs) == sqref_coverage(output_refs))
        failed = failed or not ok
        print("  %-36s tpl=%3d out=%3d  %s" %
              (n[:36], len(template_refs), len(output_refs), "OK" if ok else "** DIFF **"))

    print("\n== conditionalFormatting sqref coverage (coverage, NOT rule count) ==")
    for n, tp in tm.items():
        op = om.get(n)
        a_refs = element_sqrefs(tpl[tp].decode('utf-8', 'replace'), 'conditionalFormatting')
        b_refs = element_sqrefs(out[op].decode('utf-8', 'replace'), 'conditionalFormatting') if op else []
        ok = sqref_coverage(a_refs) == sqref_coverage(b_refs)
        failed = failed or not ok
        print("  %-36s tpl=%3d out=%3d  %s"
              % (n[:36], len(a_refs), len(b_refs), "OK" if ok else "** DIFF **"))

    ft, fo = formulas(tpl, tm), formulas(out, om)
    print("\n== FORMULAS ==")
    print("  template: %d   output: %d" % (len(ft), len(fo)))

    lost = sorted(set(ft) - set(fo))
    new = sorted(set(fo) - set(ft))
    changed = sorted(k for k in set(ft) & set(fo) if ft[k] != fo[k])
    failed = failed or bool(lost or new or changed)

    print("\n  -- SURVIVED (%d) --" % len(set(ft) & set(fo)))
    for k in sorted(set(ft) & set(fo)):
        arr = ' ARRAY' if 'array' in ft[k][0] else ''
        print("     %-46s%s  =%s" % (k, arr, ft[k][1]))

    print("\n  -- LOST (%d) --" % len(lost))
    print("     (none)" if not lost else "\n".join("     %s  =%s" % (k, ft[k][1]) for k in lost))

    if new:
        print("\n  -- NEW in output (%d) --" % len(new))
        for k in new:
            print("     %s  =%s" % (k, fo[k][1]))

    if changed:
        print("\n  -- CHANGED (%d) ** investigate ** --" % len(changed))
        for k in changed:
            print("     %s\n        tpl=%s\n        out=%s" % (k, ft[k], fo[k]))

    return 1 if failed else 0


raise SystemExit(main())
