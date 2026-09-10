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
import sys, zipfile, re


def parts(p):
    with zipfile.ZipFile(p) as z:
        return {n: z.read(n) for n in z.namelist()}


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


def sqrefs(x):
    s = set()
    for m in re.findall(r'<conditionalFormatting[^>]*sqref="([^"]+)"', x):
        s.update(m.split())
    return s


def formulas(P):
    """address -> formula text, for every sheet part."""
    out = {}
    for k, v in P.items():
        if not (k.startswith('xl/worksheets/') and k.endswith('.xml')):
            continue
        text = v.decode('utf-8', 'replace')
        # Drop self-closing cell tags first. Otherwise `<c r="BS3" s="1"/>` matches
        # `<c r="..."[^>]*>(.*?)</c>` by swallowing the NEXT cell's body, and every
        # formula after an empty cell gets reported under the wrong address.
        text = re.sub(r'<c\b[^>]*/>', '', text)
        for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*>(.*?)</c>', text, re.S):
            addr, body = m.group(1), m.group(2)
            fm = re.search(r'<f([^>]*)>(.*?)</f>', body, re.S)
            if fm:
                out[k + '!' + addr] = (fm.group(1).strip(), fm.group(2))
    return out


def main():
    tpl, out = parts(sys.argv[1]), parts(sys.argv[2])

    print("== part counts ==")
    print("  template: %d   output: %d" % (len(tpl), len(out)))

    missing = sorted(set(tpl) - set(out))
    added = sorted(set(out) - set(tpl))
    print("\n== parts MISSING from output ==")
    print("  (none)" if not missing else "\n".join("  - " + m for m in missing))
    print("\n== parts ADDED in output ==")
    print("  (none)" if not added else "\n".join("  + " + a for a in added))

    tm, om = sheet_map(tpl), sheet_map(out)

    print("\n== sheet -> part mapping (aligned BY NAME) ==")
    for n in tm:
        print("  %-36s tpl=%-24s out=%s" % (n[:36], tm[n], om.get(n, '** MISSING **')))

    print("\n== dataValidation counts ==")
    for n, tp in tm.items():
        op = om.get(n)
        tc = tpl[tp].decode('utf-8', 'replace').count('<dataValidation ')
        oc = out[op].decode('utf-8', 'replace').count('<dataValidation ') if op else -1
        print("  %-36s tpl=%3d out=%3d  %s" % (n[:36], tc, oc, "OK" if tc == oc else "** DIFF **"))

    print("\n== conditionalFormatting sqref coverage (coverage, NOT rule count) ==")
    for n, tp in tm.items():
        op = om.get(n)
        a = sqrefs(tpl[tp].decode('utf-8', 'replace'))
        b = sqrefs(out[op].decode('utf-8', 'replace')) if op else set()
        flag = "OK" if a == b else "** DIFF **"
        print("  %-36s tpl=%3d out=%3d  lost=%s gained=%s  %s"
              % (n[:36], len(a), len(b), sorted(a - b), sorted(b - a), flag))

    ft, fo = formulas(tpl), formulas(out)
    print("\n== FORMULAS ==")
    print("  template: %d   output: %d" % (len(ft), len(fo)))

    lost = sorted(set(ft) - set(fo))
    new = sorted(set(fo) - set(ft))
    changed = sorted(k for k in set(ft) & set(fo) if ft[k] != fo[k])

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


main()
