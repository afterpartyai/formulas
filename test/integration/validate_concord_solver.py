"""Slow integration test: run the AcquiOS fork's iterative circular solver
against Concord's Remington workbook (doc 1052) and verify the headline
output cells match Excel-cached values within tolerance.

Per spec FORMULAS_ITERATIVE_CIRCULAR_SOLVER.md §9 (validation against actual goal):
acceptance is `Annual Proforma!D221` (levered IRR) = 0.21040 ± 0.002, plus the
other cells in §2.4 within 1% relative.

NOT a CI unit test. Compile time ~6-12 min, peak RSS ~6 GB. Run manually:

    cd /var/www/acquios_ai_v1/formulas
    /var/www/acquios_ai_v1/excel_service/ve_excel_service_p311/bin/python3 \\
        test/integration/validate_concord_solver.py [path_to_workbook]

If the workbook arg is omitted, defaults to /tmp/recalc_sanitized_1052.xlsx
(the deactivation+DataTable-neutralized Remington copy used during the
Concord debugging effort).
"""
import os
import os.path as osp
import sys
import time

# Acceptance targets from FORMULAS_ITERATIVE_CIRCULAR_SOLVER.md §2.4.
# (sheet, cell, expected_value, label, tolerance_pct_or_abs)
TARGETS = [
    ('Annual Proforma', 'D221',  0.21039991,    'Levered IRR',         0.002),     # absolute
    ('Annual Proforma', 'G134',  3392305.45,    'Y1 NOI',              0.01),      # 1% relative
    ('Annual Proforma', 'G115',  7039337.08,    'Y1 Gross Income',     0.01),
    ('Annual Proforma', 'G132',  3647031.63,    'Y1 OpEx',             0.01),
    ('Annual Proforma', 'G198',  1283412.08,    'Y1 Net Cash Flow',    0.01),
    ('Annual Proforma', 'D220',  2.4124,        'Equity Multiple',     0.01),
    # These two aren't in the SCC; should solve regardless.
    ('Assumptions',     'E15',   54350000.0,    'Purchase Price',      0.001),
    ('Assumptions',     'D37',   0.0565,        'Exit Cap Rate',       0.001),
]


def find_in_solution(sol, sheet, cell):
    """Locate a value in the solution dict by (sheet, cell)."""
    su, cu = sheet.upper(), cell.upper()
    for k in sol.keys():
        ks = str(k).upper()
        if su in ks and (ks.endswith(f"!{cu}") or ks.endswith(f"!{cu}'")):
            v = sol[k]
            val = v.value if hasattr(v, 'value') else v
            try:
                if hasattr(val, 'ravel'):
                    val = val.ravel()[0] if val.size else val
            except Exception:
                pass
            return val
    return None


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else '/tmp/recalc_sanitized_1052.xlsx'
    if not osp.exists(path):
        print(f"ERROR: workbook not found: {path}")
        return 2

    # Sanity: confirm we're loading the fork, not site-packages
    import formulas
    fork_file = osp.realpath(formulas.__file__)
    expected_fork = '/var/www/acquios_ai_v1/formulas/formulas/__init__.py'
    if not fork_file.startswith(osp.dirname(osp.dirname(osp.realpath(expected_fork)))):
        print(f"WARNING: `formulas` resolves to {fork_file}, not the fork tree.")
        print(f"         expected fork dir: {osp.dirname(osp.dirname(osp.realpath(expected_fork)))}")
        print("         Edits to the fork may not be live. Aborting.")
        return 3
    print(f"formulas: {fork_file}")
    print(f"workbook: {path}")
    print()

    print("Compiling workbook + invoking solver (this is the slow step)...")
    t0 = time.time()
    m = formulas.ExcelModel().loads(path).finish(circular=True, iterate=True)
    t_compile = time.time() - t0
    print(f"  compile+solve: {t_compile:.1f}s")

    t0 = time.time()
    sol = m.calculate()
    t_calc = time.time() - t0
    print(f"  calculate:     {t_calc:.1f}s")

    info = getattr(m, '_last_solve_info', None)
    print()
    print("Solve info (_last_solve_info):")
    if info:
        print(f"  iterate enabled: {info.get('iterate_enabled')}")
        for s in info.get('sccs', []):
            print(f"  SCC size={s['size']}, converged={s['converged']}, "
                  f"method={s['method']}")
    else:
        print("  (no _last_solve_info attribute — solver didn't run)")
    print()

    # Verify each acceptance target.
    print("Acceptance check vs Excel-cached values:")
    print(f"  {'cell':28} {'label':22} {'expected':>14}  {'got':>14}  {'Δ':>10}  status")
    print("  " + "-" * 100)
    n_pass = n_fail = 0
    failures = []
    for sheet, cell, expected, label, tol in TARGETS:
        val = find_in_solution(sol, sheet, cell)
        addr = f"{sheet}!{cell}"
        if not isinstance(val, (int, float)):
            print(f"  {addr:28} {label:22} {expected:>14.4f}  {str(val)[:14]:>14}  "
                  f"{'':>10}  ✗ NOT NUMERIC")
            n_fail += 1
            failures.append((addr, label, expected, val))
            continue
        delta = val - expected
        # Pass if either absolute delta (for tight tolerances like IRR) is OK,
        # or relative delta < tol when expected is non-tiny.
        if abs(delta) < tol or (abs(expected) > 1e-6 and abs(delta) / abs(expected) < tol):
            print(f"  {addr:28} {label:22} {expected:>14.4f}  {val:>14.4f}  "
                  f"{delta:>+10.4f}  ✓")
            n_pass += 1
        else:
            print(f"  {addr:28} {label:22} {expected:>14.4f}  {val:>14.4f}  "
                  f"{delta:>+10.4f}  ✗ FAIL (tol={tol})")
            n_fail += 1
            failures.append((addr, label, expected, val))

    print()
    print(f"Result: {n_pass}/{len(TARGETS)} PASS, {n_fail} FAIL")
    if failures:
        print("\nFailures:")
        for addr, label, exp, got in failures:
            print(f"  {addr} ({label}): expected {exp}, got {got}")
    return 0 if n_fail == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
