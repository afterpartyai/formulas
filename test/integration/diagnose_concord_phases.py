"""H1 diagnostic — Phase A cost in isolation.

Per FORMULAS_ITERATIVE_CIRCULAR_SOLVER.md §3.5c H1: is the wall-clock cost in
Phase A graph-rewriting (cycle enumeration + IF-rewrite) or in Phase B iterative
solving?

Times three configurations on a workbook (default Concord Remington 1052):
  A. finish(circular=False)                    — load + compile, NO cycle work
  B. finish(circular=True, iterate=False)      — + Phase A only
  C. finish(circular=True, iterate=True)       — + Phase B iterative solve

Hard timeout per phase: 900s (15 min). Earlier termination is itself a finding.

Run:
    /var/www/acquios_ai_v1/excel_service/ve_excel_service_p311/bin/python3 \\
        test/integration/diagnose_concord_phases.py [path]
"""
import os.path as osp
import signal
import sys
import time

DEFAULT_PATH = '/tmp/recalc_sanitized_1052.xlsx'
PHASE_TIMEOUT_SEC = 900


class PhaseTimeout(Exception):
    pass


def _alarm_handler(signum, frame):
    raise PhaseTimeout()


def run_phase(label, fn):
    """Execute fn() with a hard timeout. Return (elapsed, result_or_None, status)."""
    signal.signal(signal.SIGALRM, _alarm_handler)
    signal.alarm(PHASE_TIMEOUT_SEC)
    t0 = time.perf_counter()
    try:
        result = fn()
        elapsed = time.perf_counter() - t0
        signal.alarm(0)
        return elapsed, result, 'OK'
    except PhaseTimeout:
        elapsed = time.perf_counter() - t0
        return elapsed, None, f'TIMEOUT_{PHASE_TIMEOUT_SEC}s'
    except Exception as e:
        elapsed = time.perf_counter() - t0
        signal.alarm(0)
        return elapsed, None, f'EXCEPTION: {type(e).__name__}: {e}'


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PATH
    if not osp.exists(path):
        print(f"ERROR: workbook not found: {path}")
        return 2

    import formulas
    print(f"formulas: {osp.realpath(formulas.__file__)}")
    print(f"workbook: {path}")
    print(f"per-phase timeout: {PHASE_TIMEOUT_SEC}s")
    print()

    # --- Phase A baseline: load + compile only, no circular handling.
    print("=" * 70)
    print("H1.A — finish(circular=False)   [no cycle work]")
    print("=" * 70)
    def _A():
        m = formulas.ExcelModel().loads(path).finish(circular=False)
        return m
    t_A, m_A, st_A = run_phase('A', _A)
    print(f"  status: {st_A}")
    print(f"  elapsed: {t_A:.1f}s")
    if st_A != 'OK':
        print("  --> baseline did not complete; cannot proceed to H1.B/C.")
        return 1
    print()

    # --- + Phase A graph-rewriting only.
    print("=" * 70)
    print("H1.B — finish(circular=True, iterate=False)   [+ Phase A only]")
    print("=" * 70)
    def _B():
        m = formulas.ExcelModel().loads(path).finish(circular=True, iterate=False)
        return m
    t_B, m_B, st_B = run_phase('B', _B)
    print(f"  status: {st_B}")
    print(f"  elapsed: {t_B:.1f}s")
    print(f"  Phase A cost (B - A): {(t_B - t_A) if st_B == 'OK' else 'N/A (B did not complete)'}s")
    print()

    # --- + Phase B iterative.
    print("=" * 70)
    print("H1.C — finish(circular=True, iterate=True)    [+ Phase B iterative]")
    print("=" * 70)
    def _C():
        m = formulas.ExcelModel().loads(path).finish(circular=True, iterate=True)
        return m
    t_C, m_C, st_C = run_phase('C', _C)
    print(f"  status: {st_C}")
    print(f"  elapsed: {t_C:.1f}s")
    if st_C == 'OK' and st_B == 'OK':
        print(f"  Phase B cost (C - B): {(t_C - t_B):.1f}s")
    elif st_C != 'OK' and st_B == 'OK':
        print(f"  Phase B exceeded timeout BEYOND completed Phase A.")
        print(f"  Conclusion: Phase B is at least {(PHASE_TIMEOUT_SEC):.0f}s on Concord.")
    print()

    # --- Summary
    print("=" * 70)
    print("H1 SUMMARY")
    print("=" * 70)
    print(f"  load+compile only           A: {t_A:>7.1f}s   {st_A}")
    print(f"  + Phase A graph-rewriting   B: {t_B:>7.1f}s   {st_B}")
    print(f"  + Phase B iterative solve   C: {t_C:>7.1f}s   {st_C}")
    if st_A == 'OK' and st_B == 'OK':
        print(f"\n  Phase A isolated cost:  {t_B - t_A:.1f}s")
    if st_A == 'OK' and st_B == 'OK' and st_C == 'OK':
        print(f"  Phase B isolated cost:  {t_C - t_B:.1f}s")
    elif st_B == 'OK' and st_C != 'OK':
        print(f"  Phase B isolated cost:  >= {PHASE_TIMEOUT_SEC}s (timed out)")

    print()
    print("Hypothesis verdict:")
    if st_A == 'OK' and st_B == 'OK':
        phase_a_cost = t_B - t_A
        if phase_a_cost > 300:
            print(f"  H1 SUPPORTED: Phase A alone is {phase_a_cost:.0f}s (> 5 min).")
            print(f"  Conclusion: cycle enumeration / IF-rewrite is on the critical path.")
        else:
            print(f"  H1 REJECTED: Phase A is fast ({phase_a_cost:.0f}s).")
            if st_C != 'OK':
                print(f"  --> blowup is in Phase B iteration; proceed to H2/H3/H4.")
            else:
                print(f"  --> whole pipeline completed; recheck what 41-min kill was diagnosing.")

    return 0


if __name__ == '__main__':
    sys.exit(main())
