"""H2 — Iterative solver scaling sweep.

Per FORMULAS_ITERATIVE_CIRCULAR_SOLVER.md §3.5c H2: does our solver scale
pathologically with SCC size? Builds synthetic linear-convergent SCCs at
sizes [10, 50, 100, 200, 400, 800] using build_synthetic_sccs.build() and
times finish(circular=True, iterate=True) on each.

Output: a markdown table of size → Phase A time, Phase B time, total, convergence
status. Fits the per-cell timing into _last_solve_info.

Per-size hard timeout 300s — anything bigger means the scaling itself is the
finding.

Run:
    /var/www/acquios_ai_v1/excel_service/ve_excel_service_p311/bin/python3 \\
        test/integration/h2_scc_scaling_sweep.py [--pattern ring]

Reusable: imports build() from build_synthetic_sccs; can be invoked with
--sizes / --pattern args to extend the sweep without script edits.
"""
import argparse
import math
import os.path as osp
import signal
import sys
import tempfile
import time

sys.path.insert(0, osp.dirname(osp.abspath(__file__)))
from build_synthetic_sccs import build as build_scc

DEFAULT_SIZES = [10, 50, 100, 200, 400, 800]
DEFAULT_PATTERN = 'ring'
PER_SIZE_TIMEOUT_SEC = 300


class SweepTimeout(Exception):
    pass


def _alarm(signum, frame):
    raise SweepTimeout()


def _time_solve(path):
    """Return (t_phase_a, t_phase_b_or_total, status, solve_info)."""
    import formulas
    signal.signal(signal.SIGALRM, _alarm)
    signal.alarm(PER_SIZE_TIMEOUT_SEC)
    try:
        t0 = time.perf_counter()
        m = formulas.ExcelModel().loads(path).finish(circular=True, iterate=False)
        t_a = time.perf_counter() - t0

        t0 = time.perf_counter()
        m = formulas.ExcelModel().loads(path).finish(circular=True, iterate=True)
        t_ab = time.perf_counter() - t0
        signal.alarm(0)

        info = getattr(m, '_last_solve_info', None)
        return t_a, t_ab - t_a, 'OK', info
    except SweepTimeout:
        return None, None, f'TIMEOUT_{PER_SIZE_TIMEOUT_SEC}s', None
    except Exception as e:
        signal.alarm(0)
        return None, None, f'EXCEPTION: {type(e).__name__}: {e}', None


def _converged_count(info):
    if not info:
        return '-'
    sccs = info.get('sccs', [])
    if not sccs:
        return '0/0'
    n = sum(1 for s in sccs if s.get('converged'))
    return f"{n}/{len(sccs)}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pattern', default=DEFAULT_PATTERN)
    ap.add_argument('--sizes', default=','.join(str(s) for s in DEFAULT_SIZES))
    args = ap.parse_args()

    sizes = [int(s) for s in args.sizes.split(',')]
    pattern = args.pattern

    import formulas
    print(f"formulas: {osp.realpath(formulas.__file__)}")
    print(f"pattern:  {pattern}")
    print(f"sizes:    {sizes}")
    print(f"per-size timeout: {PER_SIZE_TIMEOUT_SEC}s")
    print()

    print(f"{'size':>6} {'phase A (s)':>12} {'phase B (s)':>12} {'total (s)':>10} "
          f"{'B/A ratio':>10} {'converged':>10}  status")
    print('-' * 90)

    rows = []
    for n in sizes:
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tf:
            scc_path = tf.name
        build_scc(pattern, n, scc_path)
        t_a, t_b, status, info = _time_solve(scc_path)
        if t_a is None:
            print(f"{n:>6} {'-':>12} {'-':>12} {'-':>10} {'-':>10} {'-':>10}  {status}")
            rows.append((n, None, None, status))
        else:
            ratio = (t_b / t_a) if t_a > 0 else float('inf')
            print(f"{n:>6} {t_a:>12.3f} {t_b:>12.3f} {t_a + t_b:>10.3f} "
                  f"{ratio:>10.2f} {_converged_count(info):>10}  {status}")
            rows.append((n, t_a, t_b, status))

    # Estimate scaling exponent for Phase B from the OK rows.
    ok = [(n, b) for n, _, b, st in rows if st == 'OK' and b and b > 0]
    if len(ok) >= 3:
        # Fit log(B) = a + k * log(N) -- two-point first/last as a cheap estimator.
        n0, b0 = ok[0]
        n1, b1 = ok[-1]
        k = math.log(b1 / b0) / math.log(n1 / n0)
        print()
        print(f"Phase B scaling estimate: B ∝ N^{k:.2f}  (from size {n0} → {n1})")
        if k > 2.5:
            print("  H2 SUPPORTED: super-quadratic scaling. Solver itself is pathological.")
        elif k > 1.5:
            print("  H2 PARTIAL: super-linear but sub-cubic. Algorithm choice may need review.")
        else:
            print("  H2 REJECTED: near-linear scaling. Blowup is not in the solver per se.")

    return 0


if __name__ == '__main__':
    sys.exit(main())
