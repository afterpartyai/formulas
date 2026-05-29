"""H1.1 — Per-step Phase A timing breakdown on Concord.

Per spec §3.5c next-actions: H1 showed Phase A > 900s on Concord but did not
isolate which step (simple_cycles enumeration vs IF-rewrite pass vs DSP
mutation). This script runs finish(circular=True, iterate=False) with
FORMULAS_TIMING=1 set, so each Phase A sub-step logs its wall clock.

Hard timeout: 1800s. We expect simple_cycles to be the offender (it ran 58 min
on May 28 against this same graph). If simple_cycles never prints its
[FORMULAS_TIMING] line, it didn't return — that IS the diagnosis.

Run:
    /var/www/acquios_ai_v1/excel_service/ve_excel_service_p311/bin/python3 \\
        test/integration/h1_1_phase_a_breakdown.py [path]
"""
import logging
import os
import os.path as osp
import signal
import sys
import time

# Capture log.warning() output to stderr so [FORMULAS_TIMING] lines are visible.
logging.basicConfig(level=logging.WARNING, format='%(message)s')

DEFAULT_PATH = '/tmp/recalc_sanitized_1052.xlsx'
TIMEOUT_SEC = 1800


def _alarm(signum, frame):
    raise TimeoutError(f"hit {TIMEOUT_SEC}s timeout")


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_PATH
    if not osp.exists(path):
        print(f"ERROR: workbook not found: {path}", file=sys.stderr)
        return 2

    os.environ['FORMULAS_TIMING'] = '1'

    import formulas
    print(f"formulas: {osp.realpath(formulas.__file__)}", flush=True)
    print(f"workbook: {path}", flush=True)
    print(f"timeout:  {TIMEOUT_SEC}s", flush=True)
    print(f"FORMULAS_TIMING=1 — Phase A sub-step lines will appear below.", flush=True)
    print(flush=True)

    signal.signal(signal.SIGALRM, _alarm)
    signal.alarm(TIMEOUT_SEC)

    t0 = time.perf_counter()
    try:
        m = formulas.ExcelModel().loads(path).finish(
            circular=True, iterate=False
        )
        elapsed = time.perf_counter() - t0
        signal.alarm(0)
        print(f"\nfinish(circular=True, iterate=False) completed in {elapsed:.1f}s", flush=True)
        return 0
    except TimeoutError as e:
        elapsed = time.perf_counter() - t0
        print(f"\nTIMEOUT after {elapsed:.1f}s: {e}", flush=True)
        print("Phase A timing lines above identify how far we got.", flush=True)
        return 1
    except Exception as e:
        elapsed = time.perf_counter() - t0
        signal.alarm(0)
        print(f"\nEXCEPTION after {elapsed:.1f}s: {type(e).__name__}: {e}", flush=True)
        return 2


if __name__ == '__main__':
    sys.exit(main())
