#!/usr/bin/env python
# -*- coding: UTF-8 -*-
#
# Copyright 2016-2026 European Commission (JRC);
# Licensed under the EUPL (the 'Licence');
# You may not use this work except in compliance with the Licence.
# You may obtain a copy of the Licence at: http://ec.europa.eu/idabc/eupl
import os
import copy
import dill
import time
import json
import shutil
import unittest
import platform
import os.path as osp
import schedula as sh
from formulas.ranges import Ranges
from formulas.excel.xlreader import load_workbook
from formulas.excel import (
    ExcelModel, BOOK, ERR_CIRCULAR, _book2dict, _file2books, _convert_complex
)

EXTRAS = os.environ.get('EXTRAS', 'all')

mydir = osp.join(osp.dirname(__file__), 'test_files')
_filename = 'test.xlsx'
_filename_ods = 'test.ods'
_filename_compile = 'excel.xlsx'
_filename_full_range = 'full-range.xlsx'
_link_filename = 'test_link.xlsx'
_link_filename_ods = 'test_link.ods'
_filename_circular = 'circular.xlsx'
_filename_datatable = 'datatable.xlsx'
_filename_iter_divergent = 'iter_divergent.xlsx'
_filename_iter_linear = 'iter_convergent_linear.xlsx'
_filename_iter_nonlinear = 'iter_convergent_nonlinear.xlsx'


@unittest.skipIf(EXTRAS not in ('all', 'excel'), 'Not for extra %s.' % EXTRAS)
class TestExcelModel(unittest.TestCase):
    def setUp(self):
        self.filename = osp.join(mydir, _filename)
        self.link_filename = osp.join(mydir, _link_filename)
        self.filename_compile = osp.join(mydir, _filename_compile)
        self.filename_circular = osp.join(mydir, _filename_circular)
        self.filename_full_range = osp.join(mydir, _filename_full_range)
        self.filename_ods = osp.join(mydir, _filename_ods)
        self.link_filename_ods = osp.join(mydir, _link_filename_ods)

        self.results = _convert_complex(_file2books(
            self.filename, self.link_filename, _raw_data=True
        ))
        self.results_ods = _convert_complex(_file2books(
            self.filename_ods, self.link_filename_ods, _raw_data=True
        ))
        sh.get_nested_dicts(self.results, 'EXTRA.XLSX', 'EXTRA').update({
            'A1': 1, 'B1': 1
        })
        sh.get_nested_dicts(self.results, 'TEST.XLSX', 'EXTRA').update({
            'H1': 1, 'I1': 3, 'J1': 4, 'K1': 1, 'L1': 3
        })
        self.results_compile = _book2dict(
            load_workbook(self.filename_compile, data_only=True)
        )['DATA']
        self.results_circular = _convert_complex(_file2books(
            self.filename_circular, _raw_data=True
        ))
        self.results_full_range = {
            'TEST_FILES/%s' % k: v for k, v in _convert_complex(_file2books(
                self.filename_full_range, _raw_data=True
            )).items()}
        sh.get_nested_dicts(
            self.results_full_range, 'TEST_FILES/FULL-RANGE.XLSX', 'DATA'
        ).update({'A6': 5, 'A7': 1, 'B2': 19, 'B5': 63})
        self.maxDiff = None

    def _compare(self, xl_mdl, target, **kw):
        errors = xl_mdl.compare(target=target, absolute_tolerance=.0001, **kw)
        self.assertTrue('No differences.' == errors, errors)
        return len(tuple(sh.stack_nested_keys(target, depth=3)))

    def test_ods_model(self):
        start = time.time()
        _msg = '[info] test_ods_model: '
        xl_mdl = ExcelModel()

        print('\n%sLoading ods-model.' % _msg)
        s = time.time()

        xl_mdl.loads(self.filename_ods)

        msg = '%sLoaded ods-model in %.2fs.\n%sFinishing ods-model.'
        print(msg % (_msg, time.time() - s, _msg))
        s = time.time()

        xl_mdl.finish()

        print('%sFinished ods-model in %.2fs.' % (_msg, time.time() - s))

        n_test = 0

        print('%sCalculate ods-model.' % _msg)
        s = time.time()

        xl_mdl.calculate()

        msg = '%sCalculated ods-model in %.2fs.\n%s' \
              'Comparing overwritten results.'
        print(msg % (_msg, time.time() - s, _msg))
        s = time.time()

        n_test += self._compare(
            xl_mdl, self.results_ods, books=xl_mdl.books,
            solution=xl_mdl.dsp.solution
        )

        msg = '%sCompared overwritten results in %.2fs.\n' \
              '%sComparing fresh written results.'
        print(msg % (_msg, time.time() - s, _msg))
        s = time.time()

        n_test += self._compare(
            xl_mdl, self.results_ods, solution=xl_mdl.dsp.solution
        )

        msg = '%sCompared fresh written results in %.2fs.\n%sRan %d tests in %.2fs'
        print(msg % (_msg, time.time() - s, _msg, n_test, time.time() - start))

    def test_excel_model(self):
        start = time.time()
        _msg = '[info] test_excel_model: '
        xl_mdl = ExcelModel()

        print('\n%sLoading excel-model.' % _msg)
        s = time.time()

        xl_mdl.loads(self.filename)
        xl_mdl.add_book(self.link_filename)

        msg = '%sLoaded excel-model in %.2fs.\n%sFinishing excel-model.'
        print(msg % (_msg, time.time() - s, _msg))
        s = time.time()

        xl_mdl.finish()

        print('%sFinished excel-model in %.2fs.' % (_msg, time.time() - s))

        n_test, calculate = 0, True

        for i in range(4):
            if calculate:
                print('%sCalculate excel-model.' % _msg)
                s = time.time()

                xl_mdl({
                    "'[EXTRA.XLSX]EXTRA'!A1:B1": [[1, 1]],
                    "'[test.xlsx]EXTRA'!H1:I1": [[1, 3]]
                })

                msg = '%sCalculated excel-model in %.2fs.\n%s' \
                      'Comparing overwritten results.'
                print(msg % (_msg, time.time() - s, _msg))
                s = time.time()

                n_test += self._compare(
                    xl_mdl, self.results, books=xl_mdl.books,
                    solution=xl_mdl.dsp.solution
                )

                msg = '%sCompared overwritten results in %.2fs.\n' \
                      '%sComparing fresh written results.'
                print(msg % (_msg, time.time() - s, _msg))
                s = time.time()

                n_test += self._compare(
                    xl_mdl, self.results, solution=xl_mdl.dsp.solution
                )

                msg = '%sCompared fresh written results in %.2fs.'
                print(msg % (_msg, time.time() - s))
                calculate = False

            if i == 2 and platform.python_version() >= '3.8':
                print('%sSaving excel-model dill.' % _msg)
                s = time.time()

                xl_copy = dill.dumps(xl_mdl)

                msg = '%sSaved excel-model dill in %.2fs.\n' \
                      '%sLoading excel-model dill.'
                print(msg % (_msg, time.time() - s, _msg))
                s = time.time()

                xl_mdl = dill.loads(xl_copy)
                del xl_copy

                msg = '%sLoaded excel-model dill in %.2fs.'
                print(msg % (_msg, time.time() - s))
                calculate = True
            elif i == 1:
                print('%sDeep-copying excel-model.' % _msg)
                s = time.time()

                xl_mdl = copy.deepcopy(xl_mdl)

                msg = '%sDeep-copied excel-model in %.2fs.'
                print(msg % (_msg, time.time() - s))
                calculate = True
            elif i == 0:
                print('%sSaving JSON excel-model.' % _msg)
                s = time.time()

                xl_json = json.dumps(xl_mdl.to_dict())

                msg = '%sSaved JSON excel-model in %.2fs.\n' \
                      '%sLoading JSON excel-model.'
                print(msg % (_msg, time.time() - s, _msg))
                s = time.time()
                xl_mdl = ExcelModel().from_dict(json.loads(xl_json))
                del xl_json

                msg = '%sLoaded JSON excel-model in %.2fs.'
                print(msg % (_msg, time.time() - s))
                s = time.time()

                xl_mdl.finish(complete=False)

                print('%sFinished JSON excel-model in %.2fs.' % (
                    _msg, time.time() - s))
                calculate = True

        print('%sSaving excel-model xlsx.' % _msg)
        s = time.time()

        dirpath = osp.join(mydir, 'tmp')
        xl_mdl.write(dirpath=dirpath)

        msg = '%sSaved excel-model xlsx in %.2fs.\n%sComparing saved results.'
        print(msg % (_msg, time.time() - s, _msg))
        s = time.time()

        n_test += self._compare(
            xl_mdl, self.results, actual=_convert_complex(_file2books(*(
                osp.join(dirpath, fp) for fp in xl_mdl.books
            ), _raw_data=True))
        )

        msg = '%sCompared saved results in %.2fs.\n%sRan %d tests in %.2fs'
        print(msg % (_msg, time.time() - s, _msg, n_test, time.time() - start))

    def test_excel_model_compile(self):
        xl_model = ExcelModel().loads(self.filename_compile).finish()
        inputs = ["A%d" % i for i in range(2, 5)]
        outputs = ["C%d" % i for i in range(2, 5)]
        func = xl_model.compile(
            ["'[excel.xlsx]DATA'!%s" % i for i in inputs],
            ["'[excel.xlsx]DATA'!%s" % i for i in outputs]
        )
        i = sh.selector(inputs, self.results_compile, output_type='list')
        res = sh.selector(outputs, self.results_compile, output_type='list')
        self.assertEqual([x.value[0, 0] for x in func(*i)], res)
        func1 = xl_model.compile(
            ["'[excel.xlsx]DATA'!INPUT_%s" % i for i in "ABC"],
            ["'[excel.xlsx]DATA'!%s" % i for i in outputs]
        )
        self.assertEqual([x.value[0, 0] for x in func1(*i)], res)
        self.assertIsNot(xl_model, copy.deepcopy(xl_model))
        self.assertIsNot(func, copy.deepcopy(func))
        xl_model = ExcelModel().loads(self.filename_circular).finish(circular=1)
        func = xl_model.compile(
            ["'[circular.xlsx]DATA'!A10"],
            ["'[circular.xlsx]DATA'!E10"]
        )
        self.assertEqual(func(False).value[0, 0], 2.0)
        self.assertIs(func(True).value[0, 0], ERR_CIRCULAR)
        self.assertIsNot(xl_model, copy.deepcopy(xl_model))
        self.assertIsNot(func, copy.deepcopy(func))

    def test_excel_model_cycles(self):
        xl_model = ExcelModel().loads(self.filename_circular).finish(circular=1)
        xl_model.calculate()

        self._compare(
            xl_model, self.results_circular, actual=_convert_complex({
                k: _book2dict(v[BOOK])
                for k, v in xl_model.write(xl_model.books).items()
            })
        )

    def test_iterative_tarjan_no_recursion_limit(self):
        # AcquiOS fork 2026-05-29: _strong_connect is now iterative.
        # A 2000-node cycle would overflow Python's default recursion limit
        # (~1000 frames) under the original recursive implementation.
        # See specs/FORMULAS_ITERATIVE_CIRCULAR_SOLVER.md §4 Step 1.
        import sys
        from formulas.excel.cycle import _strongly_connected_components
        N = 2000
        graph = {i: {(i + 1) % N} for i in range(N)}
        orig_limit = sys.getrecursionlimit()
        sys.setrecursionlimit(1000)
        try:
            sccs = _strongly_connected_components(graph)
        finally:
            sys.setrecursionlimit(orig_limit)
        big = [s for s in sccs if len(s) > 1]
        self.assertEqual(len(big), 1)
        self.assertEqual(len(big[0]), N)

    def _solve_iterate(self, filename):
        # Helper: load fixture, run finish(circular=True, iterate=True),
        # return (solution_dict, last_solve_info). Solution values are
        # extracted as scalars for convenient assertions.
        path = osp.join(mydir, filename)
        xl_mdl = ExcelModel().loads(path).finish(circular=True, iterate=True)
        sol = xl_mdl.calculate()
        scalars = {}
        for k, v in sol.items():
            val = v.value if hasattr(v, 'value') else v
            if hasattr(val, 'tolist'):
                val = val.tolist()
            while isinstance(val, list) and len(val) == 1:
                val = val[0]
            scalars[str(k)] = val
        return scalars, xl_mdl._last_solve_info

    def test_iterate_divergent_does_not_crash(self):
        # A1=B1+1; B1=C1+1; C1=A1+1 has no fixed point. Must:
        #   - not raise
        #   - mark the SCC as not converged
        #   - return bounded last-iterate values (not inf/NaN)
        #   - leave non-cyclic cells (D1, E1) computing correctly
        scalars, info = self._solve_iterate(_filename_iter_divergent)
        self.assertEqual(scalars["'[iter_divergent.xlsx]S'!D1"], 42)
        self.assertEqual(scalars["'[iter_divergent.xlsx]S'!E1"], 84.0)
        self.assertEqual(len(info['sccs']), 1)
        scc = info['sccs'][0]
        self.assertFalse(scc['converged'])
        self.assertEqual(scc['size'], 3)
        # Bounded — Excel's iterate_count default is 100, so |x| <= ~100.
        for c in ('A1', 'B1', 'C1'):
            v = scalars[f"'[iter_divergent.xlsx]S'!{c}"]
            self.assertLess(abs(v), 1e6)

    def test_iterate_convergent_linear(self):
        # gross=1000; fee=0.05*EGI; EGI=gross-fee.
        # Closed-form: EGI=1000/1.05=952.380952...; fee=47.619047...
        scalars, info = self._solve_iterate(_filename_iter_linear)
        self.assertEqual(len(info['sccs']), 1)
        scc = info['sccs'][0]
        self.assertTrue(scc['converged'])
        self.assertEqual(scc['size'], 2)
        # Convergence to within Excel's default 0.001 tolerance.
        a2 = scalars["'[iter_convergent_linear.xlsx]S'!A2"]
        a3 = scalars["'[iter_convergent_linear.xlsx]S'!A3"]
        a4 = scalars["'[iter_convergent_linear.xlsx]S'!A4"]
        self.assertAlmostEqual(a2, 1000 * 0.05 / 1.05, delta=0.01)
        self.assertAlmostEqual(a3, 1000 / 1.05, delta=0.1)
        # Downstream cell A4=A3*2 must also appear in solution.
        self.assertAlmostEqual(a4, (1000 / 1.05) * 2, delta=0.2)

    def test_iterate_convergent_nonlinear(self):
        # fee=MIN(0.05*EGI, cap=40); EGI=gross-fee. MIN engages (5%*EGI=47.62 > 40),
        # so fee=40, EGI=960. Tests that a non-linear (MIN-bounded) cycle still
        # converges — either via fixed_point or broyden1 fallback.
        scalars, info = self._solve_iterate(_filename_iter_nonlinear)
        self.assertEqual(len(info['sccs']), 1)
        scc = info['sccs'][0]
        self.assertTrue(scc['converged'])
        a2 = scalars["'[iter_convergent_nonlinear.xlsx]S'!A2"]
        a3 = scalars["'[iter_convergent_nonlinear.xlsx]S'!A3"]
        self.assertAlmostEqual(a2, 40.0, delta=0.01)
        self.assertAlmostEqual(a3, 960.0, delta=0.1)

    def test_iterate_off_by_default_preserves_legacy_behavior(self):
        # Without iterate=True, finish(circular=True) should still produce
        # the original ERR_CIRCULAR-marked cells, proving we haven't disturbed
        # the existing graph-rewriting path.
        path = osp.join(mydir, _filename_iter_linear)
        xl_mdl = ExcelModel().loads(path).finish(circular=True)
        sol = xl_mdl.calculate()
        a3 = sol["'[iter_convergent_linear.xlsx]S'!A3"]
        v = a3.value if hasattr(a3, 'value') else a3
        if hasattr(v, 'tolist'):
            v = v.tolist()
        while isinstance(v, list) and len(v) == 1:
            v = v[0]
        # SCC cells carry the ERR_CIRCULAR sentinel (XlCircular, which
        # stringifies as "0" and is a subclass of XlError).
        self.assertIsInstance(v, ERR_CIRCULAR.__class__)
        # _last_solve_info must report iterate disabled and no SCCs solved.
        self.assertFalse(xl_mdl._last_solve_info['iterate_enabled'])
        self.assertEqual(xl_mdl._last_solve_info['sccs'], [])

    def test_iterate_method_used_pinned_for_linear(self):
        # Pin the solver method for the linear case so we detect silent
        # solver-strategy regressions (e.g. scipy.optimize.fixed_point
        # being downgraded behind the scenes).
        _, info = self._solve_iterate(_filename_iter_linear)
        self.assertEqual(info['sccs'][0]['method'], 'fixed_point')

    def test_excel_model_full_range(self):
        fname = osp.basename(self.filename_full_range)
        xl_model = ExcelModel()
        xl_model.basedir = osp.dirname(__file__)
        xl_model.complete([
            f"'test_files/[{fname}]DATA'!B5"
        ]).finish(complete=False)
        sheet_name = f'test_files/[{fname}]DATA'
        xl_model.calculate({
            f"'{sheet_name}'!A6": 5,
            f"'{sheet_name}'!A7": Ranges().push(f"'{sheet_name}'!A7", 1)
        })
        self._compare(
            xl_model, self.results_full_range, actual=_convert_complex({
                k: _book2dict(v[BOOK])
                for k, v in xl_model.write(xl_model.books).items()
            })
        )

    def test_excel_from_dict(self):
        xl_model = ExcelModel().from_dict({
            'A1': 1, 'B2': '=R[-1]C[-1]', 'A': 2, 'B': '=2', 'C': '=A1'
        }).finish()
        self.assertEqual({'A1': 2, 'B2': 2, 'A': 2, 'B': 2, 'C': 2}, {
            k: v.value.ravel()[0] if isinstance(v, Ranges) else v
            for k, v in xl_model.calculate({'C': 2}).items()
        })

        self.assertEqual(
            "\n\nErrors(4):\nAddition [D] -> 1\nChange [A1]: 2 -> 1\n"
            "Change [B2]: 2 -> 1\nChange [C]: 2 -> 1\n",
            xl_model.compare(
                target={'A1': 2, 'B2': 2, 'A': 2, 'B': 2, 'C': 2},
                actual={
                    k: v.value.ravel()[0] if isinstance(v, Ranges) else v
                    for k, v in xl_model.calculate({'D': 1}).items()
                }
            )
        )
        xl_model = ExcelModel().from_dict({
            "A1": 5, "A2": 0, "A3": 7, "B4": 1, "B5": 3, "A7": 3,
            "B1": "=SUM(A1:A2)", "C1": "=(B2 - B1)", "B2": "=SUM(A1:A8)",
            "A4:A5": "=B4:B5"
        }).finish(complete=False)
        self.assertEqual({
            "A1:A8": (1, 2, 3, 4, 5, 6, 7, 8), sh.SELF: xl_model.dsp,
            "A8": (8,), "A6": (6,), "A1": (1,), "A2": (2,), "A3": (3,),
            "A4:A5": (4, 5), "A7": (7,), "B2": (36,), "A1:A2": (1, 2),
            "B1": (3,), "C1": (33,), "B4": (1,), "B5": (3,), "B4:B5": (1, 3)
        }, {
            k: tuple(v.value.ravel()) if isinstance(v, Ranges) else v
            for k, v in xl_model({"A1:A8": [1, 2, 3, 4, 5, 6, 7, 8]}).items()
        })

    def test_excel_model_datatable(self):
        # Fork patch: an Excel Data Table (What-If) cell is stored as a
        # DataTableFormula with no `.text`. Before the patch this crashed load
        # with AttributeError. The patch substitutes an empty value and logs a
        # warning. Assert the model loads AND ordinary formulas still compute.
        fname = osp.join(mydir, _filename_datatable)
        xl_model = ExcelModel().loads(fname).finish()
        sol = xl_model.calculate()
        a2 = next(
            (v for k, v in sol.items() if str(k).upper().endswith('!A2')), None
        )
        self.assertIsNotNone(a2, 'A2 missing from solution')
        val = a2.value if isinstance(a2, Ranges) else a2
        self.assertEqual(float(val.ravel()[0]), 20.0)

    def tearDown(self) -> None:
        shutil.rmtree(osp.join(mydir, 'tmp'), ignore_errors=True)
