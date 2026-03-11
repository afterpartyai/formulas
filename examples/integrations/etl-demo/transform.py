#!/usr/bin/env python

import json
import sys
from pathlib import Path

import formulas


def extract_value(value):
    if isinstance(value, formulas.Ranges):
        data = value.value.tolist()
        if len(data) == 1 and len(data[0]) == 1:
            return data[0][0]
        return data
    return value


def iter_jsonl(path):
    with open(path) as fh:
        for line in fh:
            line = line.strip()
            if line:
                yield json.loads(line)


def main():
    root = Path(__file__).resolve().parents[3]
    workbook = root / 'test' / 'test_files' / 'excel.xlsx'
    input_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name('input.jsonl')

    model = formulas.ExcelModel().loads(str(workbook)).finish()
    output_ref = "'[excel.xlsx]DATA'!C2"

    for record in iter_jsonl(input_path):
        solution = model.calculate(
            inputs={
                "'[excel.xlsx]'!INPUT_A": record['input_a'],
                "'[excel.xlsx]DATA'!B3": record['b3'],
            },
            outputs=[output_ref]
        )
        transformed = {
            'id': record['id'],
            'result': extract_value(solution[output_ref]),
        }
        print(json.dumps(transformed, sort_keys=True))


if __name__ == '__main__':
    main()
