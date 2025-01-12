#!/usr/bin/env python

import pandas as pd

import openhtf as htf
from openhtf.output.callbacks import tonly_factory
from openhtf.output.callbacks import console_summary


from openhtf.plugs import user_input


@htf.measures(htf.Measurement("test_pass").in_range(minimum=10, maximum=20))
@htf.measures(htf.Measurement("test_fail").in_range(minimum=10, maximum=20))
@htf.measures(htf.Measurement("test_low_lim_only").in_range(minimum=10))
@htf.measures(htf.Measurement("test_high_lim_only").in_range(maximum=10))
@htf.measures(htf.Measurement("test_no_lim"))
@htf.measures(htf.Measurement("test_string"))
def numeric_phase(test):
    test.logger.info("Hello World!")
    test.measurements.test_pass = 11
    test.measurements.test_fail = 22
    test.measurements.test_low_lim_only = 12.0
    test.measurements.test_high_lim_only = 9.0
    test.measurements.test_no_lim = 10.0
    test.measurements.test_string = "hello world"


def attach_csv(test):
    df = pd.DataFrame(
        {
            "month": [1, 4, 7, 10],
            "year": [2012, 2014, 2013, 2014],
            "sale": [55, 40, 84, 31],
        }
    )
    csv = df.to_csv()
    test.attach("example_data.csv", csv.encode("utf-8"))


if __name__ == "__main__":
    test = htf.Test(
        numeric_phase,
        attach_csv,
        test_name="Tonly_Example_Report",
        test_version="1.2.3",
    )

    test.add_output_callbacks(
        tonly_factory.OutputToTonly(
            "{dut_id}_{metadata[test_name]}_{start_time_millis}.txt"
        )
    )

    test.add_output_callbacks(console_summary.ConsoleSummary())
    test.execute(test_start=user_input.prompt_for_test_start())
    # test.execute(test_start=dut_id_f())

    # xls_f = xls_factory.OutputToXLS('./blop.xls')
    # d = xls_f.convert_to_dict(test)
    # print(d)
