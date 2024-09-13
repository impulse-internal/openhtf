"""Module for outputting test record to XLS files."""

from openhtf.core import test_record
from openhtf.output import callbacks
from openhtf.util import data
import six
from io import StringIO
from io import BytesIO
import pandas as pd
from collections import OrderedDict
import datetime
import base64


class OutputToXLSX(callbacks.OutputToFile):
    """Return an output callback that writes Excel Test Records.

    Example filename_patterns might be:
      '/data/test_records/{dut_id}.{metadata[test_name]}.xls', indent=4)) or
      '/data/test_records/%(dut_id)s.%(start_time_millis)s'
    To use this output mechanism:
      test = openhtf.Test(PhaseOne, PhaseTwo)
      test.add_output_callback(openhtf.output.callbacks.OutputToXLS(
          '/data/test_records/{dut_id}.{metadata[test_name]}.xls'))

    Args:
      filename_pattern: A format string specifying the filename to write to. Must end with ".xlsx"
    """

    def __init__(self, filename_pattern=None, inline_attachments=True, **kwargs):
        if not filename_pattern.endswith(".xlsx"):
            raise AssertionError(
                "Invalid filename pattern: %s" % repr(filename_pattern)
            )
        super(OutputToXLSX, self).__init__(filename_pattern)
        self.inline_attachments = inline_attachments

    def validators_to_limits(self, validators):
        low_lim = None
        high_lim = None
        if len(validators) == 1:
            v = validators[0].split(" ")
            if v[1] == "<=" and v[2] == "x":
                low_lim = float(v[0])
            if v[0] == "x" and v[1] == "<=":
                high_lim = float(v[2])
            if len(v) == 5:
                if v[2] == "x" and v[3] == "<=":
                    high_lim = float(v[4])
        return (low_lim, high_lim)

    def write_sheets(self, test_record, writer):
        logs_sheet_name = "Test Logs"
        test_rec_sheet_name = "Test Record"

        # Defining cell format
        workbook = writer.book
        cell_format = workbook.add_format()
        cell_format.set_align("left")

        rec = self.convert_to_dict(test_record)
        start_t_ms = rec["start_time_millis"]
        start_t_date = datetime.datetime.fromtimestamp(start_t_ms / 1000.0)
        try:
            test_name = rec["metadata"]["test_name"]
        except:
            test_version = "Unset"
        try:
            test_version = rec["metadata"]["test_version"]
        except:
            test_version = "Unset"

        header_dict = OrderedDict(
            [
                ("Test Name", test_name),
                ("Test Station", [rec["station_id"]]),
                ("Software Version", test_version),
                ("DUT ID", [rec["dut_id"]]),
                ("Start Time", [start_t_date]),
                ("Start Time (ms)", [start_t_ms]),
                ("Test Result", [rec["outcome"]]),
            ]
        )
        df = pd.DataFrame.from_dict(header_dict)
        df.to_excel(writer, sheet_name=test_rec_sheet_name, index=False)

        # Add conditional formatting to the sheet for PASS / FAIL

        # Add formats to the workbook
        # Light red fill with dark red text.
        fail_format = workbook.add_format(
            {"bg_color": "#FFC7CE", "font_color": "#9C0006"}
        )
        # Light green fill with dark green text.
        pass_format = workbook.add_format(
            {"bg_color": "#C6EFCE", "font_color": "#006100"}
        )

        # Write a conditional formats to the records sheet.
        writer.sheets[test_rec_sheet_name].conditional_format(
            "A1:Z1000",
            {
                "type": "cell",
                "criteria": "==",
                "value": '"PASS"',
                "format": pass_format,
            },
        )
        writer.sheets[test_rec_sheet_name].conditional_format(
            "A1:Z1000",
            {
                "type": "cell",
                "criteria": "==",
                "value": '"FAIL"',
                "format": fail_format,
            },
        )

        xls_dict = OrderedDict(
            [
                ("Measurement", []),
                ("Value", []),
                ("Low Limit", []),
                ("High Limit", []),
                ("Pass/Fail", []),
            ]
        )

        # Iterate throuhg the measurements and build them into a table
        phases = rec["phases"]
        for phase in phases:
            measurements = phase["measurements"]
            for m in measurements:
                value = measurements[m]["measured_value"]
                if "validators" in measurements[m]:
                    v = measurements[m]["validators"]
                else:
                    v = []
                (low_lim, high_lim) = self.validators_to_limits(v)
                pass_fail = measurements[m]["outcome"]
                xls_dict["Value"].append(value)
                xls_dict["Measurement"].append(m)
                xls_dict["Low Limit"].append(low_lim)
                xls_dict["High Limit"].append(high_lim)
                xls_dict["Pass/Fail"].append(pass_fail)

        df = pd.DataFrame.from_dict(xls_dict)
        df.to_excel(writer, sheet_name=test_rec_sheet_name, startrow=3, index=False)

        # Insert any csv attachments as extra sheets
        phases = rec["phases"]
        for phase in phases:
            attachments = phase["attachments"]
            for a in attachments:
                if a.endswith(".csv"):
                    csv_data = attachments[a].data.decode("utf-8")
                    df = pd.read_csv(StringIO(csv_data))
                    data_sheet_name = a.replace(".csv", "")
                    df.to_excel(writer, sheet_name=data_sheet_name, index=False)
                    data_sheet = writer.sheets[data_sheet_name]
                    data_sheet.set_column(0, 10, 15, cell_format)
                if a.endswith(".png"):
                    try:
                        print("image: %s" % a)
                        attachment = attachments[a]
                        image_data = BytesIO(attachment.data)
                        image_sheet = workbook.add_worksheet(a)
                        image_sheet.insert_image("A1", a, {"image_data": image_data})
                    except Exception as e:
                        print(e)
        # Insert tester logs as a sheet
        log_fields = [
            "level",
            "logger_name",
            "source",
            "lineno",
            "timestamp_millis",
            "millis_since_test_start",
            "message",
        ]
        log_dict = OrderedDict()
        for f in log_fields:
            log_dict[f] = []

        logs = rec["log_records"]
        for log in logs:
            for f in log_fields:
                if f == "millis_since_test_start":
                    # Add column for times from test start
                    log_dict[f].append(
                        log["timestamp_millis"] - rec["start_time_millis"]
                    )
                else:
                    log_dict[f].append(log[f])

        df = pd.DataFrame.from_dict(log_dict)
        df.to_excel(writer, sheet_name=logs_sheet_name, index=False)

        logs_sheet = writer.sheets[logs_sheet_name]
        column_widths = [10, 30, 30, 10, 20, 20, 100]
        for col, width in enumerate(column_widths):
            logs_sheet.set_column(col, col, width, cell_format)

        test_rec_sheet = writer.sheets[test_rec_sheet_name]
        column_widths = [30, 20, 20, 20, 20, 20, 20]
        for col, width in enumerate(column_widths):
            test_rec_sheet.set_column(col, col, width, cell_format)

        return str(xls_dict)

    def convert_to_dict(self, test_record):
        as_dict = data.convert_to_base_types(test_record)
        if self.inline_attachments:
            for phase, original_phase in zip(as_dict["phases"], test_record.phases):
                for name, attachment in six.iteritems(original_phase.attachments):
                    phase["attachments"][name] = attachment
        return as_dict

    def __call__(self, test_record):
        filename = self.create_file_name(test_record)
        if test_record.dut_id not in ["exit", "quit", "EXIT", "QUIT"]:
            with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
                try:
                    self.write_sheets(test_record, writer)
                except Exception as e:
                    print(e)
