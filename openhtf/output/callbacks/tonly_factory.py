"""Module for outputting test record to XLS files."""

from openhtf.core import test_record
from openhtf.output import callbacks
from openhtf.util import data
import six
from io import StringIO
import pandas as pd
from collections import OrderedDict
import datetime
import platform
import os
import sys
from uuid import getnode
from datetime import datetime
import socket


class OutputToTonly(callbacks.OutputToFile):
    """Return an output callback that writes Tonly formatted reports.

    Example filename_patterns might be:
      '/data/test_records/{dut_id}.{metadata[test_name]}.txt')) or
      '/data/test_records/%(dut_id)s.%(start_time_millis)s'
    To use this output mechanism:
      test = openhtf.Test(PhaseOne, PhaseTwo)
      test.add_output_callback(openhtf.output.callbacks.OutputToTonly(
          '/data/test_records/{dut_id}.{metadata[test_name]}.txt'))

    Args:
      filename_pattern: A format string specifying the filename to write to. Must end with ".txt"
    """

    def __init__(self, filename_pattern=None, **kwargs):
        if not filename_pattern.endswith(".txt"):
            raise AssertionError(
                "Invalid filename pattern: %s" % repr(filename_pattern)
            )
        super(OutputToTonly, self).__init__(filename_pattern)

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

    def long_break(self, report):
        report.write("===================================\n")

    def section_break(self, report):
        report.write("===============================\n")

    def phase_break(self, report):
        report.write("===============\n")

    def linefeed(self, report):
        report.write("\n")

    def timestamp_ms_to_date_str(self, ts_ms):
        ts_ms = int(ts_ms)
        t = datetime.fromtimestamp(ts_ms / 1000.0)
        return str(t.strftime("%Y/%m/%d %H:%M:%S"))

    def write_tonly_report(self, test_record, report):
        rec = self.convert_to_dict(test_record)
        print(rec["metadata"])

        self.section_break(report)

        # Operating system
        report.write("操作系统版本：%s\n" % platform.platform())

        # User name
        report.write("用户名称：%s\n" % os.getlogin())

        # Program name
        if "test_name" in rec["metadata"]:
            report.write("程序名称：%s\n" % rec["metadata"]["test_name"])
        else:
            report.write("程序名称：unknown\n")

        # Program version
        if "test_version" in rec["metadata"]:
            report.write("程序版本：%s\n" % rec["metadata"]["test_version"])
        else:
            report.write("程序版本：unknown\n")

        # Config file - not implemented
        report.write("配置文件: Not Implemented\n")

        # Config file verification
        report.write("配置文件校验: Not Implemented\n")

        self.section_break(report)

        # Batch number?
        report.write("程序设置的批次号：\n")

        # Model name?
        report.write("程序设置的机型名：\n")

        # Component number
        report.write("程序设置的组件号：\n")

        # Station name
        report.write("程序设置的工站名：\n")

        # Computer name + MAC
        mac = getnode()
        mac_str = "-".join(("%012X" % mac)[i : i + 2] for i in range(0, 12, 2))
        print("Mac str: %s" % mac_str)
        hostname = socket.gethostname()
        print("hostname: %s" % hostname)
        report.write("机架号及穴位号：%s_%s\n" % (hostname, mac_str))

        self.long_break(report)

        # Test start time
        start_t_str = self.timestamp_ms_to_date_str(test_record.start_time_millis)
        report.write("开始测试时间: ")
        report.write(start_t_str + "\n")

        # Test end time
        end_t_str = self.timestamp_ms_to_date_str(test_record.end_time_millis)
        report.write("结束测试时间: ")
        report.write(end_t_str + "\n")

        # Test duration in seconds
        duration_ms = test_record.end_time_millis - test_record.start_time_millis
        report.write("总的测试时间: ")
        report.write("%0.1f S\n" % (duration_ms / 1000.0))

        self.long_break(report)
        self.linefeed(report)
        self.long_break(report)

        # Start test message
        report.write(start_t_str + "\n")
        self.linefeed(report)
        report.write("Test Start !\n")
        self.phase_break(report)

        # Print out the phases one by one
        phases = rec["phases"]
        print(phases)
        print("")
        report.write("phases\n")
        for phase in phases:
            measurements = phase["measurements"]
            for m in measurements:
                print("measurement: %s" % m)
                report.write("boop\n")
                report.write("measurement: %s" % m)
                value = measurements[m]["measured_value"]
                if "validators" in measurements[m]:
                    v = measurements[m]["validators"]
                else:
                    v = []
                (low_lim, high_lim) = self.validators_to_limits(v)
                pass_fail = measurements[m]["outcome"]

                # report.write("%s: %f (%s) - %s\n" % (m, value, v, pass_fail))

        report.write(test_record.dut_id + "\n")
        report.write("bloop\n")
        report.write("blop\n")

    def convert_to_dict(self, test_record):
        return data.convert_to_base_types(test_record)

    def __call__(self, test_record):
        filename = self.create_file_name(test_record)
        if test_record.dut_id not in ["exit", "quit", "EXIT", "QUIT"]:
            with open(filename, "w", encoding="utf-8") as report:
                self.write_tonly_report(test_record, report)
