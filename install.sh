#!/bin/bash

pip uninstall openhtf -y
python setup.py build
python setup.py install

