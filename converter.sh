#!/bin/sh

# конвертация *.ui --> *.py
pyuic5 main.ui -o main.py
pyuic5 settings.ui -o settings.py
