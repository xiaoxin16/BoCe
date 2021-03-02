#!/bin/python3

import datetime
import random
from webbrowser import Chrome

import dns.resolver
import platform
import ipdb
import json
import math
import os
import re
import threading
import time
from datetime import datetime
from urllib.parse import urlparse

import requests
import tldextract
from PIL import Image, ImageDraw, ImageFont
import xlrd
import docx
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side, colors
from requests.adapters import HTTPAdapter
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, WebDriverException, \
    NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.by import By
import sqlite3

# 读取配置文件json
from selenium.webdriver.support.wait import WebDriverWait


def main():
    print("start...")


if __name__ == "__main__":
    main()