# pyinstaller -F Spider.py -p util.py
# pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
# git rm -r --cached .
# git add .
# git commit -m "update"
# taskkill /im chromedriver.exe /f
# taskkill /im chrome.exe /f

# pip freeze >requirements.txt
# pip download -d packages -r requirements.txt
# pip install --no-index --find-links=packages -r requirements.txt

# C:\Windows\system32>FOR /F "usebackq delims=" %A IN (`python -c "from importlib import util;import os;print(os.path.join(os.path.dirname(util.find_spec('sasl').origin),'sasl2'))"`) DO (
#   REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Carnegie Mellon\Project Cyrus\SASL Library" /v SearchPath /t REG_SZ /d "%A"
# )
# https://raw.githubusercontent.com/publicsuffix/list/master/public_suffix_list.dat

# traceback.py
# def print_exception(etype, value, tb, limit=None, file=None, chain=True, url=None):
#     """Print exception up to 'limit' stack trace entries from 'tb' to 'file'.
#
#     This differs from print_tb() in the following ways: (1) if
#     traceback is not None, it prints a header "Traceback (most recent
#     call last):"; (2) it prints the exception type and value after the
#     stack trace; (3) if type is SyntaxError and value has the
#     appropriate format, it prints the line where the syntax error
#     occurred with a caret on the next line indicating the approximate
#     position of the error.
#     """
#     # format_exception has ignored etype for some time, and code such as cgitb
#     # passes in bogus values as a result. For compatibility with such code we
#     # ignore it here (rather than in the new TracebackException API).
#     if file is None:
#         file = sys.stderr
#     if url:
#        print("***** ", url, file=file, end="\n")
#     for line in TracebackException(
#             type(value), value, tb, limit=limit).format(chain=chain):
#         print(line, file=file, end="")

