from distutils.core import setup
import py2exe

setup(windows=['zhw.py'], options={'py2exe':  {'includes':['sip']}})