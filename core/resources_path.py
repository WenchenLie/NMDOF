import sys
import os


def resources_path(rela_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath('.')
    return os.path.join(base_path, rela_path).replace('\\', '/')