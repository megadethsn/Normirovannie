import os
import shutil
import sys


def app_dir():
    """Directory that contains the exe in production or sources in development."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def bundled_dir():
    """PyInstaller unpacked directory, if the app was built with --add-data."""
    return getattr(sys, "_MEIPASS", app_dir())


def templates_dir():
    path = os.path.join(app_dir(), "templates")
    if not os.path.isdir(path):
        bundled_templates = os.path.join(bundled_dir(), "templates")
        if os.path.isdir(bundled_templates):
            shutil.copytree(bundled_templates, path)
    os.makedirs(path, exist_ok=True)
    return path


def config_path():
    return os.path.join(app_dir(), "centers.json")


def log_path():
    return os.path.join(app_dir(), "app.log")


def resolve_template(filename):
    if not filename:
        return ""
    if os.path.isabs(filename):
        return filename
    return os.path.join(templates_dir(), filename)
