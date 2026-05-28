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


def assets_dir():
    path = os.path.join(app_dir(), "assets")
    if not os.path.isdir(path):
        bundled_assets = os.path.join(bundled_dir(), "assets")
        if os.path.isdir(bundled_assets):
            shutil.copytree(bundled_assets, path)
    os.makedirs(path, exist_ok=True)
    return path


def resolve_asset(filename):
    if not filename:
        return ""
    if os.path.isabs(filename):
        return filename
    return os.path.join(assets_dir(), filename)


def config_path():
    return os.path.join(app_dir(), "centers.json")


def people_config_path():
    return os.path.join(app_dir(), "people.json")


def log_path():
    return os.path.join(app_dir(), "app.log")


def resolve_template(filename):
    if not filename:
        return ""
    if os.path.isabs(filename):
        return filename
    return os.path.join(templates_dir(), filename)
