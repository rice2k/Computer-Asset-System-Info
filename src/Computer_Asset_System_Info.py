import argparse
import csv
import ctypes
import getpass
import io
import json
import os
import platform
import queue
import re
import socket
import subprocess
import sys
import threading
import textwrap
import time
import urllib.request
import webbrowser
import winreg
from collections import OrderedDict
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import BooleanVar, Menu, StringVar, Tk, filedialog, messagebox
import tkinter as tk
from tkinter import ttk

try:
    import psutil
except Exception:
    psutil = None

try:
    import pyperclip
except Exception:
    pyperclip = None

try:
    import wmi
except Exception:
    wmi = None

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None

try:
    from PIL import Image, ImageDraw, ImageGrab, ImageTk
except Exception:
    Image = None
    ImageDraw = None
    ImageGrab = None
    ImageTk = None


APP_TITLE = "Computer Asset System Info"
APP_VERSION = "3.7"
OUTPUT_PREFIX = "ComputerAsset"
ICON_PNG = "computer_asset_icon.png"
ICON_ICO = "computer_asset_icon.ico"
SETTINGS_FILE = "ComputerAssetSystemInfo.settings.json"

SIMPLE_SECTIONS = (
    "Dashboard",
    "Helpdesk",
    "Problems",
    "Overview",
    "Asset & Ownership",
    "Security",
    "Hardware",
    "Storage",
    "Network",
)

TAB_LABELS = {
    "Asset & Ownership": "Asset",
    "Operating System": "Windows",
    "Windows & Updates": "Updates",
    "Access & Startup": "Access",
    "Peripherals & Shares": "Peripherals",
    "Activity & Health": "Health",
}

NAV_ICON_KEYS = {
    "Dashboard": "dashboard",
    "Problems": "warning",
    "Helpdesk": "helpdesk",
    "Search": "search",
    "Overview": "overview",
    "Asset & Ownership": "asset",
    "Security": "shield",
    "Hardware": "chip",
    "Storage": "storage",
    "Network": "wifi",
    "About": "help",
    "Operating System": "windows",
    "Windows & Updates": "updates",
    "Management": "gear",
    "Access & Startup": "key",
    "Applications": "apps",
    "Peripherals & Shares": "plug",
    "Activity & Health": "health",
}

TOOLTIP_TEXT = {
    "TPM Status": "Trusted Platform Module. This hardware chip helps Windows protect encryption keys and sign-in secrets.",
    "TPM": "Trusted Platform Module. This hardware chip helps Windows protect encryption keys and sign-in secrets.",
    "BitLocker Status": "Windows drive encryption status. It helps protect data if a device is lost or stolen.",
    "MDM": "Mobile Device Management. Usually Intune or another tool that manages settings, apps, and compliance.",
    "Intune / MDM Enrollment": "Shows whether this computer appears enrolled in a device management system.",
    "SMART": "Drive health reporting built into many disks. It can warn about possible disk failure.",
    "Disk SMART Health": "Drive health reporting built into many disks. It can warn about possible disk failure.",
    "Secure Boot": "A firmware security feature that helps block untrusted startup software.",
    "EDR / Security Agent Status": "Endpoint Detection and Response. This is the advanced security agent used by many IT teams.",
    "Defender Definition Age": "How old the Microsoft Defender security definitions are.",
    "Pending Reboot": "Windows or an installer has changes waiting for a restart.",
    "Azure AD Joined": "Shows whether the device is joined to Microsoft Entra ID, formerly Azure Active Directory.",
    "SCCM / ConfigMgr Client": "Microsoft Configuration Manager client status, if used by the organization.",
    "Wi-Fi SSID": "The wireless network name. Some Windows permissions or adapters may hide this even while the computer has a valid IP address.",
    "Primary IP Address": "The main network address Windows is using. A valid IP usually means the machine is connected.",
}

SEARCH_TAGS = {
    "wifi": ("wifi", "wi-fi", "wireless", "ssid", "network", "internet", "ip"),
    "network": ("network", "wifi", "wi-fi", "ip", "mac", "adapter", "traffic", "download", "upload"),
    "speed": ("speed", "download", "upload", "internet", "network", "traffic", "mbps"),
    "disk": ("disk", "drive", "storage", "c:", "free", "used", "space", "smart"),
    "storage": ("storage", "disk", "drive", "space", "free", "used", "smart"),
    "battery": ("battery", "charge", "health", "power"),
    "security": ("security", "antivirus", "defender", "firewall", "bitlocker", "tpm", "secure boot", "edr"),
    "antivirus": ("antivirus", "defender", "virus", "security", "protection"),
    "updates": ("update", "updates", "windows update", "patch", "reboot", "hotfix"),
    "windows": ("windows", "os", "build", "activation", "license", "edition"),
    "hardware": ("hardware", "cpu", "ram", "memory", "bios", "serial", "model"),
    "asset": ("asset", "tag", "assigned", "owner", "department", "location", "warranty"),
    "helpdesk": ("helpdesk", "support", "computer name", "model", "serial", "ip", "os"),
}

TAG_MENU = ("wifi", "speed", "disk", "battery", "security", "updates", "windows", "hardware", "asset", "helpdesk")

SPEED_TEST_DOWNLOAD_SOURCES = (
    ("Cloudflare", "https://speed.cloudflare.com/__down?bytes=25000000", 12_000_000, 12),
    ("Cloudflare fallback", "https://speed.cloudflare.com/__down?bytes=10000000", 8_000_000, 12),
    ("OVH fallback", "https://proof.ovh.net/files/10Mb.dat", 5_000_000, 15),
)
SPEED_TEST_UPLOAD_URL = "https://speed.cloudflare.com/__up"

HELPDESK_FIELDS = (
    ("Overview", "Computer Name"),
    ("Asset & Ownership", "Asset Tag"),
    ("Asset & Ownership", "Assigned User"),
    ("Overview", "Manufacturer"),
    ("Overview", "Model"),
    ("Overview", "Serial Number"),
    ("Overview", "Operating System"),
    ("Overview", "Windows Release"),
    ("Operating System", "Build Number"),
    ("Hardware", "RAM Installed"),
    ("Hardware", "CPU Name"),
    ("Storage", "C: Free Space"),
    ("Network", "Primary IP Address"),
    ("Network", "Wi-Fi SSID"),
    ("Security", "BitLocker Status"),
    ("Security", "Virus Protection"),
    ("Security", "Pending Reboot"),
)

SIMPLE_FIELDS = {
    "Overview": (
        "Computer Name",
        "Asset Tag",
        "Assigned User",
        "Manufacturer",
        "Model",
        "Serial Number",
        "Operating System",
        "Windows Release",
        "RAM Installed",
        "C: Free Space",
        "Security Agent",
        "Pending Reboot",
    ),
    "Asset & Ownership": (
        "Asset Tag",
        "Assigned User",
        "Department",
        "Device Location",
        "Warranty Expiration",
        "Device Age",
        "Manufacturer",
        "Model",
        "Serial Number",
    ),
    "Security": (
        "Virus Protection",
        "Firewall Profile Status",
        "Microsoft Defender Status",
        "Defender Definition Age",
        "EDR / Security Agent Status",
        "BitLocker Status",
        "TPM Status",
        "Secure Boot",
        "Pending Reboot",
    ),
    "Hardware": (
        "Manufacturer",
        "Model",
        "Case Type",
        "Asset Tag",
        "Device Age",
        "CPU Name",
        "RAM Installed",
        "Battery Health",
    ),
    "Storage": ("C: Total Space", "C: Used Space", "C: Free Space", "C: Percent Used"),
    "Network": ("Hostname", "Primary IP Address", "Wi-Fi SSID", "Machine Domain", "Part of Domain", "Admin User"),
}

SENSITIVE_TERMS = (
    "serial",
    "ip address",
    "ipv4",
    "ipv6",
    "mac address",
    "ssid",
    "user id",
    "user domain",
    "logged-in user",
    "assigned user",
    "last logged-in user",
    "registered owner",
    "owner",
    "computer name",
    "hostname",
    "domain controller",
    "machine domain",
    "local path",
    "network path",
    "share path",
    "profile path",
)

WMI_CACHE = {}


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(name):
    external = app_dir() / name
    if external.exists():
        return external
    bundled = Path(getattr(sys, "_MEIPASS", app_dir())) / name
    return bundled


def settings_path():
    return Path.home() / "Documents" / SETTINGS_FILE


def pictures_folder():
    candidates = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders") as key:
            value, _kind = winreg.QueryValueEx(key, "My Pictures")
            candidates.append(Path(os.path.expandvars(value)))
    except Exception:
        pass
    one_drive = os.environ.get("OneDrive")
    if one_drive:
        candidates.append(Path(one_drive) / "Pictures")
    candidates.append(Path.home() / "Pictures")
    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            return candidate
        except Exception:
            continue
    fallback = Path.home() / "Pictures"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback


def load_settings():
    defaults = {
        "privacy": False,
        "advanced": False,
        "dark_mode": False,
        "export_folder": str(Path.home() / "Documents"),
    }
    try:
        path = settings_path()
        if path.exists():
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                defaults.update({key: data[key] for key in defaults if key in data})
    except Exception:
        pass
    return defaults


def save_settings(values):
    try:
        settings_path().write_text(json.dumps(values, indent=2), encoding="utf-8")
    except Exception:
        pass


def now_stamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def safe_filename(value, fallback="computer"):
    value = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "").strip()).strip("_")
    return value or fallback


def clean(value, default="Unavailable"):
    if value is None:
        return default
    value = str(value).strip()
    return value if value else default


def yes_no(value):
    if value is None:
        return "Unavailable"
    return "Yes" if bool(value) else "No"


def format_gb(value):
    try:
        return f"{float(value) / (1024 ** 3):.2f} GB"
    except Exception:
        return "Unavailable"


def format_uptime(seconds):
    try:
        return str(timedelta(seconds=int(seconds)))
    except Exception:
        return "Unavailable"


def wmi_datetime(value):
    value = clean(value, "")
    if not value:
        return "Unavailable"
    try:
        return datetime.strptime(value[:14], "%Y%m%d%H%M%S").strftime("%Y-%m-%d %I:%M:%S %p")
    except Exception:
        return value


def parse_date(value):
    value = clean(value, "")
    if not value:
        return None
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m-%d-%Y"):
        try:
            return datetime.strptime(value, fmt)
        except Exception:
            pass
    return None


def parse_wmi_datetime(value):
    value = clean(value, "")
    if not value or value == "Unavailable":
        return None
    try:
        return datetime.strptime(value[:14], "%Y%m%d%H%M%S")
    except Exception:
        return None


def age_since(value):
    if isinstance(value, str):
        value = parse_wmi_datetime(value) or parse_date(value)
    if not value:
        return "Unavailable"
    try:
        days = max((datetime.now() - value).days, 0)
        years = days // 365
        months = (days % 365) // 30
        if years:
            return f"{years} years, {months} months"
        return f"{months} months" if months else f"{days} days"
    except Exception:
        return "Unavailable"


def run_command(args, timeout=12):
    try:
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = subprocess.run(args, capture_output=True, text=True, timeout=timeout, creationflags=flags)
        return (result.stdout or "") + (result.stderr or "")
    except Exception as exc:
        return f"Unavailable: {exc}"


def run_powershell(script, timeout=15):
    return run_command(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        timeout=timeout,
    )


def powershell_json(script, timeout=20):
    out = run_powershell(script, timeout=timeout).strip()
    if not out:
        return None
    start = min([pos for pos in [out.find("{"), out.find("[")] if pos >= 0] or [0])
    text = out[start:]
    try:
        return json.loads(text)
    except Exception:
        return None


def as_list(value):
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def get_wmi(namespace="root\\cimv2"):
    if wmi is None:
        return None
    if namespace not in WMI_CACHE:
        try:
            WMI_CACHE[namespace] = wmi.WMI(namespace=namespace)
        except Exception:
            WMI_CACHE[namespace] = None
    return WMI_CACHE[namespace]


def wmi_first(class_name, namespace="root\\cimv2"):
    client = get_wmi(namespace)
    if client is None:
        return None
    try:
        rows = getattr(client, class_name)()
        return rows[0] if rows else None
    except Exception:
        return None


def wmi_all(class_name, namespace="root\\cimv2"):
    client = get_wmi(namespace)
    if client is None:
        return []
    try:
        return list(getattr(client, class_name)())
    except Exception:
        return []


def wmi_query(query, namespace="root\\cimv2"):
    client = get_wmi(namespace)
    if client is None:
        return []
    try:
        return list(client.query(query))
    except Exception:
        return []


def attr(obj, name, default="Unavailable"):
    try:
        return clean(getattr(obj, name), default)
    except Exception:
        return default


def registry_get(root, path, name, default="Unavailable"):
    try:
        with winreg.OpenKey(root, path) as key:
            value, _ = winreg.QueryValueEx(key, name)
            return clean(value, default)
    except Exception:
        return default


def registry_key_exists(root, path):
    try:
        with winreg.OpenKey(root, path):
            return True
    except Exception:
        return False


def scan_uninstall_apps():
    apps = []
    hives = (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER)
    paths = (
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    )
    for hive in hives:
        for path in paths:
            try:
                with winreg.OpenKey(hive, path) as key:
                    count = winreg.QueryInfoKey(key)[0]
                    for index in range(count):
                        try:
                            sub_name = winreg.EnumKey(key, index)
                            with winreg.OpenKey(key, sub_name) as sub_key:
                                name = registry_value(sub_key, "DisplayName")
                                if not name:
                                    continue
                                apps.append(
                                    {
                                        "Name": name,
                                        "Version": registry_value(sub_key, "DisplayVersion") or "Unavailable",
                                        "Publisher": registry_value(sub_key, "Publisher") or "Unavailable",
                                    }
                                )
                        except Exception:
                            continue
            except Exception:
                continue
    return apps


def registry_value(open_key, name):
    try:
        value, _ = winreg.QueryValueEx(open_key, name)
        return clean(value, "")
    except Exception:
        return ""


def find_installed(apps, terms):
    terms = [term.lower() for term in terms]
    for app in apps:
        name = app.get("Name", "").lower()
        if any(term in name for term in terms):
            version = clean(app.get("Version"), "")
            return f"Installed{f' ({version})' if version and version != 'Unavailable' else ''}"
    return "Not installed"


def decode_wmi_text(values):
    try:
        return "".join(chr(int(v)) for v in values if int(v) > 0).strip()
    except Exception:
        return ""


def chassis_type_label(value):
    labels = {
        1: "Other",
        2: "Unknown",
        3: "Desktop",
        4: "Low Profile Desktop",
        5: "Pizza Box",
        6: "Mini Tower",
        7: "Tower",
        8: "Portable",
        9: "Laptop",
        10: "Notebook",
        11: "Hand Held",
        12: "Docking Station",
        13: "All-in-One",
        14: "Sub-Notebook",
        15: "Space-saving",
        16: "Lunch Box",
        17: "Main System Chassis",
        30: "Tablet",
        31: "Convertible",
        32: "Detachable",
    }
    try:
        if isinstance(value, (list, tuple)):
            value = value[0] if value else None
        return labels.get(int(value), f"Type {value}")
    except Exception:
        return "Unavailable"


def current_user():
    try:
        return os.getlogin()
    except Exception:
        return getpass.getuser()


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def windows_release_info():
    base = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    display = registry_get(winreg.HKEY_LOCAL_MACHINE, base, "DisplayVersion", "")
    release = registry_get(winreg.HKEY_LOCAL_MACHINE, base, "ReleaseId", "")
    build = registry_get(winreg.HKEY_LOCAL_MACHINE, base, "CurrentBuildNumber", "")
    ubr = registry_get(winreg.HKEY_LOCAL_MACHINE, base, "UBR", "")
    edition = registry_get(winreg.HKEY_LOCAL_MACHINE, base, "EditionID", "")
    product = registry_get(winreg.HKEY_LOCAL_MACHINE, base, "ProductName", "")
    build_full = f"{build}.{ubr}" if build and ubr else build
    return {
        "Product": product,
        "Edition": edition,
        "Release": display or release or "Unavailable",
        "Build": build_full or "Unavailable",
    }


def pending_reboot_status():
    checks = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending", "Component servicing"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired", "Windows Update"),
    ]
    reasons = [label for root, path, label in checks if registry_key_exists(root, path)]
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager") as key:
            winreg.QueryValueEx(key, "PendingFileRenameOperations")
            reasons.append("Pending file rename")
    except Exception:
        pass
    return "Yes: " + ", ".join(reasons) if reasons else "No"


def last_hotfix():
    rows = []
    for fix in wmi_all("Win32_QuickFixEngineering"):
        installed = attr(fix, "InstalledOn", "")
        date = parse_date(installed)
        rows.append((date or datetime.min, installed, attr(fix, "HotFixID"), attr(fix, "Description")))
    if not rows:
        return "Unavailable"
    rows.sort(key=lambda item: item[0], reverse=True)
    _, installed, hotfix, description = rows[0]
    return f"{installed} ({hotfix}, {description})"


def activation_status():
    status_map = {
        0: "Unlicensed",
        1: "Licensed",
        2: "Out-of-box grace",
        3: "Out-of-tolerance grace",
        4: "Non-genuine grace",
        5: "Notification",
        6: "Extended grace",
    }
    products = wmi_query("SELECT Name, LicenseStatus, PartialProductKey FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL")
    for product in products:
        name = attr(product, "Name", "")
        if "windows" in name.lower():
            try:
                code = int(getattr(product, "LicenseStatus"))
            except Exception:
                code = -1
            return status_map.get(code, f"Status {code}")
    return "Unavailable"


def secure_boot_status():
    out = run_powershell("try { Confirm-SecureBootUEFI } catch { 'Unsupported or unavailable' }", timeout=10)
    value = out.strip().splitlines()[-1] if out.strip() else "Unavailable"
    if value.lower() == "true":
        return "Enabled"
    if value.lower() == "false":
        return "Disabled"
    return value


def tpm_status():
    out = run_powershell(
        "try { Get-Tpm | Select-Object TpmPresent,TpmReady,TpmEnabled,TpmActivated | ConvertTo-Json -Compress } catch { '{}' }",
        timeout=10,
    ).strip()
    try:
        data = json.loads(out)
        if not data:
            return "Unavailable"
        return ", ".join(f"{key}={yes_no(value)}" for key, value in data.items())
    except Exception:
        return clean(out)


def bitlocker_status():
    out = run_command(["manage-bde.exe", "-status", "C:"], timeout=15)
    fields = {}
    for line in out.splitlines():
        if ":" in line:
            key, value = line.split(":", 1)
            key = key.strip()
            value = value.strip()
            if key in ("Conversion Status", "Percentage Encrypted", "Protection Status", "Lock Status"):
                fields[key] = value
    if fields:
        return "; ".join(f"{k}: {v}" for k, v in fields.items())
    return "Unavailable"


def antivirus_status():
    products = []
    client = get_wmi("root\\SecurityCenter2")
    if client is not None:
        try:
            for product in client.AntiVirusProduct():
                products.append(attr(product, "displayName"))
        except Exception:
            pass
    if products:
        return ", ".join(sorted(set(products)))
    return "Not detected"


def wifi_ssid():
    out = run_command(["netsh.exe", "wlan", "show", "interfaces"], timeout=10)
    for line in out.splitlines():
        text = line.strip()
        if text.startswith("SSID") and not text.startswith("BSSID") and ":" in text:
            value = text.split(":", 1)[1].strip()
            if value:
                return value
    profile = run_powershell(
        "try { Get-NetConnectionProfile | Where-Object { $_.InterfaceAlias -match 'Wi-Fi|Wireless|WLAN' -and $_.IPv4Connectivity -ne 'Disconnected' } | Select-Object -First 1 -ExpandProperty Name } catch { '' }",
        timeout=8,
    ).strip()
    if profile:
        return profile.splitlines()[-1].strip()
    return "Not connected or unavailable"


def battery_status():
    rows = wmi_all("Win32_Battery")
    if not rows:
        return "No battery detected"
    battery = rows[0]
    charge = attr(battery, "EstimatedChargeRemaining", "")
    status = attr(battery, "BatteryStatus", "")
    health = []
    full = wmi_first("BatteryFullChargedCapacity", namespace="root\\wmi")
    designed = wmi_first("BatteryStaticData", namespace="root\\wmi")
    try:
        full_capacity = float(getattr(full, "FullChargedCapacity"))
        designed_capacity = float(getattr(designed, "DesignedCapacity"))
        if designed_capacity > 0:
            health.append(f"Health {full_capacity / designed_capacity * 100:.0f}%")
    except Exception:
        pass
    if charge:
        health.append(f"Charge {charge}%")
    if status:
        health.append(f"Status {status}")
    return "; ".join(health) if health else "Detected"


def collect_about_info():
    return OrderedDict(
        [
            ("Program", APP_TITLE),
            ("Version", APP_VERSION),
            ("Purpose", "Collects computer asset, Windows, hardware, security, network, and health information."),
            ("Privacy Mode", "Hides serials, users, IPs, MAC addresses, profile paths, and similar identifiers in views and exports."),
            ("Export Formats", "Text, CSV, JSON, Excel, PDF, screenshot"),
            ("New View Areas", "Asset ownership, management status, access, startup, services, updates, certificates, peripherals, USB history, disk health, and recent errors"),
            ("Generated On", datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")),
        ]
    )


def device_asset_tag():
    enclosure = wmi_first("Win32_SystemEnclosure")
    tag = attr(enclosure, "SMBIOSAssetTag", "")
    return tag if tag and tag.lower() not in ("none", "no asset tag", "to be filled by o.e.m.", "default string") else "Not set"


def registered_owner():
    return registry_get(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", "Unavailable")


def assigned_user():
    computer = wmi_first("Win32_ComputerSystem")
    return attr(computer, "UserName", current_user())


def last_logged_in_user():
    base = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"
    return registry_get(winreg.HKEY_LOCAL_MACHINE, base, "LastLoggedOnUser", "Unavailable")


def device_age():
    bios = wmi_first("Win32_BIOS")
    release = parse_wmi_datetime(attr(bios, "ReleaseDate", ""))
    if release:
        return age_since(release)
    os_info = wmi_first("Win32_OperatingSystem")
    return age_since(attr(os_info, "InstallDate", ""))


def license_channel():
    products = wmi_query("SELECT Name, PartialProductKey, ProductKeyChannel FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL")
    for product in products:
        name = attr(product, "Name", "")
        if "windows" in name.lower():
            return attr(product, "ProductKeyChannel")
    return "Unavailable"


def dsreg_status():
    fields = {}
    out = run_command(["dsregcmd.exe", "/status"], timeout=15)
    for line in out.splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        key = key.strip()
        value = value.strip()
        if key:
            fields[key] = value
    return fields


def mdm_enrollment_status(dsreg=None):
    dsreg = dsreg or dsreg_status()
    mdm_url = clean(dsreg.get("MdmUrl"), "")
    if mdm_url and mdm_url != "Unavailable":
        return "Enrolled"
    enrollment_root = r"SOFTWARE\Microsoft\Enrollments"
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, enrollment_root) as key:
            count = winreg.QueryInfoKey(key)[0]
            return f"Enrollment keys found ({count})" if count else "Not detected"
    except Exception:
        return "Not detected"


def service_state(name):
    service = None
    rows = wmi_query(f"SELECT State FROM Win32_Service WHERE Name='{name}'")
    if rows:
        service = rows[0]
    if service is None:
        return "Not found"
    return attr(service, "State")


def sccm_status():
    state = service_state("CcmExec")
    if state != "Not found":
        return f"Client service {state}"
    if get_wmi("root\\ccm") is not None:
        return "Client namespace detected"
    return "Not detected"


def group_policy_last_update():
    out = run_command(["gpresult.exe", "/r", "/scope", "computer"], timeout=20)
    for line in out.splitlines():
        text = line.strip()
        if text.lower().startswith("last time group policy was applied"):
            return text.split(":", 1)[1].strip() if ":" in text else text
    return "Unavailable"


def windows_update_source():
    policy = r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    au = policy + r"\AU"
    use_wsus = registry_get(winreg.HKEY_LOCAL_MACHINE, au, "UseWUServer", "")
    server = registry_get(winreg.HKEY_LOCAL_MACHINE, policy, "WUServer", "")
    if str(use_wsus) == "1" or (server and server != "Unavailable"):
        return f"WSUS / managed update source{f': {server}' if server and server != 'Unavailable' else ''}"
    return "Default Windows Update / Microsoft Update"


def pending_update_summary():
    script = (
        "$ErrorActionPreference='Stop'; "
        "$s=New-Object -ComObject Microsoft.Update.Session; "
        "$r=$s.CreateUpdateSearcher().Search(\"IsInstalled=0 and Type='Software'\"); "
        "[pscustomobject]@{Count=$r.Updates.Count} | ConvertTo-Json -Compress"
    )
    data = powershell_json(script, timeout=25)
    if isinstance(data, dict) and "Count" in data:
        return f"{data.get('Count')} update(s) found"
    return "Unavailable"


def firewall_profile_status():
    out = run_command(["netsh.exe", "advfirewall", "show", "allprofiles", "state"], timeout=12)
    statuses = []
    current = ""
    for line in out.splitlines():
        text = line.strip()
        if text.lower().endswith("profile settings:"):
            current = text.replace("Profile Settings:", "").strip()
        elif text.lower().startswith("state") and current:
            statuses.append(f"{current}: {text.split()[-1]}")
    return "; ".join(statuses) if statuses else "Unavailable"


def defender_status():
    script = (
        "try { Get-MpComputerStatus | Select-Object "
        "AMServiceEnabled,AntivirusEnabled,RealTimeProtectionEnabled,AntispywareEnabled,IoavProtectionEnabled,NISEnabled "
        "| ConvertTo-Json -Compress } catch { '{}' }"
    )
    data = powershell_json(script, timeout=15)
    if not isinstance(data, dict) or not data:
        return "Unavailable"
    labels = {
        "AMServiceEnabled": "Service",
        "AntivirusEnabled": "Antivirus",
        "RealTimeProtectionEnabled": "Real-time",
        "AntispywareEnabled": "Antispyware",
        "IoavProtectionEnabled": "Downloaded files",
        "NISEnabled": "Network inspection",
    }
    return "; ".join(f"{labels.get(key, key)}={yes_no(value)}" for key, value in data.items())


def defender_definition_age():
    script = (
        "try { $d=(Get-MpComputerStatus).AntivirusSignatureLastUpdated; "
        "if ($d) { [pscustomobject]@{Updated=$d.ToString('yyyy-MM-dd HH:mm:ss'); DaysOld=(New-TimeSpan -Start $d -End (Get-Date)).Days} "
        "| ConvertTo-Json -Compress } } catch { '{}' }"
    )
    data = powershell_json(script, timeout=15)
    if isinstance(data, dict) and data.get("Updated"):
        return f"{data.get('Updated')} ({data.get('DaysOld')} days old)"
    return "Unavailable"


def endpoint_agent_status(apps):
    terms = [
        "defender for endpoint",
        "carbon black",
        "crowdstrike",
        "falcon",
        "sentinelone",
        "trellix endpoint",
        "mcafee endpoint",
        "tanium",
        "sophos endpoint",
        "symantec endpoint",
    ]
    detected = find_installed(apps, terms)
    sense = service_state("Sense")
    if sense != "Not found":
        return f"Microsoft Defender for Endpoint service {sense}"
    return detected.replace("Installed", "Detected")


def collect_asset_ownership_info(system_info, hardware_info, network_info):
    computer = wmi_first("Win32_ComputerSystem")
    return OrderedDict(
        [
            ("Asset Tag", device_asset_tag()),
            ("Assigned User", assigned_user()),
            ("Registered Owner", registered_owner()),
            ("Department", "Not set"),
            ("Device Location", "Not set"),
            ("Purchase Date", "Not set"),
            ("Warranty Expiration", "Not set"),
            ("Lease Refresh Date", "Not set"),
            ("Device Age", device_age()),
            ("Manufacturer", hardware_info.get("Manufacturer")),
            ("Model", hardware_info.get("Model")),
            ("Serial Number", hardware_info.get("Serial Number")),
            ("Computer Name", system_info.get("Computer Name")),
            ("Domain / Workgroup", attr(computer, "Domain")),
            ("Primary IP Address", network_info.get("Primary IP Address")),
        ]
    )


def collect_management_info(dsreg, pending_updates=None):
    return OrderedDict(
        [
            ("Azure AD Joined", clean(dsreg.get("AzureAdJoined"), "Unavailable")),
            ("Domain Joined", clean(dsreg.get("DomainJoined"), "Unavailable")),
            ("Device Auth Status", clean(dsreg.get("DeviceAuthStatus"), "Unavailable")),
            ("Tenant Name", clean(dsreg.get("TenantName"), "Unavailable")),
            ("Intune / MDM Enrollment", mdm_enrollment_status(dsreg)),
            ("SCCM / ConfigMgr Client", sccm_status()),
            ("Last Group Policy Update", group_policy_last_update()),
            ("Windows Update Source", windows_update_source()),
            ("Missing / Pending Updates", pending_updates or pending_update_summary()),
            ("Pending Reboot", pending_reboot_status()),
        ]
    )


def collect_access_info(local_admin_rows=None, user_profile_rows=None, startup_rows=None, running_service_rows=None, critical_service_rows=None, task_summary=None):
    local_admin_rows = collect_local_admin_rows() if local_admin_rows is None else local_admin_rows
    user_profile_rows = collect_user_profile_rows() if user_profile_rows is None else user_profile_rows
    startup_rows = collect_startup_rows() if startup_rows is None else startup_rows
    running_service_rows = collect_running_service_rows() if running_service_rows is None else running_service_rows
    critical_service_rows = collect_critical_service_rows() if critical_service_rows is None else critical_service_rows
    task_summary = scheduled_task_summary() if task_summary is None else task_summary
    return OrderedDict(
        [
            ("Assigned User", assigned_user()),
            ("Last Logged-In User", last_logged_in_user()),
            ("Current User", current_user()),
            ("Current User Is Admin", yes_no(is_admin())),
            ("Local Administrator Count", str(len(local_admin_rows))),
            ("User Profile Count", str(len(user_profile_rows))),
            ("Startup App Count", str(len(startup_rows))),
            ("Running Service Count", str(len(running_service_rows))),
            ("Critical Stopped Services", critical_stopped_summary(critical_service_rows)),
            ("Scheduled Tasks Summary", task_summary),
        ]
    )


def collect_update_info(system_info, pending_updates=None):
    return OrderedDict(
        [
            ("Windows License Channel", license_channel()),
            ("Windows Activation Status", system_info.get("Windows Activation")),
            ("Windows Edition", system_info.get("Windows Edition")),
            ("Windows Release", system_info.get("Windows Release")),
            ("Build Number", system_info.get("Build Number")),
            ("Last Windows Update", system_info.get("Last Windows Update")),
            ("Missing / Pending Updates", pending_updates or pending_update_summary()),
            ("Windows Update Source", windows_update_source()),
            ("Pending Reboot", system_info.get("Pending Reboot")),
        ]
    )


def collect_local_admin_rows():
    out = run_command(["net.exe", "localgroup", "administrators"], timeout=10)
    rows = []
    capture = False
    for line in out.splitlines():
        text = line.strip()
        if not text:
            continue
        if text.startswith("---"):
            capture = True
            continue
        if "command completed" in text.lower():
            break
        if capture:
            rows.append(OrderedDict([("Account", text)]))
    return rows


def collect_user_profile_rows():
    rows = []
    for profile in wmi_all("Win32_UserProfile"):
        local_path = attr(profile, "LocalPath", "")
        if not local_path or local_path == "Unavailable":
            continue
        if str(attr(profile, "Special", "")).lower() == "true":
            continue
        rows.append(
            OrderedDict(
                [
                    ("User", Path(local_path).name),
                    ("Profile Path", local_path),
                    ("Loaded", yes_no(str(attr(profile, "Loaded", "")).lower() == "true")),
                    ("Last Use", wmi_datetime(attr(profile, "LastUseTime", ""))),
                ]
            )
        )
    return rows


def collect_startup_rows():
    rows = []
    run_keys = [
        ("HKLM Run", winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"),
        ("HKLM Run 32-bit", winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"),
        ("HKCU Run", winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"),
    ]
    for source, root, path in run_keys:
        try:
            with winreg.OpenKey(root, path) as key:
                count = winreg.QueryInfoKey(key)[1]
                for index in range(count):
                    name, value, _ = winreg.EnumValue(key, index)
                    rows.append(OrderedDict([("Source", source), ("Name", clean(name)), ("Command", clean(value))]))
        except Exception:
            continue
    folders = [
        ("Current User Startup", Path(os.environ.get("APPDATA", "")) / r"Microsoft\Windows\Start Menu\Programs\Startup"),
        ("All Users Startup", Path(os.environ.get("PROGRAMDATA", "")) / r"Microsoft\Windows\Start Menu\Programs\Startup"),
    ]
    for source, folder in folders:
        try:
            if folder.exists():
                for item in folder.iterdir():
                    rows.append(OrderedDict([("Source", source), ("Name", item.name), ("Command", str(item))]))
        except Exception:
            continue
    return rows


def collect_running_service_rows():
    rows = []
    for service in wmi_query("SELECT Name, DisplayName, State, StartMode FROM Win32_Service WHERE State='Running'"):
        rows.append(
            OrderedDict(
                [
                    ("Name", attr(service, "Name")),
                    ("Display Name", attr(service, "DisplayName")),
                    ("State", attr(service, "State")),
                    ("Start Mode", attr(service, "StartMode")),
                ]
            )
        )
    return rows


def collect_critical_service_rows():
    services = OrderedDict(
        [
            ("WinDefend", "Microsoft Defender Antivirus"),
            ("Sense", "Microsoft Defender for Endpoint"),
            ("mpssvc", "Windows Firewall"),
            ("wuauserv", "Windows Update"),
            ("BITS", "Background Intelligent Transfer"),
            ("EventLog", "Windows Event Log"),
            ("Winmgmt", "Windows Management Instrumentation"),
            ("Dhcp", "DHCP Client"),
            ("Dnscache", "DNS Client"),
        ]
    )
    rows = []
    for name, display in services.items():
        rows.append(OrderedDict([("Service", name), ("Display Name", display), ("State", service_state(name))]))
    return rows


def critical_stopped_summary(rows=None):
    rows = collect_critical_service_rows() if rows is None else rows
    stopped = [row["Display Name"] for row in rows if row["State"].lower() not in ("running", "not found")]
    return ", ".join(stopped) if stopped else "None detected"


def scheduled_task_summary():
    script = "Get-ScheduledTask | Group-Object State | Select-Object Name,Count | ConvertTo-Json -Compress"
    data = powershell_json(script, timeout=20)
    rows = as_list(data)
    if not rows:
        return "Unavailable"
    return "; ".join(f"{clean(row.get('Name'))}: {clean(row.get('Count'))}" for row in rows if isinstance(row, dict))


def collect_hotfix_rows(limit=30):
    rows = []
    for fix in wmi_all("Win32_QuickFixEngineering"):
        installed = attr(fix, "InstalledOn", "")
        date = parse_date(installed) or datetime.min
        rows.append((date, OrderedDict([("Hotfix", attr(fix, "HotFixID")), ("Description", attr(fix, "Description")), ("Installed On", installed), ("Installed By", attr(fix, "InstalledBy"))])))
    rows.sort(key=lambda item: item[0], reverse=True)
    return [row for _date, row in rows[:limit]]


def collect_vpn_rows(apps):
    terms = ("vpn", "anyconnect", "globalprotect", "forticlient", "fortinet", "sonicwall", "pulse secure", "checkpoint", "tailscale", "wireguard", "openvpn", "zscaler")
    rows = []
    for app in apps:
        name = app.get("Name", "")
        if any(term in name.lower() for term in terms):
            rows.append(OrderedDict([("Name", name), ("Version", app.get("Version", "Unavailable")), ("Publisher", app.get("Publisher", "Unavailable"))]))
    return rows


def collect_certificate_rows():
    script = (
        "$limit=(Get-Date).AddDays(60); "
        "$paths=@('Cert:\\LocalMachine\\My','Cert:\\CurrentUser\\My'); "
        "$items=@(); foreach ($p in $paths) { if (Test-Path $p) { "
        "$items += Get-ChildItem $p | Where-Object {$_.NotAfter -lt $limit} | "
        "Select-Object @{n='Store';e={$p}}, Subject, @{n='Expires';e={$_.NotAfter.ToString('yyyy-MM-dd')}}, Thumbprint } } "
        "$items | Sort-Object Expires | Select-Object -First 40 | ConvertTo-Json -Compress"
    )
    data = powershell_json(script, timeout=20)
    rows = []
    for item in as_list(data):
        if isinstance(item, dict):
            rows.append(OrderedDict([("Store", clean(item.get("Store"))), ("Subject", clean(item.get("Subject"))), ("Expires", clean(item.get("Expires"))), ("Thumbprint", clean(item.get("Thumbprint")))]))
    return rows


def collect_mapped_drive_rows():
    rows = []
    for drive in wmi_all("Win32_MappedLogicalDisk"):
        rows.append(
            OrderedDict(
                [
                    ("Drive", attr(drive, "DeviceID")),
                    ("Name", attr(drive, "Name")),
                    ("Network Path", attr(drive, "ProviderName")),
                    ("File System", attr(drive, "FileSystem")),
                ]
            )
        )
    return rows


def collect_printer_rows():
    rows = []
    for printer in wmi_all("Win32_Printer"):
        rows.append(
            OrderedDict(
                [
                    ("Name", attr(printer, "Name")),
                    ("Default", yes_no(str(attr(printer, "Default", "")).lower() == "true")),
                    ("Network", yes_no(str(attr(printer, "Network", "")).lower() == "true")),
                    ("Port", attr(printer, "PortName")),
                    ("Driver", attr(printer, "DriverName")),
                ]
            )
        )
    return rows


def collect_share_rows():
    rows = []
    for share in wmi_all("Win32_Share"):
        rows.append(OrderedDict([("Name", attr(share, "Name")), ("Share Path", attr(share, "Path")), ("Type", attr(share, "Type")), ("Description", attr(share, "Description"))]))
    return rows


def collect_usb_rows():
    rows = []
    root = r"SYSTEM\CurrentControlSet\Enum\USBSTOR"
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, root) as key:
            device_count = winreg.QueryInfoKey(key)[0]
            for device_index in range(device_count):
                device_name = winreg.EnumKey(key, device_index)
                with winreg.OpenKey(key, device_name) as device_key:
                    instance_count = winreg.QueryInfoKey(device_key)[0]
                    for instance_index in range(instance_count):
                        instance = winreg.EnumKey(device_key, instance_index)
                        with winreg.OpenKey(device_key, instance) as instance_key:
                            rows.append(
                                OrderedDict(
                                    [
                                        ("Device", registry_value(instance_key, "FriendlyName") or device_name),
                                        ("Manufacturer", registry_value(instance_key, "Mfg") or "Unavailable"),
                                        ("Serial Number", instance),
                                        ("Status", registry_value(instance_key, "DeviceDesc") or "Unavailable"),
                                    ]
                                )
                            )
    except Exception:
        pass
    return rows[:60]


def collect_battery_rows():
    rows = []
    for battery in wmi_all("Win32_Battery"):
        rows.append(
            OrderedDict(
                [
                    ("Name", attr(battery, "Name")),
                    ("Charge", clean(attr(battery, "EstimatedChargeRemaining", ""), "") + "%"),
                    ("Status", attr(battery, "Status")),
                    ("Chemistry", attr(battery, "Chemistry")),
                    ("Design Voltage", attr(battery, "DesignVoltage")),
                ]
            )
        )
    full = wmi_first("BatteryFullChargedCapacity", namespace="root\\wmi")
    designed = wmi_first("BatteryStaticData", namespace="root\\wmi")
    if full is not None or designed is not None:
        rows.append(
            OrderedDict(
                [
                    ("Name", "Battery Capacity"),
                    ("Charge", attr(full, "FullChargedCapacity")),
                    ("Status", "Full charged capacity"),
                    ("Chemistry", "Design capacity"),
                    ("Design Voltage", attr(designed, "DesignedCapacity")),
                ]
            )
        )
    return rows


def collect_disk_smart_rows():
    rows = []
    for item in wmi_all("MSStorageDriver_FailurePredictStatus", namespace="root\\wmi"):
        rows.append(
            OrderedDict(
                [
                    ("Instance", attr(item, "InstanceName")),
                    ("Predict Failure", yes_no(str(attr(item, "PredictFailure", "")).lower() == "true")),
                    ("Reason", attr(item, "Reason")),
                ]
            )
        )
    if rows:
        return rows
    script = "Get-PhysicalDisk | Select-Object FriendlyName, HealthStatus, OperationalStatus, MediaType, Size | ConvertTo-Json -Compress"
    data = powershell_json(script, timeout=15)
    for item in as_list(data):
        if isinstance(item, dict):
            rows.append(
                OrderedDict(
                    [
                        ("Instance", clean(item.get("FriendlyName"))),
                        ("Predict Failure", clean(item.get("HealthStatus"))),
                        ("Reason", clean(item.get("OperationalStatus"))),
                    ]
                )
            )
    return rows


def collect_drive_temperature_rows():
    script = (
        "Get-PhysicalDisk | ForEach-Object { $disk=$_; "
        "try { $rel=$disk | Get-StorageReliabilityCounter; "
        "[pscustomobject]@{Name=$disk.FriendlyName; DeviceId=$disk.DeviceId; Temperature=$rel.Temperature; ReadErrors=$rel.ReadErrorsTotal; WriteErrors=$rel.WriteErrorsTotal} } catch {} } "
        "| ConvertTo-Json -Compress"
    )
    data = powershell_json(script, timeout=20)
    rows = []
    for item in as_list(data):
        if isinstance(item, dict):
            temp = clean(item.get("Temperature"))
            rows.append(
                OrderedDict(
                    [
                        ("Drive", clean(item.get("Name"))),
                        ("Device ID", clean(item.get("DeviceId"))),
                        ("Temperature", f"{temp} C" if temp not in ("Unavailable", "") else "Unavailable"),
                        ("Read Errors", clean(item.get("ReadErrors"))),
                        ("Write Errors", clean(item.get("WriteErrors"))),
                    ]
                )
            )
    return rows


def collect_recent_error_rows():
    script = (
        "$start=(Get-Date).AddDays(-7); "
        "Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2; StartTime=$start} -MaxEvents 40 | "
        "Select-Object @{n='Time';e={$_.TimeCreated.ToString('yyyy-MM-dd HH:mm')}}, Id, LevelDisplayName, ProviderName, "
        "@{n='Message';e={($_.Message -replace \"`r|`n\", ' ').Substring(0, [Math]::Min(220, ($_.Message -replace \"`r|`n\", ' ').Length))}} "
        "| ConvertTo-Json -Compress"
    )
    data = powershell_json(script, timeout=25)
    rows = []
    for item in as_list(data):
        if isinstance(item, dict):
            rows.append(
                OrderedDict(
                    [
                        ("Time", clean(item.get("Time"))),
                        ("Event ID", clean(item.get("Id"))),
                        ("Level", clean(item.get("LevelDisplayName"))),
                        ("Source", clean(item.get("ProviderName"))),
                        ("Message", clean(item.get("Message"))),
                    ]
                )
            )
    return rows


def collect_peripheral_info(mapped=None, printers=None, shares=None, usb=None, monitors=None):
    mapped = collect_mapped_drive_rows() if mapped is None else mapped
    printers = collect_printer_rows() if printers is None else printers
    shares = collect_share_rows() if shares is None else shares
    usb = collect_usb_rows() if usb is None else usb
    monitors = collect_monitor_rows() if monitors is None else monitors
    return OrderedDict(
        [
            ("Mapped Network Drives", str(len(mapped))),
            ("Installed Printers", str(len(printers))),
            ("Shared Folders", str(len(shares))),
            ("USB Storage History", str(len(usb))),
            ("Monitor Count", str(len(monitors))),
            ("Monitor Serial Numbers", ", ".join(row.get("Serial Number", "Unavailable") for row in monitors) if monitors else "Unavailable"),
        ]
    )


def collect_health_info(disk_rows=None, temp_rows=None, event_rows=None, battery_rows=None):
    disk_rows = collect_disk_smart_rows() if disk_rows is None else disk_rows
    temp_rows = collect_drive_temperature_rows() if temp_rows is None else temp_rows
    event_rows = collect_recent_error_rows() if event_rows is None else event_rows
    battery_rows = collect_battery_rows() if battery_rows is None else battery_rows
    failures = [row.get("Instance") for row in disk_rows if str(row.get("Predict Failure", "")).lower() in ("yes", "true")]
    return OrderedDict(
        [
            ("Battery Health Report", battery_status()),
            ("Battery Detail Records", str(len(battery_rows))),
            ("Disk SMART Health", "Failure predicted: " + ", ".join(failures) if failures else ("Records found" if disk_rows else "Unavailable")),
            ("Drive Temperature Records", str(len(temp_rows))),
            ("Recent Critical/Error Events", str(len(event_rows))),
            ("Recent System Error Window", "Last 7 days"),
        ]
    )


def collect_system_info(fast=False):
    os_info = wmi_first("Win32_OperatingSystem")
    release = windows_release_info()
    uptime_seconds = int(time.time() - psutil.boot_time()) if psutil else 0
    return OrderedDict(
        [
            ("Computer Name", platform.node()),
            ("Operating System", attr(os_info, "Caption", f"{platform.system()} {platform.release()}")),
            ("Windows Edition", release["Edition"]),
            ("Windows Release", release["Release"]),
            ("OS Version", attr(os_info, "Version", platform.version())),
            ("Build Number", release["Build"]),
            ("Install Date", wmi_datetime(attr(os_info, "InstallDate", ""))),
            ("Last Boot Time", wmi_datetime(attr(os_info, "LastBootUpTime", ""))),
            ("System Uptime", format_uptime(uptime_seconds)),
            ("Architecture", platform.architecture()[0]),
            ("Language", attr(os_info, "OSLanguage", "en-US")),
            ("Last Windows Update", "Run full scan to check" if fast else last_hotfix()),
            ("Pending Reboot", pending_reboot_status()),
            ("Windows Activation", "Run full scan to check" if fast else activation_status()),
            ("Windows License Channel", "Run full scan to check" if fast else license_channel()),
        ]
    )


def collect_hardware_info(fast=False):
    computer = wmi_first("Win32_ComputerSystem")
    bios = wmi_first("Win32_BIOS")
    enclosure = wmi_first("Win32_SystemEnclosure")
    board = wmi_first("Win32_BaseBoard")
    cpu = wmi_first("Win32_Processor")
    memory = psutil.virtual_memory() if psutil else None
    cpu_freq = psutil.cpu_freq() if psutil else None
    chassis = chassis_type_label(getattr(enclosure, "ChassisTypes", None))
    return OrderedDict(
        [
            ("Manufacturer", attr(computer, "Manufacturer")),
            ("Model", attr(computer, "Model")),
            ("Case Type", chassis),
            ("Asset Tag", device_asset_tag()),
            ("Device Age", device_age()),
            ("Serial Number", attr(bios, "SerialNumber")),
            ("BIOS Revision", attr(bios, "SMBIOSBIOSVersion")),
            ("BIOS Release Date", wmi_datetime(attr(bios, "ReleaseDate", ""))),
            ("Motherboard Manufacturer", attr(board, "Manufacturer")),
            ("Motherboard Model", attr(board, "Product")),
            ("CPU Name", attr(cpu, "Name", platform.processor())),
            ("CPU Speed", f"{cpu_freq.current:.2f} MHz" if cpu_freq else attr(cpu, "MaxClockSpeed", "Unavailable")),
            ("CPU Cores", attr(cpu, "NumberOfCores")),
            ("CPU Threads", attr(cpu, "NumberOfLogicalProcessors")),
            ("RAM Installed", format_gb(attr(computer, "TotalPhysicalMemory", ""))),
            ("RAM Available", format_gb(memory.available) if memory else "Unavailable"),
            ("RAM Slot Count", attr(wmi_first("Win32_PhysicalMemoryArray"), "MemoryDevices")),
            ("Battery Health", "Run full scan to check" if fast else battery_status()),
            ("TPM Status", "Run full scan to check" if fast else tpm_status()),
            ("Secure Boot", "Run full scan to check" if fast else secure_boot_status()),
        ]
    )


def collect_ram_slots():
    rows = []
    for mem in wmi_all("Win32_PhysicalMemory"):
        rows.append(
            OrderedDict(
                [
                    ("Slot", attr(mem, "DeviceLocator")),
                    ("Bank", attr(mem, "BankLabel")),
                    ("Capacity", format_gb(attr(mem, "Capacity", ""))),
                    ("Speed", clean(attr(mem, "Speed", ""), "") + (" MHz" if attr(mem, "Speed", "") != "Unavailable" else "")),
                    ("Manufacturer", attr(mem, "Manufacturer")),
                    ("Part Number", attr(mem, "PartNumber")),
                    ("Serial Number", attr(mem, "SerialNumber")),
                ]
            )
        )
    return rows


def collect_gpu_rows():
    rows = []
    for gpu in wmi_all("Win32_VideoController"):
        ram = attr(gpu, "AdapterRAM", "")
        rows.append(
            OrderedDict(
                [
                    ("Name", attr(gpu, "Name")),
                    ("Adapter RAM", format_gb(ram)),
                    ("Resolution", f"{attr(gpu, 'CurrentHorizontalResolution', '')} x {attr(gpu, 'CurrentVerticalResolution', '')}"),
                    ("Driver Version", attr(gpu, "DriverVersion")),
                ]
            )
        )
    return rows


def collect_monitor_rows():
    rows = []
    monitor_ids = wmi_all("WmiMonitorID", namespace="root\\wmi")
    for monitor in monitor_ids:
        manufacturer = decode_wmi_text(getattr(monitor, "ManufacturerName", []))
        name = decode_wmi_text(getattr(monitor, "UserFriendlyName", []))
        serial = decode_wmi_text(getattr(monitor, "SerialNumberID", []))
        rows.append(OrderedDict([("Manufacturer", manufacturer or "Unavailable"), ("Model", name or "Unavailable"), ("Serial Number", serial or "Unavailable")]))
    if not rows:
        for monitor in wmi_all("Win32_DesktopMonitor"):
            rows.append(OrderedDict([("Manufacturer", attr(monitor, "MonitorManufacturer")), ("Model", attr(monitor, "Name")), ("Serial Number", "Unavailable")]))
    return rows


def collect_storage_info():
    c_usage = None
    if psutil:
        try:
            c_usage = psutil.disk_usage("C:\\")
        except Exception:
            c_usage = None
    return OrderedDict(
        [
            ("C: Total Space", format_gb(c_usage.total) if c_usage else "Unavailable"),
            ("C: Used Space", format_gb(c_usage.used) if c_usage else "Unavailable"),
            ("C: Free Space", format_gb(c_usage.free) if c_usage else "Unavailable"),
            ("C: Percent Used", f"{c_usage.percent:.1f}%" if c_usage else "Unavailable"),
        ]
    )


def collect_logical_drive_rows():
    rows = []
    for disk in wmi_all("Win32_LogicalDisk"):
        device = attr(disk, "DeviceID")
        rows.append(
            OrderedDict(
                [
                    ("Drive", device),
                    ("Name", attr(disk, "VolumeName")),
                    ("File System", attr(disk, "FileSystem")),
                    ("Total", format_gb(attr(disk, "Size", ""))),
                    ("Free", format_gb(attr(disk, "FreeSpace", ""))),
                    ("Type", attr(disk, "DriveType")),
                ]
            )
        )
    return rows


def collect_physical_disk_rows():
    rows = []
    for disk in wmi_all("Win32_DiskDrive"):
        rows.append(
            OrderedDict(
                [
                    ("Model", attr(disk, "Model")),
                    ("Serial Number", attr(disk, "SerialNumber")),
                    ("Size", format_gb(attr(disk, "Size", ""))),
                    ("Interface", attr(disk, "InterfaceType")),
                    ("Media Type", attr(disk, "MediaType")),
                    ("Status", attr(disk, "Status")),
                ]
            )
        )
    return rows


def collect_network_info():
    hostname = socket.gethostname()
    computer = wmi_first("Win32_ComputerSystem")
    return OrderedDict(
        [
            ("Hostname", hostname),
            ("Primary IP Address", primary_ip(hostname)),
            ("Wi-Fi SSID", wifi_ssid()),
            ("Machine Domain", attr(computer, "Domain")),
            ("Part of Domain", yes_no(str(attr(computer, "PartOfDomain", "")).lower() == "true")),
            ("Domain Controller", hostname),
            ("User ID", current_user()),
            ("User Domain", os.environ.get("USERDOMAIN", hostname)),
            ("Admin User", yes_no(is_admin())),
        ]
    )


def primary_ip(hostname):
    try:
        return socket.gethostbyname(hostname)
    except Exception:
        return "Unavailable"


def collect_adapter_rows():
    rows = []
    if psutil is None:
        return rows
    stats = psutil.net_if_stats()
    addrs = psutil.net_if_addrs()
    link_family = getattr(psutil, "AF_LINK", None)
    for name, stat in stats.items():
        if not stat.isup:
            continue
        ipv4 = []
        ipv6 = []
        mac = []
        for item in addrs.get(name, []):
            if item.family == socket.AF_INET:
                ipv4.append(item.address)
            elif item.family == socket.AF_INET6:
                ipv6.append(item.address)
            elif link_family is not None and item.family == link_family:
                mac.append(item.address)
        rows.append(
            OrderedDict(
                [
                    ("Adapter", name),
                    ("Status", "Up"),
                    ("Speed", f"{stat.speed} Mbps" if stat.speed else "Unavailable"),
                    ("IPv4", ", ".join(ipv4) or "Unavailable"),
                    ("IPv6", ", ".join(ipv6) or "Unavailable"),
                    ("MAC Address", ", ".join(mac) or "Unavailable"),
                ]
            )
        )
    return rows


def collect_application_info(apps):
    vpn_rows = collect_vpn_rows(apps)
    return OrderedDict(
        [
            ("Google Chrome", find_installed(apps, ["google chrome"])),
            ("Microsoft Edge", find_installed(apps, ["microsoft edge"])),
            ("Mozilla Firefox", find_installed(apps, ["mozilla firefox"])),
            ("Brave Browser", find_installed(apps, ["brave"])),
            ("Opera Browser", find_installed(apps, ["opera"])),
            ("Internet Explorer", registry_get(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Internet Explorer", "svcVersion", registry_get(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Internet Explorer", "Version", "Not installed"))),
            ("Adobe Flash", find_installed(apps, ["adobe flash", "flash player"])),
            ("VPN Clients Installed", str(len(vpn_rows))),
            ("Installed Applications Scanned", str(len(apps))),
        ]
    )


def collect_security_info(apps, fast=False):
    def detected(label, terms):
        return find_installed(apps, terms).replace("Installed", "Detected")

    if fast:
        security_agent = endpoint_agent_status(apps)
        return OrderedDict(
            [
                ("Virus Management", "Run full scan to check"),
                ("Virus Protection", "Run full scan to check"),
                ("Firewall Profile Status", "Run full scan to check"),
                ("Microsoft Defender Status", "Run full scan to check"),
                ("Defender Definition Age", "Run full scan to check"),
                ("EDR / Security Agent Status", security_agent),
                ("BitLocker Status", "Run full scan to check"),
                ("TPM Status", "Run full scan to check"),
                ("Secure Boot", "Run full scan to check"),
                ("DLP Status", detected("DLP", ["dlp", "digital guardian", "forcepoint", "data loss prevention"])),
                ("FRP Status", detected("FRP", ["file and removable", "frp", "endpoint encryption"])),
                ("ATP Status", detected("ATP", ["defender for endpoint", "advanced threat", "trellix endpoint", "carbon black"])),
                ("Web Control", detected("Web Control", ["web control", "zscaler", "umbrella", "web security"])),
                ("Browsium Status", detected("Browsium", ["browsium"])),
                ("SMP Agent Status", detected("SMP", ["symantec management agent", "smp agent"])),
                ("Encryption Status", "Run full scan to check"),
                ("Windows Activation", "Run full scan to check"),
                ("Pending Reboot", pending_reboot_status()),
            ]
        )

    return OrderedDict(
        [
            ("Virus Management", antivirus_status()),
            ("Virus Protection", antivirus_status()),
            ("Firewall Profile Status", firewall_profile_status()),
            ("Microsoft Defender Status", defender_status()),
            ("Defender Definition Age", defender_definition_age()),
            ("EDR / Security Agent Status", endpoint_agent_status(apps)),
            ("BitLocker Status", bitlocker_status()),
            ("TPM Status", tpm_status()),
            ("Secure Boot", secure_boot_status()),
            ("DLP Status", detected("DLP", ["dlp", "digital guardian", "forcepoint", "data loss prevention"])),
            ("FRP Status", detected("FRP", ["file and removable", "frp", "endpoint encryption"])),
            ("ATP Status", detected("ATP", ["defender for endpoint", "advanced threat", "trellix endpoint", "carbon black"])),
            ("Web Control", detected("Web Control", ["web control", "zscaler", "umbrella", "web security"])),
            ("Browsium Status", detected("Browsium", ["browsium"])),
            ("SMP Agent Status", detected("SMP", ["symantec management agent", "smp agent"])),
            ("Encryption Status", bitlocker_status()),
            ("Windows Activation", activation_status()),
            ("Pending Reboot", pending_reboot_status()),
        ]
    )


def collect_report(deep=True):
    apps = scan_uninstall_apps()
    report = OrderedDict()
    system_info = collect_system_info(fast=not deep)
    hardware_info = collect_hardware_info(fast=not deep)
    network_info = collect_network_info()
    storage_info = collect_storage_info()
    dsreg = dsreg_status() if deep else {}
    pending_updates = pending_update_summary() if deep else "Run full scan to check"
    hotfix_rows = collect_hotfix_rows() if deep else []
    local_admin_rows = collect_local_admin_rows() if deep else []
    user_profile_rows = collect_user_profile_rows() if deep else []
    startup_rows = collect_startup_rows() if deep else []
    running_service_rows = collect_running_service_rows() if deep else []
    critical_service_rows = collect_critical_service_rows() if deep else []
    task_summary = scheduled_task_summary() if deep else "Run full scan to check"
    vpn_rows = collect_vpn_rows(apps)
    certificate_rows = collect_certificate_rows() if deep else []
    mapped_rows = collect_mapped_drive_rows() if deep else []
    printer_rows = collect_printer_rows() if deep else []
    share_rows = collect_share_rows() if deep else []
    usb_rows = collect_usb_rows() if deep else []
    monitor_rows = collect_monitor_rows() if deep else []
    battery_rows = collect_battery_rows() if deep else []
    disk_smart_rows = collect_disk_smart_rows() if deep else []
    drive_temp_rows = collect_drive_temperature_rows() if deep else []
    recent_error_rows = collect_recent_error_rows() if deep else []
    report["About"] = {"items": collect_about_info(), "tables": OrderedDict()}
    report["Overview"] = {
        "items": OrderedDict(
            [
                ("Report Time", datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")),
                ("Computer Name", system_info.get("Computer Name")),
                ("Asset Tag", hardware_info.get("Asset Tag")),
                ("Assigned User", assigned_user()),
                ("Manufacturer", hardware_info.get("Manufacturer")),
                ("Model", hardware_info.get("Model")),
                ("Serial Number", hardware_info.get("Serial Number")),
                ("Operating System", system_info.get("Operating System")),
                ("Windows Release", system_info.get("Windows Release")),
                ("RAM Installed", hardware_info.get("RAM Installed")),
                ("C: Free Space", storage_info.get("C: Free Space")),
                ("Primary IP Address", network_info.get("Primary IP Address")),
                ("Security Agent", endpoint_agent_status(apps)),
                ("Pending Reboot", system_info.get("Pending Reboot")),
                ("Scan Mode", "Full scan" if deep else "Fast startup scan"),
            ]
        ),
        "tables": OrderedDict(),
    }
    report["Asset & Ownership"] = {"items": collect_asset_ownership_info(system_info, hardware_info, network_info), "tables": OrderedDict()}
    report["Operating System"] = {"items": system_info, "tables": OrderedDict()}
    report["Windows & Updates"] = {
        "items": collect_update_info(system_info, pending_updates),
        "tables": OrderedDict([("Installed Windows Updates Timeline", hotfix_rows)]),
    }
    report["Management"] = {"items": collect_management_info(dsreg, pending_updates), "tables": OrderedDict()}
    report["Access & Startup"] = {
        "items": collect_access_info(local_admin_rows, user_profile_rows, startup_rows, running_service_rows, critical_service_rows, task_summary),
        "tables": OrderedDict(
            [
                ("Local Administrator Accounts", local_admin_rows),
                ("User Profiles", user_profile_rows),
                ("Startup Apps", startup_rows),
                ("Running Services", running_service_rows),
                ("Critical Services", critical_service_rows),
            ]
        ),
    }
    report["Hardware"] = {
        "items": hardware_info,
        "tables": OrderedDict(
            [
                ("RAM Slots", collect_ram_slots() if deep else []),
                ("Graphics", collect_gpu_rows() if deep else []),
                ("Monitors", monitor_rows),
            ]
        ),
    }
    report["Storage"] = {
        "items": storage_info,
        "tables": OrderedDict(
            [
                ("Logical Drives", collect_logical_drive_rows() if deep else []),
                ("Physical Disks", collect_physical_disk_rows() if deep else []),
            ]
        ),
    }
    report["Network"] = {"items": network_info, "tables": OrderedDict([("Active Network Adapters", collect_adapter_rows() if deep else [])])}
    report["Applications"] = {"items": collect_application_info(apps), "tables": OrderedDict([("VPN Clients Installed", vpn_rows)])}
    report["Security"] = {
        "items": collect_security_info(apps, fast=not deep),
        "tables": OrderedDict([("Certificates Expiring Soon", certificate_rows), ("VPN / Remote Access Clients", vpn_rows)]),
    }
    report["Peripherals & Shares"] = {
        "items": collect_peripheral_info(mapped_rows, printer_rows, share_rows, usb_rows, monitor_rows),
        "tables": OrderedDict(
            [
                ("Mapped Network Drives", mapped_rows),
                ("Installed Printers", printer_rows),
                ("Shared Folders", share_rows),
                ("USB Device History", usb_rows),
                ("Monitor Serial Numbers", monitor_rows),
            ]
        ),
    }
    report["Activity & Health"] = {
        "items": collect_health_info(disk_smart_rows, drive_temp_rows, recent_error_rows, battery_rows),
        "tables": OrderedDict(
            [
                ("Battery Health Report", battery_rows),
                ("Disk SMART Health", disk_smart_rows),
                ("Drive Temperatures", drive_temp_rows),
                ("Recent System Errors and Crashes", recent_error_rows),
            ]
        ),
    }
    return report


def is_sensitive(field):
    text = str(field or "").lower()
    return any(term in text for term in SENSITIVE_TERMS)


def mask_value(field, value, privacy=False):
    if privacy and is_sensitive(field):
        return "Hidden in privacy mode"
    return clean(value)


def trim_cell(value, width):
    value = clean(value)
    return value if len(value) <= width else value[: max(width - 3, 1)] + "..."


def aligned_detail_lines(category, information, details):
    category_width = 22
    information_width = 34
    detail_width = 96
    category = trim_cell(category, category_width)
    information = trim_cell(information, information_width)
    wrapped = textwrap.wrap(clean(details), width=detail_width, replace_whitespace=False) or [""]
    lines = []
    for index, part in enumerate(wrapped):
        left = category if index == 0 else ""
        middle = information if index == 0 else ""
        lines.append(f"{left:<{category_width}}  {middle:<{information_width}}  {part}")
    return lines


def report_lines(report, privacy=False):
    category_width = 22
    information_width = 34
    lines = [
        f"{APP_TITLE} {APP_VERSION}",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')}",
        "",
        f"{'Category':<{category_width}}  {'Information':<{information_width}}  Details",
        "-" * 118,
    ]
    for section, bundle in report.items():
        for key, value in bundle.get("items", {}).items():
            lines.extend(aligned_detail_lines(section, key, mask_value(key, value, privacy)))
        for table_name, rows in bundle.get("tables", {}).items():
            if not rows:
                lines.extend(aligned_detail_lines(section, table_name, "No records found."))
                continue
            for index, row in enumerate(rows, 1):
                for key, value in row.items():
                    lines.extend(aligned_detail_lines(section, f"{table_name} #{index}: {key}", mask_value(key, value, privacy)))
        lines.append("")
    return lines


def flatten_csv_rows(report, privacy=False):
    rows = []
    for section, bundle in report.items():
        for key, value in bundle.get("items", {}).items():
            rows.append({"Section": section, "Table": "", "Row": "", "Field": key, "Value": mask_value(key, value, privacy)})
        for table_name, table_rows in bundle.get("tables", {}).items():
            for index, row in enumerate(table_rows, 1):
                for key, value in row.items():
                    rows.append({"Section": section, "Table": table_name, "Row": index, "Field": key, "Value": mask_value(key, value, privacy)})
    return rows


def masked_report(report, privacy=False):
    if not privacy:
        return report
    output = OrderedDict()
    for section, bundle in report.items():
        items = OrderedDict((key, mask_value(key, value, True)) for key, value in bundle.get("items", {}).items())
        tables = OrderedDict()
        for table_name, rows in bundle.get("tables", {}).items():
            tables[table_name] = [OrderedDict((key, mask_value(key, value, True)) for key, value in row.items()) for row in rows]
        output[section] = {"items": items, "tables": tables}
    return output


def write_text(path, report, privacy=False):
    Path(path).write_text("\n".join(report_lines(report, privacy)), encoding="utf-8")


def write_json(path, report, privacy=False):
    Path(path).write_text(json.dumps(masked_report(report, privacy), indent=2), encoding="utf-8")


def write_csv(path, report, privacy=False):
    rows = flatten_csv_rows(report, privacy)
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=["Section", "Table", "Row", "Field", "Value"])
        writer.writeheader()
        writer.writerows(rows)


def safe_sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\]", "_", name)[:31] or "Sheet"
    value = base
    counter = 2
    while value in used:
        suffix = f"_{counter}"
        value = base[: 31 - len(suffix)] + suffix
        counter += 1
    used.add(value)
    return value


def write_xlsx(path, report, privacy=False):
    if Workbook is None:
        raise RuntimeError("Excel export requires openpyxl.")
    try:
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
    except Exception:
        Font = None
        PatternFill = None
        get_column_letter = None
    wb = Workbook()
    used = set()
    first = True
    tab_colors = ["2f80ed", "12b76a", "f79009", "7a5af8", "ef4444", "06aed4", "667085"]
    for section, bundle in masked_report(report, privacy).items():
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = safe_sheet_name(section, used)
        ws.sheet_properties.tabColor = tab_colors[(len(used) - 1) % len(tab_colors)]
        row_index = 1
        ws.cell(row=row_index, column=1, value=section)
        if Font is not None:
            ws.cell(row=row_index, column=1).font = Font(bold=True, size=14)
        row_index += 2
        ws.cell(row=row_index, column=1, value="Information")
        ws.cell(row=row_index, column=2, value="Details")
        if Font is not None and PatternFill is not None:
            for col in range(1, 3):
                ws.cell(row=row_index, column=col).font = Font(bold=True, color="FFFFFF")
                ws.cell(row=row_index, column=col).fill = PatternFill("solid", fgColor="172033")
        ws.freeze_panes = "A4"
        row_index += 1
        for key, value in bundle.get("items", {}).items():
            ws.cell(row=row_index, column=1, value=key)
            ws.cell(row=row_index, column=2, value=value)
            row_index += 1
        for table_name, rows in bundle.get("tables", {}).items():
            row_index += 2
            ws.cell(row=row_index, column=1, value=table_name)
            if Font is not None:
                ws.cell(row=row_index, column=1).font = Font(bold=True)
            row_index += 1
            if rows:
                headers = list(rows[0].keys())
                for col, header in enumerate(headers, 1):
                    ws.cell(row=row_index, column=col, value=header)
                    if Font is not None and PatternFill is not None:
                        ws.cell(row=row_index, column=col).font = Font(bold=True, color="FFFFFF")
                        ws.cell(row=row_index, column=col).fill = PatternFill("solid", fgColor="344054")
                row_index += 1
                for data_row in rows:
                    for col, header in enumerate(headers, 1):
                        ws.cell(row=row_index, column=col, value=data_row.get(header))
                    row_index += 1
        for column in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(max(max_len + 2, 12), 55)
        if get_column_letter is not None:
            ws.auto_filter.ref = f"A3:{get_column_letter(max(ws.max_column, 2))}{max(ws.max_row, 3)}"
    wb.save(path)


def pdf_escape(text):
    return str(text).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)").replace("\r", "").replace("\n", " ")


def write_pdf(path, report, privacy=False):
    overview = report.get("Overview", {}).get("items", {})
    wrapped = [
        APP_TITLE,
        f"Version {APP_VERSION}",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')}",
        "",
        f"Computer: {mask_value('Computer Name', overview.get('Computer Name'), privacy)}",
        f"Model: {mask_value('Model', overview.get('Model'), privacy)}",
        f"Operating System: {mask_value('Operating System', overview.get('Operating System'), privacy)}",
        f"Pending Reboot: {mask_value('Pending Reboot', overview.get('Pending Reboot'), privacy)}",
        "",
        "Report Details",
        "-" * 72,
    ]
    for line in report_lines(report, privacy):
        if not line:
            wrapped.append("")
        else:
            wrapped.extend(textwrap.wrap(line, width=105, replace_whitespace=False) or [""])
    pages = [wrapped[i : i + 50] for i in range(0, len(wrapped), 50)] or [[]]
    objects = [None, None, b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>"]
    page_ids = []

    def add(obj):
        objects.append(obj)
        return len(objects)

    for page_lines in pages:
        stream = "BT\n/F1 8 Tf\n13 TL\n44 750 Td\n"
        for line in page_lines:
            stream += f"({pdf_escape(line)}) Tj\nT*\n"
        stream += "ET\n"
        stream_bytes = stream.encode("cp1252", "replace")
        content_id = add(b"<< /Length " + str(len(stream_bytes)).encode("ascii") + b" >>\nstream\n" + stream_bytes + b"endstream")
        page_id = add(
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents {content_id} 0 R >>".encode("ascii")
        )
        page_ids.append(page_id)
    objects[0] = b"<< /Type /Catalog /Pages 2 0 R >>"
    kids = " ".join(f"{page_id} 0 R" for page_id in page_ids)
    objects[1] = f"<< /Type /Pages /Kids [{kids}] /Count {len(page_ids)} >>".encode("ascii")
    output = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for index, obj in enumerate(objects, 1):
        offsets.append(len(output))
        output.extend(f"{index} 0 obj\n".encode("ascii"))
        output.extend(obj)
        output.extend(b"\nendobj\n")
    xref_offset = len(output)
    output.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    output.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        output.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    output.extend(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF".encode("ascii"))
    Path(path).write_bytes(output)


class ScrollFrame(ttk.Frame):
    def __init__(self, parent, bg=None):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=bg or "#f7f9fc")
        self.inner = ttk.Frame(self.canvas)
        self.window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

    def _on_inner_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.window, width=event.width)

    def _bind_mousewheel(self, _event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, _event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"


class ToolTip:
    def __init__(self, widget, text_getter):
        self.widget = widget
        self.text_getter = text_getter if callable(text_getter) else lambda: text_getter
        self.tip = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)
        self.widget.bind("<ButtonPress>", self.hide)

    def show(self, event=None):
        text = clean(self.text_getter(), "")
        if not text or self.tip is not None:
            return
        x = (event.x_root if event else self.widget.winfo_rootx()) + 14
        y = (event.y_root if event else self.widget.winfo_rooty()) + 18
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip,
            text=text,
            justify="left",
            bg="#172033",
            fg="#ffffff",
            padx=9,
            pady=6,
            wraplength=320,
            font=("Segoe UI", 9),
        )
        label.pack()

    def hide(self, _event=None):
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


class TreeToolTip:
    def __init__(self, tree, key_column_index=0):
        self.tree = tree
        self.key_column_index = key_column_index
        self.tip = None
        self.current_text = ""
        tree.bind("<Motion>", self.on_motion, add="+")
        tree.bind("<Leave>", self.hide, add="+")

    def on_motion(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            self.hide()
            return
        values = self.tree.item(row_id, "values")
        if len(values) <= self.key_column_index:
            self.hide()
            return
        key = str(values[self.key_column_index])
        text = TOOLTIP_TEXT.get(key) or TOOLTIP_TEXT.get(key.split(":")[-1].strip())
        if not text:
            self.hide()
            return
        if text == self.current_text and self.tip is not None:
            return
        self.hide()
        self.current_text = text
        self.tip = tk.Toplevel(self.tree)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{event.x_root + 14}+{event.y_root + 16}")
        tk.Label(
            self.tip,
            text=text,
            justify="left",
            bg="#172033",
            fg="#ffffff",
            padx=9,
            pady=6,
            wraplength=340,
            font=("Segoe UI", 9),
        ).pack()

    def hide(self, _event=None):
        self.current_text = ""
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


class CollapsibleSection(tk.Frame):
    def __init__(self, parent, title, open_by_default=False, colors=None):
        self.colors = colors or {"bg": "#f7f9fc", "card": "#ffffff", "border": "#dbe4ef", "text": "#172033", "muted": "#64748b"}
        super().__init__(parent, bg=self.colors["bg"])
        self.is_open = BooleanVar(value=open_by_default)
        self.header = tk.Frame(self, bg=self.colors["card"], highlightbackground=self.colors["border"], highlightthickness=1)
        self.header.pack(fill="x", pady=(8, 2))
        self.header_label = tk.Label(
            self.header,
            text=self.button_text(title),
            bg=self.colors["card"],
            fg=self.colors["text"],
            font=("Segoe UI", 9, "bold"),
            anchor="center",
            pady=7,
        )
        self.header_label.pack(fill="x")
        self.header.bind("<Button-1>", lambda _event: self.toggle(title))
        self.header_label.bind("<Button-1>", lambda _event: self.toggle(title))
        self.header.configure(cursor="hand2")
        self.header_label.configure(cursor="hand2")
        self.body = ttk.Frame(self)
        if open_by_default:
            self.body.pack(fill="x")

    def button_text(self, title):
        return ("Hide " if self.is_open.get() else "Show ") + title

    def toggle(self, title):
        if self.is_open.get():
            self.is_open.set(False)
            self.body.pack_forget()
        else:
            self.is_open.set(True)
            self.body.pack(fill="x")
        self.header_label.configure(text=self.button_text(title))


class AssetInfoApp(Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("1180x760")
        self.minsize(980, 640)
        self.settings = load_settings()
        self.report = OrderedDict()
        self.privacy = BooleanVar(value=bool(self.settings.get("privacy", False)))
        self.advanced = BooleanVar(value=bool(self.settings.get("advanced", False)))
        self.dark_mode = BooleanVar(value=bool(self.settings.get("dark_mode", False)))
        self.search_var = StringVar(value="")
        self.export_folder = StringVar(value=self.settings.get("export_folder", str(Path.home() / "Documents")))
        self.current_view = StringVar(value="Dashboard")
        self.status = StringVar(value="Ready.")
        self.recent_exports = []
        self.nav_buttons = {}
        self.icon_cache = {}
        self.full_report_loaded = False
        self.network_monitor_after = None
        self.network_last_counters = None
        self.network_last_time = None
        self.network_history = []
        self.network_widgets = {}
        self.speed_test_running = False
        self.speed_event_queue = queue.Queue()
        self.last_speed_text = "Speed test: not run yet"
        self.last_speed_level = None
        self.progress_value = tk.DoubleVar(value=0)
        self.icon_photo = None
        self.colors = {}
        self.configure_style()
        self.apply_icon()
        self.build_ui()
        self.search_var.trace_add("write", self.on_search_changed)
        self.after(100, lambda: self.refresh_scan(deep=False))

    def configure_style(self):
        self.colors = self.theme_colors()
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        self.configure(bg=self.colors["bg"])
        style.configure("TFrame", background=self.colors["bg"])
        style.configure("Card.TFrame", background=self.colors["card"], borderwidth=0)
        style.configure("Header.TLabel", background=self.colors["header"], foreground=self.colors["text"], font=("Segoe UI", 17, "bold"))
        style.configure("SubHeader.TLabel", background=self.colors["header"], foreground=self.colors["muted"], font=("Segoe UI", 10))
        style.configure("TLabel", background=self.colors["bg"], foreground=self.colors["text"], font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background=self.colors["card"], foreground=self.colors["muted"], font=("Segoe UI", 9, "bold"))
        style.configure("CardValue.TLabel", background=self.colors["card"], foreground=self.colors["text"], font=("Segoe UI", 12, "bold"))
        button_bg = self.colors["card"] if self.dark_mode.get() else "#ffffff"
        button_active = self.colors["nav_idle"] if self.dark_mode.get() else "#eef4ff"
        style.configure(
            "TButton",
            font=("Segoe UI", 9),
            padding=(10, 6),
            background=button_bg,
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
            focuscolor=button_bg,
            relief="flat",
        )
        style.map("TButton", background=[("active", button_active), ("pressed", self.colors["nav_selected"])], foreground=[("pressed", "#ffffff"), ("active", self.colors["text"])])
        style.configure("TCheckbutton", background=self.colors["bg"], foreground=self.colors["text"], font=("Segoe UI", 9))
        style.configure(
            "TMenubutton",
            font=("Segoe UI", 9),
            padding=(10, 6),
            background=button_bg,
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
            focuscolor=button_bg,
            relief="flat",
        )
        style.map("TMenubutton", background=[("active", button_active), ("pressed", self.colors["nav_selected"])], foreground=[("pressed", "#ffffff"), ("active", self.colors["text"])])
        style.configure("TEntry", fieldbackground=self.colors["card"], foreground=self.colors["text"], insertcolor=self.colors["text"], bordercolor=self.colors["border"], lightcolor=self.colors["border"], darkcolor=self.colors["border"])
        heading_bg = "#243247" if self.dark_mode.get() else "#edf2f7"
        style.configure(
            "Treeview",
            rowheight=28,
            font=("Segoe UI", 9),
            fieldbackground=self.colors["card"],
            background=self.colors["card"],
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background=heading_bg,
            foreground=self.colors["text"],
            bordercolor=self.colors["border"],
            relief="flat",
            padding=(8, 6),
        )
        style.map("Treeview.Heading", background=[("active", heading_bg)], foreground=[("active", self.colors["text"])])
        style.map("Treeview", background=[("selected", self.colors["nav_selected"])], foreground=[("selected", "#ffffff")])
        style.configure("Status.Horizontal.TProgressbar", troughcolor=self.colors["bg"], background=self.colors["nav_selected"], bordercolor=self.colors["border"], lightcolor=self.colors["nav_selected"], darkcolor=self.colors["nav_selected"])

    def theme_colors(self):
        if self.dark_mode.get():
            return {
                "bg": "#111827",
                "header": "#0f172a",
                "sidebar": "#111827",
                "card": "#1f2937",
                "border": "#334155",
                "text": "#f8fafc",
                "muted": "#cbd5e1",
                "nav_selected": "#2563eb",
                "nav_text": "#f8fafc",
                "nav_idle": "#1f2937",
            }
        return {
            "bg": "#f7f9fc",
            "header": "#ffffff",
            "sidebar": "#ffffff",
            "card": "#ffffff",
            "border": "#dbe4ef",
            "text": "#172033",
            "muted": "#64748b",
            "nav_selected": "#2f80ed",
            "nav_text": "#172033",
            "nav_idle": "#eef4ff",
        }

    def nav_icon(self, section, selected=False):
        if Image is None or ImageDraw is None or ImageTk is None:
            return None
        kind = NAV_ICON_KEYS.get(section, "overview")
        cache_key = (kind, selected, self.dark_mode.get())
        if cache_key in self.icon_cache:
            return self.icon_cache[cache_key]
        size = 22
        bg = self.colors["nav_selected"] if selected else self.colors["nav_idle"]
        fg = "#ffffff" if selected else self.colors["nav_selected"]
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle((1, 1, size - 2, size - 2), radius=5, fill=bg)
        self.draw_icon_shape(draw, kind, fg)
        photo = ImageTk.PhotoImage(img)
        self.icon_cache[cache_key] = photo
        return photo

    def draw_icon_shape(self, draw, kind, fg):
        if kind == "dashboard":
            for box in ((6, 6, 10, 10), (12, 6, 16, 10), (6, 12, 10, 16), (12, 12, 16, 16)):
                draw.rectangle(box, outline=fg, width=2)
        elif kind == "warning":
            draw.polygon([(11, 5), (17, 16), (5, 16)], outline=fg, fill=None)
            draw.line((11, 9, 11, 12), fill=fg, width=2)
            draw.point((11, 14), fill=fg)
        elif kind == "helpdesk":
            draw.arc((5, 5, 17, 17), 190, 350, fill=fg, width=2)
            draw.line((5, 12, 5, 15), fill=fg, width=2)
            draw.line((17, 12, 17, 15), fill=fg, width=2)
            draw.line((13, 16, 17, 16), fill=fg, width=2)
        elif kind == "search":
            draw.ellipse((5, 5, 13, 13), outline=fg, width=2)
            draw.line((12, 12, 17, 17), fill=fg, width=2)
        elif kind == "asset":
            draw.rounded_rectangle((5, 7, 17, 15), radius=2, outline=fg, width=2)
            draw.ellipse((7, 10, 9, 12), fill=fg)
            draw.line((11, 10, 15, 10), fill=fg, width=1)
        elif kind == "shield":
            draw.polygon([(11, 5), (16, 7), (15, 14), (11, 17), (7, 14), (6, 7)], outline=fg, width=2)
            draw.line((8, 11, 10, 13, 14, 9), fill=fg, width=2)
        elif kind == "chip":
            draw.rectangle((7, 7, 15, 15), outline=fg, width=2)
            for pos in (5, 11, 17):
                draw.line((pos, 8, pos, 14), fill=fg, width=1)
                draw.line((8, pos, 14, pos), fill=fg, width=1)
        elif kind == "storage":
            draw.ellipse((5, 5, 17, 9), outline=fg, width=2)
            draw.line((5, 7, 5, 15), fill=fg, width=2)
            draw.line((17, 7, 17, 15), fill=fg, width=2)
            draw.ellipse((5, 13, 17, 17), outline=fg, width=2)
        elif kind == "wifi":
            draw.arc((4, 6, 18, 18), 215, 325, fill=fg, width=2)
            draw.arc((7, 9, 15, 17), 215, 325, fill=fg, width=2)
            draw.ellipse((10, 15, 12, 17), fill=fg)
        elif kind == "windows":
            draw.rectangle((5, 5, 10, 10), outline=fg, width=1)
            draw.rectangle((12, 5, 17, 10), outline=fg, width=1)
            draw.rectangle((5, 12, 10, 17), outline=fg, width=1)
            draw.rectangle((12, 12, 17, 17), outline=fg, width=1)
        elif kind == "updates":
            draw.arc((5, 5, 17, 17), 35, 315, fill=fg, width=2)
            draw.polygon([(16, 5), (17, 10), (12, 8)], fill=fg)
        elif kind == "gear":
            draw.ellipse((6, 6, 16, 16), outline=fg, width=2)
            draw.ellipse((9, 9, 13, 13), outline=fg, width=1)
            for x, y in ((11, 4), (11, 18), (4, 11), (18, 11)):
                draw.line((11, 11, x, y), fill=fg, width=1)
        elif kind == "key":
            draw.ellipse((5, 7, 11, 13), outline=fg, width=2)
            draw.line((11, 10, 17, 10), fill=fg, width=2)
            draw.line((15, 10, 15, 13), fill=fg, width=1)
        elif kind == "apps":
            draw.rectangle((5, 5, 9, 9), outline=fg, width=1)
            draw.rectangle((13, 5, 17, 9), outline=fg, width=1)
            draw.rectangle((5, 13, 9, 17), outline=fg, width=1)
            draw.rectangle((13, 13, 17, 17), outline=fg, width=1)
        elif kind == "plug":
            draw.line((8, 5, 8, 10), fill=fg, width=2)
            draw.line((14, 5, 14, 10), fill=fg, width=2)
            draw.line((7, 10, 15, 10), fill=fg, width=2)
            draw.line((11, 10, 11, 17), fill=fg, width=2)
        elif kind == "health":
            draw.line((4, 12, 8, 12, 10, 8, 13, 16, 15, 12, 18, 12), fill=fg, width=2)
        elif kind == "help":
            draw.ellipse((5, 5, 17, 17), outline=fg, width=2)
            draw.text((9, 6), "?", fill=fg)
        else:
            draw.ellipse((5, 5, 17, 17), outline=fg, width=2)
            draw.line((11, 9, 11, 15), fill=fg, width=2)
            draw.point((11, 7), fill=fg)

    def apply_icon(self):
        ico = resource_path(ICON_ICO)
        png = resource_path(ICON_PNG)
        try:
            if ico.exists():
                self.iconbitmap(str(ico))
        except Exception:
            pass
        try:
            if Image is not None and ImageTk is not None and png.exists():
                img = Image.open(png).convert("RGBA").resize((44, 44))
                self.icon_photo = ImageTk.PhotoImage(img)
                self.iconphoto(True, self.icon_photo)
        except Exception:
            pass

    def build_ui(self):
        header = tk.Frame(self, bg=self.colors["header"], height=78, highlightbackground=self.colors["border"], highlightthickness=1)
        self.header = header
        header.pack(fill="x")
        brand = tk.Frame(header, bg=self.colors["header"])
        brand.pack(side="left", padx=18, pady=12)
        if self.icon_photo is not None:
            tk.Label(brand, image=self.icon_photo, bg=self.colors["header"]).pack(side="left", padx=(0, 12))
        text = tk.Frame(brand, bg=self.colors["header"])
        text.pack(side="left")
        ttk.Label(text, text=APP_TITLE, style="Header.TLabel").pack(anchor="w")
        ttk.Label(text, text="Dashboard first, deep details when you need them", style="SubHeader.TLabel").pack(anchor="w")
        top_actions = tk.Frame(header, bg=self.colors["header"])
        top_actions.pack(side="right", padx=18, pady=12)
        ttk.Label(top_actions, text=f"Version {APP_VERSION}", style="SubHeader.TLabel").pack(anchor="e", pady=(0, 6))
        about_icon = self.nav_icon("About", selected=False)
        about_btn = ttk.Button(top_actions, text="About", image=about_icon, compound="left", command=self.show_about)
        about_btn.pack(anchor="e")
        self.icon_cache[("about_button", self.dark_mode.get())] = about_icon
        ToolTip(about_btn, "Show program version, purpose, and export details.")

        toolbar = ttk.Frame(self, padding=(8, 8, 8, 5))
        toolbar.pack(fill="x")
        refresh_btn = ttk.Button(toolbar, text="Refresh", command=lambda: self.refresh_scan(deep=False))
        refresh_btn.pack(side="left", padx=(0, 5))
        ToolTip(refresh_btn, "Run a fast scan and update the dashboard.")
        full_scan_btn = ttk.Button(toolbar, text="Full Scan", command=lambda: self.refresh_scan(deep=True))
        full_scan_btn.pack(side="left", padx=5)
        ToolTip(full_scan_btn, "Run slower checks such as updates, certificates, USB history, event logs, and disk health.")
        copy_btn = ttk.Button(toolbar, text="Copy", command=self.copy_all)
        copy_btn.pack(side="left", padx=5)
        ToolTip(copy_btn, "Copy the current view in a clean Category, Information, Details format.")
        export_button = ttk.Menubutton(toolbar, text="Export")
        export_menu = Menu(export_button, tearoff=0)
        export_menu.add_command(label="Save all reports", command=self.export_all)
        export_menu.add_separator()
        for label, ext in (("Text report", "txt"), ("CSV report", "csv"), ("JSON report", "json"), ("Excel workbook", "xlsx"), ("PDF report", "pdf")):
            export_menu.add_command(label=label, command=lambda value=ext: self.export(value))
        export_button["menu"] = export_menu
        export_button.pack(side="left", padx=5)
        ToolTip(export_button, "Save one report format or all report formats at once.")
        screenshot_btn = ttk.Button(toolbar, text="Screenshot", command=self.take_screenshot)
        screenshot_btn.pack(side="left", padx=5)
        ToolTip(screenshot_btn, "Save a screenshot to Pictures and copy the image to the clipboard.")
        folder_btn = ttk.Button(toolbar, text="Folder", command=self.open_output_folder)
        folder_btn.pack(side="left", padx=5)
        ToolTip(folder_btn, "Open the folder used for reports.")
        settings_btn = ttk.Button(toolbar, text="Settings", command=self.show_settings)
        settings_btn.pack(side="left", padx=5)
        ToolTip(settings_btn, "Set privacy mode, advanced details, dark mode, and default export folder.")
        search_box = ttk.Frame(toolbar)
        search_box.pack(side="left", fill="x", expand=True, padx=(18, 8))
        ttk.Label(search_box, text="Search").pack(side="left", padx=(0, 6))
        search_entry = ttk.Entry(search_box, textvariable=self.search_var)
        search_entry.pack(side="left", fill="x", expand=True)
        ToolTip(search_entry, "Search every category, field, and detailed table. Try tags like wifi, speed, disk, battery, security, updates, hardware, or asset.")
        tags_button = ttk.Menubutton(search_box, text="Tags")
        tags_menu = Menu(tags_button, tearoff=0)
        for tag in TAG_MENU:
            tags_menu.add_command(label=tag, command=lambda value=tag: self.search_var.set(value))
        tags_button["menu"] = tags_menu
        tags_button.pack(side="left", padx=(6, 0))
        ToolTip(tags_button, "Quick search tags for common support areas.")
        clear_btn = ttk.Button(search_box, text="Clear", command=lambda: self.search_var.set(""))
        clear_btn.pack(side="left", padx=(6, 0))
        ToolTip(clear_btn, "Clear search and return to the selected view.")

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.sidebar = tk.Frame(body, bg=self.colors["sidebar"], width=180, highlightbackground=self.colors["border"], highlightthickness=1)
        self.sidebar.pack(side="left", fill="y", padx=(0, 6))
        self.sidebar.pack_propagate(False)
        self.content_host = ttk.Frame(body)
        self.content_host.pack(side="left", fill="both", expand=True)
        status_bar = tk.Frame(self, bg=self.colors["header"], highlightbackground=self.colors["border"], highlightthickness=1)
        self.status_bar = status_bar
        status_bar.pack(fill="x", padx=6, pady=(0, 6))
        status_inner = tk.Frame(status_bar, bg=self.colors["header"])
        status_inner.pack(fill="x", padx=10, pady=7)
        status_dot = tk.Label(status_inner, text=" ", bg="#12b76a", width=2, height=1)
        self.status_dot = status_dot
        status_dot.pack(side="left", padx=(0, 8))
        tk.Label(status_inner, text="Status", bg=self.colors["header"], fg=self.colors["muted"], font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 8))
        self.status_label = tk.Label(status_inner, textvariable=self.status, bg=self.colors["header"], fg=self.colors["text"], font=("Segoe UI", 9), anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)
        self.progress = ttk.Progressbar(status_inner, variable=self.progress_value, maximum=100, length=170, style="Status.Horizontal.TProgressbar")
        self.progress.pack(side="left", padx=12)
        self.recent_label = tk.Label(status_inner, text="No exports yet", bg=self.colors["header"], fg=self.colors["muted"], font=("Segoe UI", 9), anchor="e")
        self.recent_label.pack(side="right")

    def refresh_scan(self, deep=False):
        self.status.set("Starting full scan..." if deep else "Starting fast scan...")
        self.progress_value.set(10)
        self.update_idletasks()
        try:
            self.status.set("Checking core system details..." if not deep else "Checking security, storage, network, and updates...")
            self.progress_value.set(35)
            self.update_idletasks()
            self.report = collect_report(deep=deep)
            self.full_report_loaded = bool(deep)
            self.status.set("Building dashboard...")
            self.progress_value.set(80)
            self.update_idletasks()
            self.render_report()
            self.progress_value.set(100)
            self.status.set("Full scan complete." if deep else "Fast scan complete. Use Full Scan for advanced details.")
        except Exception as exc:
            self.status.set("Scan failed.")
            messagebox.showerror("Scan failed", str(exc))

    def ensure_full_report(self):
        if self.full_report_loaded:
            return
        self.status.set("Running full scan before export...")
        self.progress_value.set(15)
        self.update_idletasks()
        self.report = collect_report(deep=True)
        self.full_report_loaded = True
        self.progress_value.set(90)
        self.render_report()
        self.progress_value.set(100)
        self.status.set("Full scan complete.")

    def show_about(self):
        details = collect_about_info()
        messagebox.showinfo("About", "\n".join(f"{key}: {value}" for key, value in details.items()))

    def show_settings(self):
        window = tk.Toplevel(self)
        window.title("Settings")
        window.geometry("520x270")
        window.configure(bg=self.colors["bg"])
        window.transient(self)
        window.grab_set()
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Settings", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))
        for label, variable, tip in (
            ("Privacy mode by default", self.privacy, "Hide serials, users, IPs, MAC addresses, owner, and similar identifiers."),
            ("Show advanced details", self.advanced, "Show all detailed tables such as services, updates, certificates, USB history, and events."),
            ("Dark mode", self.dark_mode, "Switch the app to a darker visual theme."),
        ):
            row = ttk.Checkbutton(frame, text=label, variable=variable, command=self.apply_settings)
            row.pack(anchor="w", pady=3)
            ToolTip(row, tip)
        folder_row = ttk.Frame(frame)
        folder_row.pack(fill="x", pady=(12, 4))
        ttk.Label(folder_row, text="Default export folder").pack(anchor="w")
        entry_row = ttk.Frame(frame)
        entry_row.pack(fill="x")
        entry = ttk.Entry(entry_row, textvariable=self.export_folder)
        entry.pack(side="left", fill="x", expand=True)
        ToolTip(entry, "Reports save here by default.")
        ttk.Button(entry_row, text="Choose", command=self.choose_export_folder).pack(side="left", padx=(6, 0))
        ttk.Button(frame, text="Close", command=window.destroy).pack(anchor="e", pady=(16, 0))

    def choose_export_folder(self):
        folder = filedialog.askdirectory(initialdir=self.export_folder.get() or str(Path.home() / "Documents"), title="Choose default export folder")
        if folder:
            self.export_folder.set(folder)
            self.apply_settings()

    def apply_settings(self):
        self.settings.update(
            {
                "privacy": bool(self.privacy.get()),
                "advanced": bool(self.advanced.get()),
                "dark_mode": bool(self.dark_mode.get()),
                "export_folder": self.export_folder.get() or str(Path.home() / "Documents"),
            }
        )
        save_settings(self.settings)
        self.configure_style()
        self.refresh_static_colors()
        self.render_report()

    def refresh_static_colors(self):
        for widget_name in ("header", "sidebar", "status_bar"):
            widget = getattr(self, widget_name, None)
            if widget is not None:
                color = self.colors["sidebar"] if widget_name == "sidebar" else self.colors["header"]
                try:
                    widget.configure(bg=color, highlightbackground=self.colors["border"])
                except Exception:
                    pass
                self.refresh_widget_tree_colors(widget, color)
        try:
            self.status_dot.configure(bg="#12b76a")
            self.status_label.configure(bg=self.colors["header"], fg=self.colors["text"])
            self.recent_label.configure(bg=self.colors["header"], fg=self.colors["muted"])
        except Exception:
            pass

    def refresh_widget_tree_colors(self, widget, bg_color):
        for child in widget.winfo_children():
            try:
                child.configure(bg=bg_color)
            except Exception:
                pass
            try:
                if isinstance(child, tk.Label):
                    child.configure(fg=self.colors["muted"] if child.cget("text") in ("Status", "No exports yet") else self.colors["text"])
            except Exception:
                pass
            self.refresh_widget_tree_colors(child, bg_color)

    def on_search_changed(self, *_args):
        if self.search_var.get().strip():
            self.current_view.set("Search")
        elif self.current_view.get() == "Search":
            self.current_view.set("Dashboard")
        if getattr(self, "content_host", None) is not None and self.report:
            self.render_report()

    def tab_label(self, section):
        return TAB_LABELS.get(section, section)

    def sections_for_view(self):
        sections = ["Dashboard", "Problems", "Helpdesk"]
        simple = ["Overview", "Asset & Ownership", "Security", "Hardware", "Storage", "Network", "Windows & Updates", "Activity & Health"]
        sections.extend([section for section in simple if section in self.report])
        if self.advanced.get():
            sections.extend(section for section in self.report.keys() if section not in sections)
        if self.search_var.get().strip():
            sections.insert(0, "Search")
        return sections

    def select_view(self, section):
        self.current_view.set(section)
        self.render_report()

    def bind_click_recursive(self, widget, command):
        try:
            widget.configure(cursor="hand2")
        except Exception:
            pass
        widget.bind("<Button-1>", lambda _event: command())
        for child in widget.winfo_children():
            self.bind_click_recursive(child, command)

    def render_nav(self):
        for child in self.sidebar.winfo_children():
            child.destroy()
        tk.Label(self.sidebar, text="Views", bg=self.colors["sidebar"], fg=self.colors["muted"], font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(12, 6))
        self.nav_buttons = {}
        sections = self.sections_for_view()
        if self.current_view.get() not in sections:
            self.current_view.set("Dashboard")
        for section in sections:
            selected = section == self.current_view.get()
            icon = self.nav_icon(section, selected=selected)
            button = tk.Button(
                self.sidebar,
                text=self.tab_label(section),
                image=icon,
                compound="left",
                anchor="w",
                relief="flat",
                bd=0,
                padx=10,
                pady=7,
                bg=self.colors["nav_selected"] if selected else self.colors["sidebar"],
                fg="#ffffff" if selected else self.colors["nav_text"],
                activebackground=self.colors["nav_idle"],
                activeforeground=self.colors["nav_text"],
                font=("Segoe UI", 9, "bold" if selected else "normal"),
                command=lambda value=section: self.select_view(value),
            )
            button.pack(fill="x", padx=8, pady=1)
            self.nav_buttons[section] = button
            ToolTip(button, self.nav_tooltip(section))

    def nav_tooltip(self, section):
        tips = {
            "Dashboard": "A quick summary of the computer, health, and most important warnings.",
            "Problems": "Only the items that need attention or review.",
            "Helpdesk": "A compact support view with the fields most often requested.",
            "Search": "Search results from every visible and advanced report field.",
        }
        return tips.get(section, f"View {self.tab_label(section)} information.")

    def visible_items(self, section, items):
        if self.advanced.get():
            return items
        keys = SIMPLE_FIELDS.get(section)
        if not keys:
            return items
        return OrderedDict((key, items.get(key, "Unavailable")) for key in keys if key in items)

    def search_results(self):
        query = self.search_var.get().strip().lower()
        if not query:
            return []
        terms = set([query])
        for tag, aliases in SEARCH_TAGS.items():
            if query == tag or query in aliases:
                terms.update(alias.lower() for alias in aliases)
        rows = []
        for section, bundle in self.report.items():
            for key, value in bundle.get("items", {}).items():
                text = f"{section} {key} {value}".lower()
                if any(term in text for term in terms):
                    rows.append(OrderedDict([("Category", self.tab_label(section)), ("Table", ""), ("Information", key), ("Details", value)]))
            for table_name, table_rows in bundle.get("tables", {}).items():
                for index, row in enumerate(table_rows, 1):
                    for key, value in row.items():
                        text = f"{section} {table_name} {key} {value}".lower()
                        if any(term in text for term in terms):
                            rows.append(OrderedDict([("Category", self.tab_label(section)), ("Table", f"{table_name} #{index}"), ("Information", key), ("Details", value)]))
                            if len(rows) >= 200:
                                return rows
        return rows

    def health_findings(self):
        findings = []
        overview = self.report.get("Overview", {}).get("items", {})
        storage = self.report.get("Storage", {}).get("items", {})
        security = self.report.get("Security", {}).get("items", {})
        hardware = self.report.get("Hardware", {}).get("items", {})
        network = self.report.get("Network", {}).get("items", {})
        reboot = clean(overview.get("Pending Reboot"), "")
        if reboot.lower().startswith("yes"):
            findings.append(("Action needed", "Restart pending"))
        try:
            used = float(str(storage.get("C: Percent Used", "")).replace("%", ""))
            if used >= 90:
                findings.append(("Action needed", "C: drive is nearly full"))
            elif used >= 80:
                findings.append(("Watch", "C: drive space is getting low"))
        except Exception:
            pass
        firewall = clean(security.get("Firewall Profile Status"), "").lower()
        if "off" in firewall:
            findings.append(("Action needed", "Firewall profile is off"))
        defender = clean(security.get("Microsoft Defender Status"), "").lower()
        if "real-time=no" in defender or "antivirus=no" in defender:
            findings.append(("Action needed", "Defender protection may be off"))
        if clean(network.get("Primary IP Address"), "Unavailable") == "Unavailable":
            findings.append(("Watch", "Network address unavailable"))
        battery = clean(hardware.get("Battery Health"), "")
        if "health" in battery.lower():
            try:
                pct = int(re.search(r"Health\s+(\d+)%", battery).group(1))
                if pct < 70:
                    findings.append(("Watch", "Battery health is low"))
            except Exception:
                pass
        if not findings:
            findings.append(("Healthy", "No urgent warnings found"))
        return findings

    def render_status_banner(self, parent, privacy):
        findings = self.health_findings()
        first_label, first_message = findings[0]
        colors = {
            "Healthy": ("#ecfdf3", "#027a48", "#abefc6"),
            "Watch": ("#fffaeb", "#b54708", "#fedf89"),
            "Action needed": ("#fff1f3", "#c01048", "#fecdd6"),
        }
        bg, fg, border = colors.get(first_label, ("#eff6ff", "#175cd3", "#bfdbfe"))
        banner = tk.Frame(parent, bg=bg, highlightbackground=border, highlightthickness=1)
        banner.pack(fill="x", padx=4, pady=(6, 4))
        tk.Label(banner, text=first_label, bg=bg, fg=fg, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=(8, 2))
        messages = "; ".join(message for _label, message in findings[:3])
        tk.Label(banner, text=mask_value("Health Summary", messages, privacy), bg=bg, fg="#344054", font=("Segoe UI", 10)).pack(anchor="w", padx=10, pady=(0, 8))

    def render_report(self):
        self.render_nav()
        for child in self.content_host.winfo_children():
            child.destroy()
        privacy = self.privacy.get()
        frame = ScrollFrame(self.content_host, bg=self.colors["bg"])
        frame.pack(fill="both", expand=True)
        section = self.current_view.get()
        if section == "Search":
            results = self.search_results()
            self.render_page_title(frame.inner, "Search", f"{len(results)} match(es) found")
            if results:
                self.render_table(frame.inner, f"Search results for \"{self.search_var.get().strip()}\"", results, privacy, open_by_default=True)
            else:
                ttk.Label(frame.inner, text="No matches found.", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=10)
        elif section == "Dashboard":
            self.render_dashboard(frame.inner, privacy)
        elif section == "Problems":
            self.render_problems(frame.inner, privacy)
        elif section == "Helpdesk":
            self.render_helpdesk(frame.inner, privacy)
        else:
            self.render_section(frame.inner, section, privacy)

    def render_page_title(self, parent, title, subtitle=""):
        box = tk.Frame(parent, bg=self.colors["bg"])
        box.pack(fill="x")
        tk.Label(box, text=title, bg=self.colors["bg"], fg=self.colors["text"], font=("Segoe UI", 16, "bold"), anchor="center").pack(fill="x", padx=4, pady=(8, 2))
        if subtitle:
            tk.Label(box, text=subtitle, bg=self.colors["bg"], fg=self.colors["muted"], font=("Segoe UI", 10), anchor="center").pack(fill="x", padx=4, pady=(0, 8))

    def badge(self, parent, text, level="Good"):
        colors = {
            "Good": ("#ecfdf3", "#027a48", "#abefc6"),
            "Check": ("#eff6ff", "#175cd3", "#bfdbfe"),
            "Warning": ("#fffaeb", "#b54708", "#fedf89"),
            "Needs Action": ("#fff1f3", "#c01048", "#fecdd6"),
        }
        bg, fg, border = colors.get(level, colors["Check"])
        label = tk.Label(parent, text=text, bg=bg, fg=fg, font=("Segoe UI", 9, "bold"), padx=8, pady=3, highlightbackground=border, highlightthickness=1)
        return label

    def valid_network_ip(self, value):
        text = clean(value, "")
        if not text or text == "Unavailable":
            return False
        if text.startswith("127.") or text.startswith("169.254.") or text == "0.0.0.0":
            return False
        return bool(re.match(r"^\d{1,3}(\.\d{1,3}){3}$", text))

    def metric_level(self, name, value):
        text = clean(value, "").lower()
        if name == "Disk":
            try:
                used = float(str(self.report.get("Storage", {}).get("items", {}).get("C: Percent Used", "")).replace("%", ""))
                if used >= 90:
                    return "Needs Action"
                if used >= 80:
                    return "Warning"
            except Exception:
                return "Check"
            return "Good"
        if name == "Security":
            if "run full scan" in text:
                return "Check"
            if "off" in text or "not detected" in text or "real-time=no" in text:
                return "Needs Action"
            return "Good"
        if name == "Updates":
            if "run full scan" in text:
                return "Check"
            return "Needs Action" if text.lower().startswith("yes") else "Good"
        if name == "Battery":
            if "run full scan" in text:
                return "Check"
            if "no battery" in text:
                return "Check"
            if "health" in text:
                try:
                    pct = int(re.search(r"health\s+(\d+)%", text).group(1))
                    return "Warning" if pct < 70 else "Good"
                except Exception:
                    pass
            return "Check"
        if name == "Network":
            ip = self.report.get("Network", {}).get("items", {}).get("Primary IP Address", "")
            return "Good" if self.valid_network_ip(ip) else "Warning"
        return "Check"

    def metric_reason(self, name, value, detail):
        level = self.metric_level(name, f"{value} {detail}")
        if name == "Network":
            ip = self.report.get("Network", {}).get("items", {}).get("Primary IP Address", "")
            ssid = clean(self.report.get("Network", {}).get("items", {}).get("Wi-Fi SSID"), "")
            if self.valid_network_ip(ip):
                if "not connected" in ssid.lower() or "unavailable" in ssid.lower():
                    return "A valid IP address is present, so the computer is connected. Windows did not allow the app to read the Wi-Fi name, often because Location access or elevation is required."
                return f"A valid IP address is present and Windows reports the network name as {ssid}."
            return "No valid primary IP address was found, so the network tile needs review."
        if name == "Disk":
            used = self.report.get("Storage", {}).get("items", {}).get("C: Percent Used", "Unavailable")
            return f"C: drive usage is {used}. Warning starts at 80%; Needs Action starts at 90%."
        if name == "Security":
            if level == "Check":
                return "Fast startup scan skips slower security checks. Use Full Scan or export a report for complete security status."
            return "Security warning is based on antivirus, firewall, Defender, or encryption text returned by Windows."
        if name == "Updates":
            if level == "Check":
                return "Fast startup scan skips slower update checks. Use Full Scan or export a report to check pending updates."
            return "Updates warning appears when Windows reports a pending reboot or pending update condition."
        if name == "Battery":
            return "Battery status is based on Windows battery health and charge data when available."
        return "Status is based on the latest scan data for this category."

    def network_tile_detail(self, network):
        ip = network.get("Primary IP Address", "")
        ssid = clean(network.get("Wi-Fi SSID"), "")
        if self.valid_network_ip(ip) and ("not connected" in ssid.lower() or "unavailable" in ssid.lower()):
            return "Connected; network name unavailable"
        return ssid

    def render_dashboard(self, parent, privacy):
        self.render_page_title(parent, "Dashboard", "The important computer details at a glance")
        self.render_status_banner(parent, privacy)
        tiles = ttk.Frame(parent, padding=(4, 6, 4, 4))
        tiles.pack(fill="x")
        storage = self.report.get("Storage", {}).get("items", {})
        security = self.report.get("Security", {}).get("items", {})
        hardware = self.report.get("Hardware", {}).get("items", {})
        network = self.report.get("Network", {}).get("items", {})
        overview = self.report.get("Overview", {}).get("items", {})
        tile_data = [
            ("Disk", storage.get("C: Free Space", "Unavailable"), storage.get("C: Percent Used", "Unavailable"), "Disk free space and usage."),
            ("Security", security.get("Virus Protection", "Unavailable"), security.get("BitLocker Status", "Unavailable"), "Antivirus and encryption state."),
            ("Updates", overview.get("Pending Reboot", "Unavailable"), self.report.get("Windows & Updates", {}).get("items", {}).get("Missing / Pending Updates", "Unavailable"), "Pending restart and update state."),
            ("Battery", hardware.get("Battery Health", "Unavailable"), "", "Battery health and charge when available."),
            ("Network", network.get("Primary IP Address", "Unavailable"), self.network_tile_detail(network), "Primary network and wireless status."),
        ]
        for index, (title, value, detail, tip) in enumerate(tile_data):
            self.render_status_tile(tiles, index, title, value, detail, tip, privacy)
        self.render_quick_facts(parent, privacy)
        self.render_plain_summary(parent)

    def render_status_tile(self, parent, index, title, value, detail, tip, privacy):
        level = self.metric_level(title, f"{value} {detail}")
        reason = self.metric_reason(title, value, detail)
        card = tk.Frame(parent, bg=self.colors["card"], highlightbackground=self.colors["border"], highlightthickness=1)
        card.grid(row=index // 3, column=index % 3, sticky="nsew", padx=4, pady=4)
        parent.grid_columnconfigure(index % 3, weight=1)
        parent.grid_rowconfigure(index // 3, weight=1, minsize=132)
        top = tk.Frame(card, bg=self.colors["card"])
        top.pack(fill="x", padx=10, pady=(10, 3))
        tile_icon_map = {"Disk": "Storage", "Security": "Security", "Updates": "Windows & Updates", "Battery": "Activity & Health", "Network": "Network"}
        icon = self.nav_icon(tile_icon_map.get(title, title), selected=False)
        if icon is not None:
            tk.Label(top, image=icon, bg=self.colors["card"]).pack(side="left", padx=(0, 6))
        tk.Label(top, text=title, bg=self.colors["card"], fg=self.colors["text"], font=("Segoe UI", 11, "bold")).pack(side="left")
        self.badge(top, level, level).pack(side="right")
        tk.Label(card, text=mask_value(title, value, privacy), bg=self.colors["card"], fg=self.colors["text"], font=("Segoe UI", 12, "bold"), wraplength=260, justify="left").pack(anchor="w", padx=10, fill="x")
        if detail:
            tk.Label(card, text=mask_value(title, detail, privacy), bg=self.colors["card"], fg=self.colors["muted"], font=("Segoe UI", 9), wraplength=260, justify="left").pack(anchor="w", padx=10, pady=(2, 10), fill="x")
        if title == "Disk":
            self.render_usage_bar(card, self.percent_value(self.report.get("Storage", {}).get("items", {}).get("C: Percent Used")), compact=True)
        ToolTip(card, f"{tip}\n\nStatus: {level}\nReason: {reason}")
        target_map = {
            "Disk": "Storage",
            "Security": "Security",
            "Updates": "Windows & Updates",
            "Battery": "Activity & Health",
            "Network": "Network",
        }
        target = target_map.get(title)
        if target:
            self.bind_click_recursive(card, lambda value=target: self.select_view(value))

    def render_quick_facts(self, parent, privacy):
        overview = self.report.get("Overview", {}).get("items", {})
        storage = self.report.get("Storage", {}).get("items", {})
        network = self.report.get("Network", {}).get("items", {})
        facts = OrderedDict(
            [
                ("Computer Name", overview.get("Computer Name")),
                ("Model", overview.get("Model")),
                ("Serial Number", overview.get("Serial Number")),
                ("Operating System", overview.get("Operating System")),
                ("Windows Release", overview.get("Windows Release")),
                ("RAM Installed", overview.get("RAM Installed")),
                ("Primary IP Address", network.get("Primary IP Address")),
                ("C: Free Space", storage.get("C: Free Space")),
            ]
        )
        box = tk.Frame(parent, bg=self.colors["card"], highlightbackground=self.colors["border"], highlightthickness=1)
        box.pack(fill="x", padx=4, pady=(8, 4))
        tk.Label(box, text="Computer Snapshot", bg=self.colors["card"], fg=self.colors["text"], font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(10, 2))
        tk.Label(box, text="Most-used support details, lined up for quick reading.", bg=self.colors["card"], fg=self.colors["muted"], font=("Segoe UI", 9)).pack(anchor="w", padx=12, pady=(0, 8))
        grid = tk.Frame(box, bg=self.colors["card"])
        grid.pack(fill="x", padx=10, pady=(0, 10))
        for column in range(2):
            grid.grid_columnconfigure(column, weight=1, uniform="snapshot")
        for index, (key, value) in enumerate(facts.items()):
            cell = tk.Frame(grid, bg=self.colors["card"])
            cell.grid(row=index // 2, column=index % 2, sticky="ew", padx=2, pady=2)
            label = tk.Frame(cell, bg=self.colors["card"])
            label.pack(fill="x", padx=6, pady=5)
            tk.Label(label, text=key, bg=self.colors["card"], fg=self.colors["muted"], font=("Segoe UI", 8, "bold"), width=18, anchor="w").pack(side="left")
            tk.Label(label, text=mask_value(key, value, privacy), bg=self.colors["card"], fg=self.colors["text"], font=("Segoe UI", 10, "bold"), anchor="w", justify="left", wraplength=330).pack(side="left", fill="x", expand=True)

    def render_plain_summary(self, parent):
        findings = self.health_findings()
        box = tk.Frame(parent, bg=self.colors["card"], highlightbackground=self.colors["border"], highlightthickness=1)
        box.pack(fill="x", padx=4, pady=8)
        tk.Label(box, text="Plain English Summary", bg=self.colors["card"], fg=self.colors["text"], font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=(10, 2))
        if findings and findings[0][0] == "Healthy":
            text = "This computer does not show urgent warnings in the current scan."
        else:
            text = "Review these items first: " + "; ".join(message for _level, message in findings[:4])
        tk.Label(box, text=text, bg=self.colors["card"], fg=self.colors["muted"], wraplength=880, justify="left").pack(anchor="w", padx=10, pady=(0, 10))

    def render_recent_exports(self, parent):
        box = ttk.Frame(parent, padding=(4, 4, 4, 10))
        box.pack(fill="x")
        ttk.Label(box, text="Recent Exports", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        values = self.recent_exports[-5:] or ["No reports exported in this session."]
        for value in reversed(values):
            ttk.Label(box, text=str(value)).pack(anchor="w", pady=1)

    def row_colors(self):
        if self.dark_mode.get():
            return self.colors["card"], "#263244"
        return self.colors["card"], "#f8fafc"

    def percent_value(self, value):
        try:
            return max(min(float(str(value).replace("%", "").strip()), 100), 0)
        except Exception:
            return None

    def render_usage_bar(self, parent, percent, compact=False, label="Used"):
        if percent is None:
            return
        width = 260 if compact else 520
        height = 10 if compact else 18
        free = 100 - percent
        bar = tk.Canvas(parent, width=width, height=height, bg=self.colors["card"], highlightthickness=0)
        bar.pack(anchor="w", padx=10 if compact else 4, pady=(3, 10 if compact else 8))
        used_width = int(width * percent / 100)
        color = "#12b76a" if percent < 80 else "#f79009" if percent < 90 else "#f04438"
        available_color = "#dbe4ef" if not self.dark_mode.get() else "#334155"
        bar.create_rectangle(0, 0, width, height, fill=available_color, outline="")
        bar.create_rectangle(0, 0, used_width, height, fill=color, outline="")
        if not compact:
            legend = tk.Frame(parent, bg=self.colors["card"])
            legend.pack(anchor="w", padx=4, pady=(0, 8))
            used_marker = tk.Label(legend, text="  ", bg=color, font=("Segoe UI", 7), width=2)
            used_marker.pack(side="left", padx=(0, 5))
            tk.Label(legend, text=f"{label}: {percent:.1f}%", bg=self.colors["card"], fg=self.colors["text"], font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 16))
            free_marker = tk.Label(legend, text="  ", bg=available_color, font=("Segoe UI", 7), width=2)
            free_marker.pack(side="left", padx=(0, 5))
            tk.Label(legend, text=f"Available / not used: {free:.1f}%", bg=self.colors["card"], fg=self.colors["muted"], font=("Segoe UI", 9, "bold")).pack(side="left")

    def render_problems(self, parent, privacy):
        findings = self.health_findings()
        self.render_page_title(parent, "Problems Found", "Only items that need attention or review")
        rows = [OrderedDict([("Status", level), ("Information", "Finding"), ("Details", message)]) for level, message in findings]
        self.render_table(parent, "Problems Found", rows, privacy, open_by_default=True)

    def render_helpdesk(self, parent, privacy):
        self.render_page_title(parent, "Helpdesk View", "Compact support information")
        self.render_table(parent, "Helpdesk Fields", self.helpdesk_rows(), privacy, open_by_default=True)

    def render_section(self, parent, section, privacy):
        if section == "Network":
            self.render_network_page(parent, privacy)
            return
        if section == "Storage":
            self.render_storage_page(parent, privacy)
            return
        bundle = self.report.get(section, {})
        items = self.visible_items(section, bundle.get("items", {}))
        self.render_page_title(parent, self.tab_label(section), f"{len(items)} visible item(s)")
        self.render_items(parent, section, items, privacy)
        if self.advanced.get():
            for table_name, rows in bundle.get("tables", {}).items():
                self.render_table(parent, table_name, rows, privacy)

    def render_storage_page(self, parent, privacy):
        bundle = self.report.get("Storage", {})
        items = self.visible_items("Storage", bundle.get("items", {}))
        self.render_page_title(parent, "Storage", "Used and not-used drive space")
        visual = tk.Frame(parent, bg=self.colors["card"], highlightbackground=self.colors["border"], highlightthickness=1)
        visual.pack(fill="x", padx=4, pady=(6, 8))
        tk.Label(visual, text="C: Drive Usage", bg=self.colors["card"], fg=self.colors["text"], font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=(10, 2))
        self.render_usage_bar(visual, self.percent_value(items.get("C: Percent Used")), compact=False)
        stat_row = tk.Frame(visual, bg=self.colors["card"])
        stat_row.pack(fill="x", padx=10, pady=(0, 10))
        for title, key in (("Total", "C: Total Space"), ("Used", "C: Used Space"), ("Available", "C: Free Space")):
            cell = tk.Frame(stat_row, bg=self.colors["card"])
            cell.pack(side="left", padx=(0, 28))
            tk.Label(cell, text=title, bg=self.colors["card"], fg=self.colors["muted"], font=("Segoe UI", 8, "bold")).pack(anchor="w")
            tk.Label(cell, text=mask_value(key, items.get(key), privacy), bg=self.colors["card"], fg=self.colors["text"], font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.render_items(parent, "Storage", items, privacy)
        if self.advanced.get():
            for table_name, rows in bundle.get("tables", {}).items():
                self.render_table(parent, table_name, rows, privacy)

    def render_network_page(self, parent, privacy):
        bundle = self.report.get("Network", {})
        items = self.visible_items("Network", bundle.get("items", {}))
        self.render_page_title(parent, "Network / Wi-Fi", "Live traffic, connection details, and speed-test help")
        live = tk.Frame(parent, bg=self.colors["card"], highlightbackground=self.colors["border"], highlightthickness=1)
        live.pack(fill="x", padx=4, pady=(6, 8))
        top = tk.Frame(live, bg=self.colors["card"])
        top.pack(fill="x", padx=10, pady=(10, 4))
        tk.Label(top, text="Live Traffic", bg=self.colors["card"], fg=self.colors["text"], font=("Segoe UI", 12, "bold")).pack(side="left")
        speedtest_btn = ttk.Button(top, text="Open Speedtest", command=self.open_speed_test)
        speedtest_btn.pack(side="right", padx=(6, 0))
        quick_btn = ttk.Button(top, text="Run Quick Test", command=self.run_quick_speed_test)
        quick_btn.pack(side="right")
        ToolTip(speedtest_btn, "Open Speedtest.net for a full browser test with download, upload, and ping.")
        ToolTip(quick_btn, "Run a quick download and upload test inside this app, then show an easy rating.")
        rates = tk.Frame(live, bg=self.colors["card"])
        rates.pack(fill="x", padx=10)
        self.network_widgets = {
            "download": tk.Label(rates, text="Incoming: --", bg=self.colors["card"], fg="#175cd3", font=("Segoe UI", 11, "bold")),
            "upload": tk.Label(rates, text="Outgoing: --", bg=self.colors["card"], fg="#027a48", font=("Segoe UI", 11, "bold")),
            "speed": tk.Label(live, text=self.last_speed_text, bg=self.colors["card"], fg=self.speed_level_color(self.last_speed_level), font=("Segoe UI", 10, "bold"), justify="left", wraplength=900),
            "speed_button": quick_btn,
        }
        if self.speed_test_running:
            quick_btn.configure(state="disabled")
        self.network_widgets["download"].pack(side="left", padx=(0, 28), pady=(4, 4))
        self.network_widgets["upload"].pack(side="left", pady=(4, 4))
        tk.Label(
            live,
            text="Incoming means data this computer is receiving. Outgoing means data this computer is sending. The graph updates every second while this page is open.",
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
            justify="left",
            wraplength=900,
        ).pack(anchor="w", padx=10, pady=(2, 2))
        self.network_canvas = tk.Canvas(live, width=760, height=120, bg=self.colors["card"], highlightthickness=0)
        self.network_canvas.pack(fill="x", padx=10, pady=(4, 8))
        self.network_widgets["speed"].pack(anchor="w", padx=10, pady=(0, 10))
        ToolTip(live, "Incoming is download traffic. Outgoing is upload traffic. These numbers are live local network activity, not an internet speed-test result.")
        self.start_network_monitor()
        self.render_items(parent, "Network", items, privacy)
        if self.advanced.get():
            for table_name, rows in bundle.get("tables", {}).items():
                self.render_table(parent, table_name, rows, privacy)

    def format_rate(self, bytes_per_second):
        if bytes_per_second >= 1024 * 1024:
            return f"{bytes_per_second / (1024 * 1024):.2f} MB/s"
        if bytes_per_second >= 1024:
            return f"{bytes_per_second / 1024:.1f} KB/s"
        return f"{bytes_per_second:.0f} B/s"

    def start_network_monitor(self):
        if psutil is None:
            if "download" in self.network_widgets:
                self.network_widgets["download"].configure(text="Incoming: unavailable")
                self.network_widgets["upload"].configure(text="Outgoing: unavailable")
            return
        self.network_last_counters = psutil.net_io_counters()
        self.network_last_time = time.time()
        self.network_history = []
        if self.network_monitor_after is not None:
            try:
                self.after_cancel(self.network_monitor_after)
            except Exception:
                pass
        self.network_monitor_after = self.after(1000, self.update_network_monitor)

    def update_network_monitor(self):
        if self.current_view.get() != "Network" or psutil is None or not getattr(self, "network_widgets", None):
            return
        now = time.time()
        counters = psutil.net_io_counters()
        elapsed = max(now - (self.network_last_time or now), 0.1)
        down = max((counters.bytes_recv - self.network_last_counters.bytes_recv) / elapsed, 0)
        up = max((counters.bytes_sent - self.network_last_counters.bytes_sent) / elapsed, 0)
        self.network_last_counters = counters
        self.network_last_time = now
        self.network_history.append((down, up))
        self.network_history = self.network_history[-60:]
        self.network_widgets["download"].configure(text=f"Incoming: {self.format_rate(down)}")
        self.network_widgets["upload"].configure(text=f"Outgoing: {self.format_rate(up)}")
        self.draw_network_graph()
        self.network_monitor_after = self.after(1000, self.update_network_monitor)

    def draw_network_graph(self):
        canvas = getattr(self, "network_canvas", None)
        if canvas is None:
            return
        canvas.delete("all")
        width = max(canvas.winfo_width(), 400)
        height = 120
        canvas.create_rectangle(0, 0, width, height, fill=self.colors["card"], outline="")
        canvas.create_text(8, 10, anchor="w", text="Incoming", fill="#175cd3", font=("Segoe UI", 8, "bold"))
        canvas.create_text(90, 10, anchor="w", text="Outgoing", fill="#027a48", font=("Segoe UI", 8, "bold"))
        if not self.network_history:
            canvas.create_text(width / 2, height / 2, text="Waiting for traffic...", fill=self.colors["muted"], font=("Segoe UI", 10))
            return
        max_rate = max(max(pair) for pair in self.network_history) or 1
        step = width / max(len(self.network_history) - 1, 1)
        down_points = []
        up_points = []
        for index, (down, up) in enumerate(self.network_history):
            x = index * step
            down_points.append((x, height - 18 - (down / max_rate) * 76))
            up_points.append((x, height - 18 - (up / max_rate) * 76))
        if len(down_points) > 1:
            canvas.create_line(*[coord for point in down_points for coord in point], fill="#175cd3", width=2, smooth=True)
            canvas.create_line(*[coord for point in up_points for coord in point], fill="#027a48", width=2, smooth=True)
        canvas.create_line(0, height - 18, width, height - 18, fill=self.colors["border"])

    def open_speed_test(self):
        webbrowser.open("https://www.speedtest.net/")
        self.status.set("Opened Speedtest in your browser.")

    def speed_level_color(self, level):
        colors = {
            None: self.colors.get("muted", "#64748b"),
            "Running": "#175cd3",
            "Excellent": "#027a48",
            "Good": "#027a48",
            "Needs Work": "#b54708",
            "Bad": "#c01048",
            "Upgrade Recommended": "#c01048",
            "Error": "#c01048",
        }
        return colors.get(level, self.colors.get("muted", "#64748b"))

    def speed_rating(self, mbps):
        if mbps >= 100:
            return "Excellent", "This is strong for streaming, large downloads, video meetings, and most home or office work."
        if mbps >= 50:
            return "Good", "This should feel good for normal browsing, streaming, cloud apps, and meetings."
        if mbps >= 25:
            return "Needs Work", "This is usable, but large downloads and busy video meetings may feel slower."
        if mbps >= 10:
            return "Bad", "This may struggle with multiple users, video meetings, or large downloads."
        return "Upgrade Recommended", "This is slow. Check Wi-Fi signal, router, ISP plan, or wired connection options."

    def cache_busted_url(self, url):
        separator = "&" if "?" in url else "?"
        return f"{url}{separator}cacheBust={int(time.time() * 1000)}"

    def speed_request_headers(self):
        return {
            "User-Agent": f"{APP_TITLE}/{APP_VERSION}",
            "Cache-Control": "no-cache",
        }

    def update_speed_status(self, text, level="Running", running=None):
        self.last_speed_text = text
        self.last_speed_level = level
        if running is not None:
            self.speed_test_running = running
        speed_label = self.network_widgets.get("speed")
        if speed_label is not None:
            try:
                if speed_label.winfo_exists():
                    speed_label.configure(text=text, fg=self.speed_level_color(level))
            except Exception:
                pass
        speed_button = self.network_widgets.get("speed_button")
        if speed_button is not None:
            try:
                if speed_button.winfo_exists():
                    speed_button.configure(state="disabled" if self.speed_test_running else "normal")
            except Exception:
                pass
        self.status.set(text)

    def queue_speed_status(self, text, level="Running", running=None):
        self.speed_event_queue.put(("status", text, level, running))

    def poll_speed_events(self):
        try:
            while True:
                event = self.speed_event_queue.get_nowait()
                kind = event[0]
                if kind == "status":
                    _kind, text, level, running = event
                    self.update_speed_status(text, level=level, running=running)
                elif kind == "finish":
                    _kind, download_result, upload_result, error, upload_error = event
                    self.finish_speed_test(download_result, upload_result, error=error, upload_error=upload_error)
                self.speed_event_queue.task_done()
        except queue.Empty:
            pass
        except Exception:
            pass
        if self.speed_test_running:
            try:
                self.after(100, self.poll_speed_events)
            except Exception:
                pass

    def run_quick_speed_test(self):
        if self.speed_test_running:
            self.update_speed_status("Speed test: already running. Please wait for the current test to finish.", level="Running", running=True)
            return
        try:
            while True:
                self.speed_event_queue.get_nowait()
                self.speed_event_queue.task_done()
        except queue.Empty:
            pass
        self.update_speed_status("Speed test: starting download and upload checks...", level="Running", running=True)
        self.after(100, self.poll_speed_events)
        thread = threading.Thread(target=self._speed_test_worker, daemon=True)
        thread.start()

    def download_speed_sample(self, source_name, url, target_bytes, max_seconds):
        start = time.time()
        total = 0
        request = urllib.request.Request(self.cache_busted_url(url), headers=self.speed_request_headers())
        with urllib.request.urlopen(request, timeout=max_seconds + 5) as response:
            while total < target_bytes and time.time() - start < max_seconds:
                chunk = response.read(min(256 * 1024, target_bytes - total))
                if not chunk:
                    break
                total += len(chunk)
        elapsed = max(time.time() - start, 0.1)
        if total < 512 * 1024:
            raise RuntimeError(f"{source_name} did not return enough test data")
        return {
            "source": source_name,
            "mbps": (total * 8) / elapsed / 1_000_000,
            "bytes": total,
            "elapsed": elapsed,
        }

    def upload_speed_sample(self):
        data = os.urandom(2_000_000)
        start = time.time()
        headers = self.speed_request_headers()
        headers["Content-Type"] = "application/octet-stream"
        request = urllib.request.Request(self.cache_busted_url(SPEED_TEST_UPLOAD_URL), data=data, method="POST", headers=headers)
        with urllib.request.urlopen(request, timeout=18) as response:
            response.read(1)
        elapsed = max(time.time() - start, 0.1)
        return {
            "source": "Cloudflare",
            "mbps": (len(data) * 8) / elapsed / 1_000_000,
            "bytes": len(data),
            "elapsed": elapsed,
        }

    def _speed_test_worker(self):
        download_result = None
        upload_result = None
        errors = []
        for source_name, url, target_bytes, max_seconds in SPEED_TEST_DOWNLOAD_SOURCES:
            self.queue_speed_status(f"Speed test: checking download with {source_name}...", level="Running", running=True)
            try:
                download_result = self.download_speed_sample(source_name, url, target_bytes, max_seconds)
                break
            except Exception as exc:
                errors.append(f"{source_name}: {exc}")
        if download_result is None:
            message = "; ".join(errors[-2:]) if errors else "No speed-test source responded."
            self.speed_event_queue.put(("finish", None, None, message, None))
            return
        upload_error = None
        self.queue_speed_status("Speed test: checking upload...", level="Running", running=True)
        try:
            upload_result = self.upload_speed_sample()
        except Exception as exc:
            upload_error = str(exc)
        self.speed_event_queue.put(("finish", download_result, upload_result, None, upload_error))

    def finish_speed_test(self, download_result, upload_result=None, error=None, upload_error=None):
        if error or not download_result:
            text = f"Speed test: could not finish. {error or 'The download test did not return a result.'} Use Open Speedtest for a full browser-based test."
            self.update_speed_status(text, level="Error", running=False)
            return
        download_mbps = download_result["mbps"]
        rating, explanation = self.speed_rating(download_mbps)
        upload_text = "Upload: unavailable"
        upload_note = " Upload could not be measured by the quick test." if upload_error else ""
        if upload_result:
            upload_text = f"Upload: {upload_result['mbps']:.1f} Mbps"
        text = (
            f"Speed test: Download: {download_mbps:.1f} Mbps | {upload_text} - {rating}. "
            f"{explanation} Source: {download_result['source']}. "
            f"This quick test is a helpful estimate; Speedtest.net can confirm ping and provider-grade results.{upload_note}"
        )
        self.update_speed_status(text, level=rating, running=False)

    def render_summary_cards(self, parent, items, privacy):
        cards = ttk.Frame(parent, padding=(4, 6, 4, 4))
        cards.pack(fill="x")
        summary_keys = ["Computer Name", "Model", "Operating System", "Security Agent", "Pending Reboot", "C: Free Space"]
        for index, key in enumerate(summary_keys):
            card = tk.Frame(cards, bg=self.colors["card"], highlightbackground=self.colors["border"], highlightthickness=1)
            card.grid(row=index // 3, column=index % 3, sticky="ew", padx=3, pady=3)
            cards.grid_columnconfigure(index % 3, weight=1)
            ttk.Label(card, text=key, style="CardTitle.TLabel").pack(anchor="w", padx=10, pady=(8, 2))
            ttk.Label(card, text=mask_value(key, items.get(key), privacy), style="CardValue.TLabel", wraplength=300).pack(anchor="w", padx=10, pady=(0, 8))

    def tree_mousewheel(self, event, tree):
        try:
            row_count = len(tree.get_children(""))
            visible_rows = int(float(tree.cget("height")))
        except Exception:
            row_count = 0
            visible_rows = 0
        if row_count > visible_rows:
            tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"
        return None

    def table_column_anchor(self, header):
        return "center" if header in ("Category", "Status", "Row") else "w"

    def render_items(self, parent, section, items, privacy):
        box = ttk.Frame(parent, padding=(4, 5, 4, 6))
        box.pack(fill="x")
        tree = ttk.Treeview(box, columns=("Information", "Details"), show="headings", height=min(max(len(items), 5), 11))
        tree.heading("Information", text="Information", anchor="center")
        tree.heading("Details", text="Details", anchor="center")
        tree.column("Information", width=210, anchor="w", stretch=False)
        tree.column("Details", width=790, anchor="w", stretch=True)
        odd, even = self.row_colors()
        tree.tag_configure("odd", background=odd)
        tree.tag_configure("even", background=even)
        for index, (key, value) in enumerate(items.items()):
            tree.insert("", "end", values=(key, mask_value(key, value, privacy)), tags=("even" if index % 2 else "odd",))
        tree.bind("<Double-1>", lambda _event, t=tree: self.copy_selected(t))
        tree.bind("<MouseWheel>", lambda event, t=tree: self.tree_mousewheel(event, t))
        self.attach_menu(tree)
        TreeToolTip(tree, 0)
        tree.pack(fill="x", expand=False)

    def render_table(self, parent, title, rows, privacy, open_by_default=False):
        section = CollapsibleSection(parent, title, open_by_default=open_by_default, colors=self.colors)
        section.pack(fill="x", padx=4, pady=(4, 2))
        box = ttk.Frame(section.body, padding=(0, 0, 0, 6))
        box.pack(fill="x")
        if not rows:
            ttk.Label(box, text="No records found.").pack(anchor="w")
            return
        headers = list(rows[0].keys())
        tree = ttk.Treeview(box, columns=headers, show="headings", height=min(max(len(rows), 3), 11))
        for header in headers:
            tree.heading(header, text=header, anchor="center")
            width = {"Category": 120, "Table": 220, "Information": 210, "Details": 560}.get(header, 155)
            stretch = header in ("Details", "Value", "Message", "Command", "Name", "Description")
            tree.column(header, width=width, anchor=self.table_column_anchor(header), stretch=stretch)
        odd, even = self.row_colors()
        tree.tag_configure("odd", background=odd)
        tree.tag_configure("even", background=even)
        for index, row in enumerate(rows):
            tree.insert("", "end", values=[mask_value(header, row.get(header), privacy) for header in headers], tags=("even" if index % 2 else "odd",))
        tree.bind("<Double-1>", lambda _event, t=tree: self.copy_selected(t))
        self.attach_menu(tree)
        key_index = headers.index("Information") if "Information" in headers else 0
        TreeToolTip(tree, key_index)
        show_horizontal_scroll = len(headers) > 4
        if show_horizontal_scroll:
            x_scroll = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
            tree.configure(xscrollcommand=x_scroll.set)
        tree.bind("<MouseWheel>", lambda event, t=tree: self.tree_mousewheel(event, t))
        box.grid_columnconfigure(0, weight=1)
        box.grid_rowconfigure(0, weight=1)
        tree.grid(row=0, column=0, sticky="nsew")
        if show_horizontal_scroll:
            x_scroll.grid(row=1, column=0, sticky="ew")

    def attach_menu(self, tree):
        menu = Menu(tree, tearoff=0)
        menu.add_command(label="Copy selected row", command=lambda: self.copy_selected(tree))
        tree.bind("<Button-3>", lambda event: menu.tk_popup(event.x_root, event.y_root))

    def copy_selected(self, tree):
        item = tree.focus()
        if not item:
            return
        values = tree.item(item, "values")
        headers = list(tree["columns"])
        row = OrderedDict((headers[index], values[index]) for index in range(min(len(headers), len(values))))
        lines = self.professional_copy_lines("Selected Row", [row])
        self.copy_text("\n".join(lines))
        self.status.set("Selected row copied.")

    def copy_text(self, text):
        try:
            if pyperclip is not None:
                pyperclip.copy(text)
            else:
                self.clipboard_clear()
                self.clipboard_append(text)
        except Exception:
            self.clipboard_clear()
            self.clipboard_append(text)

    def professional_copy_lines(self, title, rows):
        lines = [
            f"{APP_TITLE} {APP_VERSION}",
            title,
            f"Generated: {datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')}",
            "",
            f"{'Category':<22}  {'Information':<34}  Details",
            "-" * 118,
        ]
        for row in rows:
            category = row.get("Category") or row.get("Section") or self.current_view.get()
            info = row.get("Information") or row.get("Field") or row.get("Table") or "Details"
            details = row.get("Details") or row.get("Value") or row.get("Status") or ""
            lines.extend(aligned_detail_lines(category, info, mask_value(info, details, self.privacy.get())))
        return lines

    def helpdesk_rows(self):
        rows = []
        for section, key in HELPDESK_FIELDS:
            bundle = self.report.get(section, {})
            value = bundle.get("items", {}).get(key)
            if value is not None:
                rows.append(OrderedDict([("Category", self.tab_label(section)), ("Information", key), ("Details", value)]))
        return rows

    def current_view_rows(self):
        section = self.current_view.get()
        if self.search_var.get().strip() or section == "Search":
            return self.search_results()
        if section == "Problems":
            return [OrderedDict([("Category", "Problems"), ("Information", level), ("Details", message)]) for level, message in self.health_findings()]
        if section == "Helpdesk":
            return self.helpdesk_rows()
        if section == "Dashboard":
            rows = []
            overview = self.report.get("Overview", {}).get("items", {})
            for key in ("Computer Name", "Model", "Operating System", "Pending Reboot"):
                rows.append(OrderedDict([("Category", "Dashboard"), ("Information", key), ("Details", overview.get(key, "Unavailable"))]))
            for level, message in self.health_findings():
                rows.append(OrderedDict([("Category", "Dashboard"), ("Information", level), ("Details", message)]))
            return rows
        bundle = self.report.get(section, {})
        return [
            OrderedDict([("Category", self.tab_label(section)), ("Information", key), ("Details", value)])
            for key, value in self.visible_items(section, bundle.get("items", {})).items()
        ]

    def copy_all(self):
        if self.search_var.get().strip():
            rows = self.search_results()
            lines = self.professional_copy_lines("Search Results", rows)
            self.copy_text("\n".join(lines))
            self.status.set("Search results copied.")
            return
        rows = self.current_view_rows()
        self.copy_text("\n".join(self.professional_copy_lines(f"{self.tab_label(self.current_view.get())} View", rows)))
        self.status.set("Current view copied.")

    def default_path(self, ext):
        computer = safe_filename(self.report.get("Overview", {}).get("items", {}).get("Computer Name"), "computer")
        folder = Path(self.export_folder.get() or Path.home() / "Documents")
        return folder / f"{OUTPUT_PREFIX}_{computer}_{now_stamp()}.{ext}"

    def report_preview_text(self, ext="all"):
        sections = len(self.report)
        rows = len(flatten_csv_rows(self.report, self.privacy.get()))
        label = "all report formats" if ext == "all" else ext.upper()
        return f"Save {label}?\n\nSections: {sections}\nReport rows: {rows}\nPrivacy mode: {'On' if self.privacy.get() else 'Off'}"

    def remember_export(self, path):
        self.recent_exports.append(str(path))
        self.recent_exports = self.recent_exports[-8:]
        if getattr(self, "recent_label", None) is not None:
            self.recent_label.configure(text=f"Last export: {Path(path).name}")

    def export(self, ext):
        self.ensure_full_report()
        default = self.default_path(ext)
        filetypes = {
            "txt": [("Text files", "*.txt")],
            "csv": [("CSV files", "*.csv")],
            "json": [("JSON files", "*.json")],
            "xlsx": [("Excel files", "*.xlsx")],
            "pdf": [("PDF files", "*.pdf")],
        }[ext]
        path = filedialog.asksaveasfilename(defaultextension=f".{ext}", initialfile=default.name, initialdir=str(default.parent), filetypes=filetypes)
        if not path:
            return
        if not messagebox.askyesno("Export preview", self.report_preview_text(ext)):
            return
        try:
            writers = {"txt": write_text, "csv": write_csv, "json": write_json, "xlsx": write_xlsx, "pdf": write_pdf}
            writers[ext](path, self.report, self.privacy.get())
            self.remember_export(path)
            self.status.set(f"Export saved: {path}")
            messagebox.showinfo("Export saved", f"Saved report:\n{path}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def export_all(self):
        self.ensure_full_report()
        folder = filedialog.askdirectory(initialdir=self.export_folder.get() or str(Path.home() / "Documents"), title="Choose where to save all reports")
        if not folder:
            return
        if not messagebox.askyesno("Export preview", self.report_preview_text("all")):
            return
        try:
            base = Path(folder) / f"{OUTPUT_PREFIX}_{safe_filename(self.report.get('Overview', {}).get('items', {}).get('Computer Name'), 'computer')}_{now_stamp()}"
            write_text(str(base.with_suffix(".txt")), self.report, self.privacy.get())
            write_csv(str(base.with_suffix(".csv")), self.report, self.privacy.get())
            write_json(str(base.with_suffix(".json")), self.report, self.privacy.get())
            write_pdf(str(base.with_suffix(".pdf")), self.report, self.privacy.get())
            self.remember_export(base.with_suffix(".txt"))
            self.remember_export(base.with_suffix(".csv"))
            self.remember_export(base.with_suffix(".json"))
            self.remember_export(base.with_suffix(".pdf"))
            saved = ["Text", "CSV", "JSON", "PDF"]
            if Workbook is not None:
                write_xlsx(str(base.with_suffix(".xlsx")), self.report, self.privacy.get())
                self.remember_export(base.with_suffix(".xlsx"))
                saved.append("Excel")
            self.status.set(f"Saved all reports to: {folder}")
            messagebox.showinfo("Reports saved", f"Saved {', '.join(saved)} reports:\n{folder}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def take_screenshot(self):
        if ImageGrab is None:
            messagebox.showerror("Screenshot unavailable", "Screenshot capture requires Pillow.")
            return
        try:
            self.update_idletasks()
            x = self.winfo_rootx()
            y = self.winfo_rooty()
            w = self.winfo_width()
            h = self.winfo_height()
            image = ImageGrab.grab(bbox=(x, y, x + w, y + h))
            path = self.screenshot_path()
            image.save(path)
            copied = self.copy_image_to_clipboard(image)
            if not copied:
                self.copy_text(str(path))
            self.remember_export(path)
            message = f"Screenshot saved to Pictures and {'copied to clipboard' if copied else 'path copied to clipboard'}: {path.name}"
            self.status.set(message)
        except Exception as exc:
            messagebox.showerror("Screenshot failed", str(exc))

    def screenshot_path(self):
        computer = safe_filename(self.report.get("Overview", {}).get("items", {}).get("Computer Name"), "computer")
        return pictures_folder() / f"{OUTPUT_PREFIX}_{computer}_screenshot_{now_stamp()}.png"

    def copy_image_to_clipboard(self, image):
        try:
            import win32clipboard

            output = io.BytesIO()
            image.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            finally:
                win32clipboard.CloseClipboard()
            return True
        except Exception:
            return False

    def open_output_folder(self):
        try:
            os.startfile(str(Path(self.export_folder.get() or Path.home() / "Documents")))
        except Exception as exc:
            messagebox.showerror("Open folder failed", str(exc))


def run_silent(args):
    report = collect_report()
    out_dir = Path(args.out_dir or Path.cwd())
    out_dir.mkdir(parents=True, exist_ok=True)
    computer = safe_filename(report.get("Overview", {}).get("items", {}).get("Computer Name"), "computer")
    privacy = bool(args.privacy)
    if args.export_text:
        write_text(args.export_text, report, privacy)
    if args.export_csv:
        write_csv(args.export_csv, report, privacy)
    if args.export_json:
        write_json(args.export_json, report, privacy)
    if args.export_xlsx:
        write_xlsx(args.export_xlsx, report, privacy)
    if args.export_pdf:
        write_pdf(args.export_pdf, report, privacy)
    if args.silent and not any([args.export_text, args.export_csv, args.export_json, args.export_xlsx, args.export_pdf]):
        base = out_dir / f"{OUTPUT_PREFIX}_{computer}_{now_stamp()}"
        write_text(str(base.with_suffix(".txt")), report, privacy)
        write_csv(str(base.with_suffix(".csv")), report, privacy)
        write_json(str(base.with_suffix(".json")), report, privacy)
        write_pdf(str(base.with_suffix(".pdf")), report, privacy)
        if Workbook is not None:
            write_xlsx(str(base.with_suffix(".xlsx")), report, privacy)


def parse_args():
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--silent", action="store_true", help="Run without the UI and export default report files.")
    parser.add_argument("--privacy", action="store_true", help="Hide serials, IPs, MAC addresses, users, and similar identifiers in exports.")
    parser.add_argument("--out-dir", help="Folder for --silent default exports.")
    parser.add_argument("--export-text", help="Write a text report to this path.")
    parser.add_argument("--export-csv", help="Write a CSV report to this path.")
    parser.add_argument("--export-json", help="Write a JSON report to this path.")
    parser.add_argument("--export-xlsx", help="Write an Excel report to this path.")
    parser.add_argument("--export-pdf", help="Write a PDF report to this path.")
    return parser.parse_args()


def main():
    args = parse_args()
    if args.silent or any([args.export_text, args.export_csv, args.export_json, args.export_xlsx, args.export_pdf]):
        run_silent(args)
        return
    app = AssetInfoApp()
    app.mainloop()


if __name__ == "__main__":
    main()
