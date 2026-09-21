import ctypes
import hashlib
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.request
import uuid
import webbrowser
import winreg

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import xml.etree.ElementTree as ET

APP_TITLE = "Office Setup Wizard"
APP_VERSION = "2.0.0"

CREATE_NO_WINDOW = 0x08000000
STILL_ACTIVE = 259
SEE_MASK_NOCLOSEPROCESS = 0x00000040

MAS_URL = (
    "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/"
    "f34d025d5102a790c75c839c8d46b672284729a5/MAS/All-In-One-Version-KL/MAS_AIO.cmd"
)
MAS_HASH = "850F979665FB93999ACAE93F4790C1FF8ED2041532060B7966A121C2D29A0BFA"

ODT_VERSION = "20228-20124"

CONFIG_URL = "https://config.office.com/deploymentsettings"

DEBUG_LOG = os.path.join(tempfile.gettempdir(), "oi_debug.log")


def dlog(msg):
    try:
        with open(DEBUG_LOG, "a", encoding="utf-8") as f:
            f.write("%s %s\n" % (time.strftime("%H:%M:%S"), msg))
    except Exception:
        pass


def check_internet(timeout=5):
    """Quick internet connectivity check."""
    try:
        req = urllib.request.Request(
            "https://config.office.com/deploymentsettings",
            method="HEAD",
            headers={"User-Agent": "OfficeSetupWizard/%s" % APP_VERSION})
        urllib.request.urlopen(req, timeout=timeout)
        return True
    except Exception:
        return False


def download_odt_if_needed(log=lambda s: None):
    """Download ODT files if setup.exe is missing from all known locations."""
    setup = find_resource("setup.exe")
    if setup:
        return setup
    log("setup.exe not found locally, attempting to download ODT...")
    odt_url = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    try:
        store = temp_store_dir()
        os.makedirs(store, exist_ok=True)
        data = download_with_retry(odt_url, max_retries=3, timeout=120, delay=10)
        dst = os.path.join(store, "setup.exe")
        with open(dst, "wb") as f:
            f.write(data)
        log("Downloaded setup.exe to: %s (%d bytes)" % (dst, len(data)))
        return dst
    except Exception as exc:
        log("Failed to download ODT: %s" % exc)
        return None


PALETTES = {
    "light": dict(
        bg="#f0f0f0", fg="#1c1c1c", sub="#444444", hint="#555555",
        status="#8a8a8a", link="#1a5fb4", banner="#1a5fb4",
        entry_bg="#ffffff", entry_fg="#1c1c1c",
        trough="#dcdcdc", accent="#1a5fb4",
        btn="#e0e0e0", btn_fg="#000000", btn_hover="#d0d0d0",
        btn_pressed="#c9c9c9", btn_disabled="#ebebeb", btn_fg_disabled="#9a9a9a",
        border="#8a8a8a", ok="#1a7f37"),
    "dark": dict(
        bg="#202124", fg="#e8eaed", sub="#bdc1c6", hint="#9aa0a6",
        status="#9aa0a6", link="#8ab4f8", banner="#8ab4f8",
        entry_bg="#303134", entry_fg="#e8eaed",
        trough="#3c4043", accent="#8ab4f8",
        btn="#3c4043", btn_fg="#ffffff", btn_hover="#4a4e54",
        btn_pressed="#2f3336", btn_disabled="#2a2d30", btn_fg_disabled="#6f737a",
        border="#5f6368", ok="#81c995"),
}

I18N = {
    "tr": {
        "title": "Office Setup Wizard",
        "banner": "Office Setup Wizard",
        "mode_frame": "Kurulum Modu",
        "mode_auto": "Otomatik Kurulum: Office LTSC Professional Plus 2024 — Seçtiğin Dil",
        "mode_custom": "Özel Configuration (kendi XML dosyam)",
        "install_lang": "Kurulum Dili:",
        "install_lang_hint": "Aramak için yazın veya bir dil kodu girin (örn: tr-tr, en-us)",
        "lang_invalid": "Geçerli bir kurulum dili seçin (örn: tr-tr, en-us).",
        "custom_hint1": "Özel configuration kullanmak için şu adresi açın:",
        "custom_hint2": "1) Dilinizi ve kurmak istediğiniz programları seçin",
        "custom_hint3": "2) 'Export' → 'Office Open XML' olarak indirin",
        "custom_hint4": "3) İndirdiğiniz dosyanın yolunu yapıştırın veya 'Gözat...' ile seçin",
        "browse": "Gözat...",
        "theme_dark": "Gece",
        "theme_light": "Aydınlık",
        "theme_auto": "Otomatik",
        "config_edit": "Configuration Ayarlarını Düzenle",
        "config_title": "Configuration Ayarları",
        "config_lang": "Dil",
        "config_edition": "Mimari",
        "config_channel": "Kanal",
        "config_apps": "Kurulacak Uygulamalar",
        "config_more": "Diğer Ayarlar",
        "config_updates": "Güncellemeleri etkinleştir",
        "config_autoactivate": "Otomatik aktifleştir",
        "config_removemsi": "Mevcut MSI sürümünü kaldır",
        "config_silent": "Sessiz kurulum (arayüz gösterme)",
        "config_force": "Uygulamaları kapanmaya zorla",
        "config_shared": "Paylaşılan bilgisayar lisansı",
        "config_device": "Cihaz tabanlı lisanslama",
        "config_scl": "SCL önbellek",
        "config_pin": "Office uygulamalarını görev çubuğuna sabitle",
        "config_migrate": "Mimari değişikliğine izin ver (MigrateArch)",
        "config_visio": "Visio dahil et",
        "config_project": "Project dahil et",
        "config_shortcuts": "Kurulum sonrası masaüstü kısayolları oluştur",
        "config_save": "Kaydet",
        "config_save_as": "Farklı Kaydet...",
        "config_cancel": "İptal Et",
        "config_saved": "Configuration kaydedildi ve kurulumda kullanılacak.",
        "config_saved_as": "Configuration dosyaya kaydedildi.",
        "st_shortcuts": "Masaüstü kısayolları oluşturuluyor...",
        "start": "Kurulumu Başlat",
        "details_show": "Detayları Göster",
        "details_hide": "Detayları Gizle",
        "ask_title": "Office Zaten Kurulu",
        "ask_msg": "Bilgisayarınızda Office zaten kurulu görünüyor. Ne yapmak istersiniz?",
        "ask_reinstall": "Kaldır ve Yeniden Kur",
        "ask_skip": "Atla (Sadece Aktifleştir)",
        "ask_cancel": "İptal",
        "cancelled": "Kurulum iptal edildi.",
        "close_blocked": "Kurulum devam ederken pencere kapatılamaz.\nLütfen kurulumun bitmesini bekleyin.",
        "done_ok": "Kurulum başarıyla tamamlandı!\n\nOffice kuruldu ve aktive edildi.",
        "done_err": "Kurulum tamamlanamadı.",
        "custom_invalid": "Geçersiz configuration dosyası.\n\nconfig.office.com/deploymentsettings "
                          "adresinden 'Office Open XML' formatında indirilen dosyayı seçmelisiniz.",
        "uac_failed": "Yönetici onayı verilmedi veya işlem başlatılamadı.\nProgramı yönetici olarak çalıştırıp tekrar deneyin.",
        "uac_retry": "Yönetici onayı verilmedi. Kurulum için yönetici izni gerekiyor.\nTekrar denensin mi?",
        "res_missing_title": "Eksik Kurulum Dosyası",
        "res_missing_msg": "Gerekli kurulum dosyası (setup.exe) bulunamadı.\n\nsetup.exe, Office Deployment Tool'un parçasıdır ve Office kurulumunu başlatır. Dosya silinmiş ya da dağıtılan sürümde eksik olabilir.\n\nAşağıdan bu dosyaların bulunduğu klasörü seçerek devam edebilirsiniz.",
        "res_missing_where": "Dosyaları şu konumlara da koyabilirsiniz:\n- {TEMP}\n- OfficeSetupWizard.exe'nin yanına\n\nsetup.exe, aka.ms/odt adresinden indirilen 'officedeploymenttool_20131-20090.exe' zip dosyasının içinden çıkarılır.",
        "res_use": "Klasörü Seç ve Devam Et",
        "res_invalid": "Seçilen klasörde setup.exe bulunamadı.\nOffice Deployment Tool dosyalarının bulunduğu klasörü seçin.",
        "helper_failed": "Arka plan işlemi beklenmedik şekilde sonlandı.",
        "st_check": "Yönetici yetkisi isteniyor...",
        "st_prep": "Kurulum hazırlığı yapılıyor...",
        "st_extract": "Kurulum dosyaları çözümleniyor...",
        "st_config": "Configuration dosyası hazırlanıyor...",
        "st_uninst": ["Office kaldırılıyor...",
                      "Mevcut Office kaldırılıyor, lütfen bekleyin...",
                      "Kaldırma işlemi sürüyor..."],
        "st_inst": ["Office dosyaları indiriliyor...",
                    "Gerekli dosyalar indiriliyor...",
                    "Dosyalar kuruluyor, lütfen bekleyin...",
                    "Programlar yapılandırılıyor...",
                    "Kurulum son aşamada, neredeyse tamam..."],
        "st_inst_start": "Office indiriliyor ve kuruluyor... (ilk kurulum birkaç dakika sürebilir)",
        "st_act_dl": "Aktivasyon scripti indiriliyor...",
        "st_act": "Office aktive ediliyor (Ohook)...",
        "st_act_done": "Aktivasyon tamamlandı.",
        "st_cleanup": "Temizlik yapılıyor...",
        "st_final": "Kurulum tamamlandı, son kontroller yapılıyor...",
        "err_installed_hint": "irm https://get.activated.win | iex  komutunu yönetici PowerShell'de çalıştırıp menüden [2] Ohook seçeneğini seçin.",
        "net_check": "İnternet bağlantısı kontrol ediliyor...",
        "net_fail": "İnternet bağlantısı bulunamadı.\nLütfen internet bağlantınızı kontrol edip tekrar deneyin.",
        "odt_dl": "Office Deployment Tool indiriliyor...",
        "odt_dl_ok": "ODT başarıyla indirildi.",
        "odt_dl_fail": "ODT indirilemedi. Lütfen setup.exe dosyasını manuel olarak yerleştirin.",
        "about_title": "Hakkında",
        "about_ver": "Sürüm: %s",
        "about_info": "Office Setup Wizard — Microsoft Office LTSC 2024'ü tek tıkla kurun.",
        "about_mas": "Aktivasyon: Microsoft Activation Scripts (MAS) v3.12",
        "about_odt": "Dağıtım Aracı: Office Deployment Tool (ODT) %s",
    },
    "en": {
        "title": "Office Setup Wizard",
        "banner": "Office Setup Wizard",
        "mode_frame": "Setup Mode",
        "mode_auto": "Automatic Setup: Office LTSC Professional Plus 2024 — Your Language",
        "mode_custom": "Custom Configuration (my own XML file)",
        "install_lang": "Install Language:",
        "install_lang_hint": "Type to search or enter a language code (e.g. en-us, tr-tr)",
        "lang_invalid": "Please select a valid install language (e.g. en-us, tr-tr).",
        "custom_hint1": "To use a custom configuration, open this address:",
        "custom_hint2": "1) Choose your language and the programs you want",
        "custom_hint3": "2) 'Export' → download as 'Office Open XML'",
        "custom_hint4": "3) Paste the path of the downloaded file below or use 'Browse...'",
        "browse": "Browse...",
        "theme_dark": "Night",
        "theme_light": "Light",
        "theme_auto": "Auto",
        "config_edit": "Edit Configuration Settings",
        "config_title": "Configuration Settings",
        "config_lang": "Language",
        "config_edition": "Architecture",
        "config_channel": "Channel",
        "config_apps": "Applications to Install",
        "config_more": "Other Settings",
        "config_updates": "Enable updates",
        "config_autoactivate": "Activate automatically",
        "config_removemsi": "Remove existing MSI version",
        "config_silent": "Silent install (no UI)",
        "config_force": "Force applications to close",
        "config_shared": "Shared computer licensing",
        "config_device": "Device based licensing",
        "config_scl": "SCL cache override",
        "config_pin": "Pin Office apps to taskbar",
        "config_migrate": "Allow architecture migration (MigrateArch)",
        "config_visio": "Include Visio",
        "config_project": "Include Project",
        "config_shortcuts": "Create desktop shortcuts after setup",
        "config_save": "Save",
        "config_save_as": "Save As...",
        "config_cancel": "Cancel",
        "config_saved": "Configuration saved and will be used for setup.",
        "config_saved_as": "Configuration saved to file.",
        "st_shortcuts": "Creating desktop shortcuts...",
        "start": "Start Setup",
        "details_show": "Show Details",
        "details_hide": "Hide Details",
        "ask_title": "Office Already Installed",
        "ask_msg": "Office appears to be already installed on this computer. What would you like to do?",
        "ask_reinstall": "Uninstall & Reinstall",
        "ask_skip": "Skip (Activate Only)",
        "ask_cancel": "Cancel",
        "cancelled": "Setup cancelled.",
        "close_blocked": "The window cannot be closed while setup is running.\nPlease wait for it to finish.",
        "done_ok": "Setup completed successfully!\n\nOffice has been installed and activated.",
        "done_err": "Setup could not be completed.",
        "custom_invalid": "Invalid configuration file.\n\nPlease select a file downloaded in "
                          "'Office Open XML' format from config.office.com/deploymentsettings.",
        "uac_failed": "Administrator approval was not granted or the process could not be started.\nPlease run the program as administrator and try again.",
        "uac_retry": "Administrator approval was not granted. Setup requires administrator rights.\nTry again?",
        "res_missing_title": "Missing Setup File",
        "res_missing_msg": "The required setup file (setup.exe) could not be found.\n\nsetup.exe is part of the Office Deployment Tool and starts the Office installation. It may have been deleted or be missing from the distributed build.\n\nYou can continue by selecting a folder that contains these files below.",
        "res_missing_where": "You can also place the files here:\n- {TEMP}\n- Next to OfficeSetupWizard.exe\n\nsetup.exe is extracted from the 'officedeploymenttool_20131-20090.exe' zip downloaded from aka.ms/odt.",
        "res_use": "Select Folder & Continue",
        "res_invalid": "setup.exe was not found in the selected folder.\nPlease select a folder that contains the Office Deployment Tool files.",
        "helper_failed": "The background process ended unexpectedly.",
        "st_check": "Requesting administrator privileges...",
        "st_prep": "Preparing setup...",
        "st_extract": "Extracting setup files...",
        "st_config": "Preparing configuration file...",
        "st_uninst": ["Uninstalling Office...",
                      "Removing existing Office, please wait...",
                      "Uninstall still in progress..."],
        "st_inst": ["Downloading Office files...",
                    "Downloading required files...",
                    "Installing files, please wait...",
                    "Configuring programs...",
                    "Finalizing setup, almost done..."],
        "st_inst_start": "Downloading and installing Office... (first install may take several minutes)",
        "st_act_dl": "Downloading activation script...",
        "st_act": "Activating Office (Ohook)...",
        "st_act_done": "Activation completed.",
        "st_cleanup": "Cleaning up...",
        "st_final": "Setup finished, running final checks...",
        "err_installed_hint": "Run  irm https://get.activated.win | iex  in an admin PowerShell and choose option [2] Ohook from the menu.",
        "net_check": "Checking internet connection...",
        "net_fail": "No internet connection found.\nPlease check your connection and try again.",
        "odt_dl": "Downloading Office Deployment Tool...",
        "odt_dl_ok": "ODT downloaded successfully.",
        "odt_dl_fail": "ODT download failed. Please place setup.exe manually.",
        "about_title": "About",
        "about_ver": "Version: %s",
        "about_info": "Office Setup Wizard — Install Microsoft Office LTSC 2024 with one click.",
        "about_mas": "Activation: Microsoft Activation Scripts (MAS) v3.12",
        "about_odt": "Deployment Tool: Office Deployment Tool (ODT) %s",
    },
}

LANG_CODES = ["tr", "en"]

OFFICE_APPS = [
    ("Access", "Access"),
    ("Excel", "Excel"),
    ("Skype for Business", "Lync"),
    ("OneDrive Desktop", "OneDrive"),
    ("OneNote", "OneNote"),
    ("Outlook (classic)", "Outlook"),
    ("PowerPoint", "PowerPoint"),
    ("Publisher", "Publisher"),
    ("Word", "Word"),
]

OFFICE_LANGS = [
    ("Türkçe", "tr-tr"),
    ("English (US)", "en-us"),
    ("English (UK)", "en-gb"),
    ("Deutsch", "de-de"),
    ("Français", "fr-fr"),
    ("Français (Canada)", "fr-ca"),
    ("Español", "es-es"),
    ("Español (México)", "es-mx"),
    ("Italiano", "it-it"),
    ("Português (Brasil)", "pt-br"),
    ("Português (Portugal)", "pt-pt"),
    ("Nederlands", "nl-nl"),
    ("Polski", "pl-pl"),
    ("Русский", "ru-ru"),
    ("Українська", "uk-ua"),
    ("العربية", "ar-sa"),
    ("עברית", "he-il"),
    ("हिन्दी", "hi-in"),
    ("Indonesia", "id-id"),
    ("Melayu", "ms-my"),
    ("ไทย", "th-th"),
    ("Tiếng Việt", "vi-vn"),
    ("中文(简体)", "zh-cn"),
    ("中文(繁體)", "zh-tw"),
    ("日本語", "ja-jp"),
    ("한국어", "ko-kr"),
    ("Čeština", "cs-cz"),
    ("Dansk", "da-dk"),
    ("Eesti", "et-ee"),
    ("Suomi", "fi-fi"),
    ("Ελληνικά", "el-gr"),
    ("Magyar", "hu-hu"),
    ("Latviešu", "lv-lv"),
    ("Lietuvių", "lt-lt"),
    ("Norsk Bokmål", "nb-no"),
    ("Română", "ro-ro"),
    ("Slovenčina", "sk-sk"),
    ("Slovenščina", "sl-si"),
    ("Srpski", "sr-latn-cs"),
    ("Svenska", "sv-se"),
]

OFFICE_CHANNELS = ["PerpetualVL2024", "PerpetualVL2021", "Current",
                   "MonthlyEnterprise", "SemiAnnualEnterprise"]

DEFAULT_APPS = {"Excel", "PowerPoint", "Word"}


def resolve_lang_code(raw):
    low = (raw or "").strip().lower()
    if not low:
        return None
    for d, c in OFFICE_LANGS:
        if d.lower() == low or c.lower() == low:
            return c
    m = re.match(r"^[a-z]{2}(?:-[a-z0-9]{2})?$", low)
    if m:
        code = m.group(0)
        if "-" in code:
            return code
        for d, c in OFFICE_LANGS:
            if c.split("-")[0] == code:
                return c
    return None


def lang_label(item):
    d, c = item
    return "%s — %s" % (c, d)


def lang_label_for_code(code):
    for d, c in OFFICE_LANGS:
        if c == code:
            return lang_label((d, c))
    return code


def lang_code_from_label(label):
    for d, c in OFFICE_LANGS:
        if lang_label((d, c)) == label:
            return c
    return None


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def temp_store_dir():
    return os.path.join(tempfile.gettempdir(), "OfficeSetupWizard")


def res_candidates(name):
    out = []
    base = os.path.dirname(sys.executable) if is_frozen() else \
        os.path.dirname(os.path.abspath(__file__))
    alt = os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), name)
    for p in (os.path.join(temp_store_dir(), name),
              os.path.join(base, name),
              alt,
              resource_path(name)):
        if p not in out:
            out.append(p)
    return out


def find_resource(name, extra_dir=None):
    if extra_dir:
        p = os.path.join(extra_dir, name)
        if os.path.isfile(p):
            return p
    for p in res_candidates(name):
        if os.path.isfile(p):
            return p
    return None


def ensure_temp_store():
    store = temp_store_dir()
    try:
        os.makedirs(os.path.join(store, "Configuration"), exist_ok=True)
    except OSError:
        return store
    for rel in ("Configuration/Configuration.xml", "configuration-Office365-x64.xml",
                "setup.exe", "officedeploymenttool_20131-20090.exe"):
        src = resource_path(rel)
        if not os.path.isfile(src):
            continue
        dst = os.path.join(store, rel)
        try:
            if not os.path.isfile(dst) or \
                    os.path.getsize(dst) != os.path.getsize(src):
                shutil.copyfile(src, dst)
        except OSError:
            pass
    return store


def config_res_path(name, extra_dir=None):
    found = find_resource(name, extra_dir)
    return found if found else resource_path(name)


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def detect_language():
    try:
        primary = ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0x3FF
        if primary == 0x1F:
            return "tr"
    except Exception:
        pass
    return "en"


def detect_theme():
    try:
        with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            val, _ = winreg.QueryValueEx(k, "AppsUseLightTheme")
            return "light" if val else "dark"
    except OSError:
        return "light"


def office_installed():
    try:
        with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration") as k:
            winreg.QueryValueEx(k, "ProductReleaseIds")
            return True
    except OSError:
        pass
    for p in (r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
              r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"):
        if os.path.isfile(p):
            return True
    return False


def open_url(url):
    try:
        os.startfile(url)
    except Exception:
        webbrowser.open(url)


def download_with_retry(url, max_retries=3, timeout=60, delay=5):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    last_err = None
    for attempt in range(max_retries):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except Exception as exc:
            last_err = exc
            if attempt < max_retries - 1:
                time.sleep(delay * (attempt + 1))
    raise last_err


def validate_config_xml(xml_content):
    try:
        root = ET.fromstring(xml_content)
        if root.tag != "Configuration":
            return False
        add = root.find("Add")
        if add is None:
            return False
        if not add.get("OfficeClientEdition"):
            return False
        products = add.findall("Product")
        if not products:
            return False
        for prod in products:
            if not prod.get("ID"):
                return False
            langs = prod.findall("Language")
            if not langs:
                return False
        return True
    except ET.ParseError:
        return False


def oem_decode(data):
    try:
        cp = ctypes.windll.kernel32.GetOEMCP()
        return data.decode("cp%d" % cp, errors="replace")
    except Exception:
        return data.decode("utf-8", errors="replace")


def set_preferred_app_mode(allow_dark=True):
    try:
        ux = ctypes.windll.uxtheme
        fn = getattr(ux, "SetPreferredAppMode", None)
        if fn is not None:
            fn(1 if allow_dark else 3)
    except Exception:
        pass


def apply_dark_titlebar(hwnd, dark):
    try:
        user32 = ctypes.windll.user32
        top = user32.GetAncestor(hwnd, 2) or hwnd
        dwm = ctypes.windll.dwmapi
        dwm.DwmSetWindowAttribute.argtypes = [
            ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_uint]
        mode = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):
            if dwm.DwmSetWindowAttribute(top, attr, ctypes.byref(mode),
                                         ctypes.sizeof(mode)) == 0:
                break
        try:
            caption = ctypes.c_ulong(0x00242120 if dark else 0x00FFFFFF)
            text = ctypes.c_ulong(0x00EDEAE8 if dark else 0x001C1C1C)
            border = ctypes.c_ulong(0x001B1B1B if dark else 0x00FFFFFF)
            for attr, val in ((35, caption), (36, text), (34, border)):
                dwm.DwmSetWindowAttribute(top, attr, ctypes.byref(val), 4)
        except Exception:
            pass
    except Exception:
        pass


_mutex_handle = None


def acquire_single_instance():
    global _mutex_handle
    try:
        _mutex_handle = ctypes.windll.kernel32.CreateMutexW(
            None, False, "Local\\OfficeSetupWizard_SingleInstance")
        if ctypes.windll.kernel32.GetLastError() == 183:
            ctypes.windll.kernel32.CloseHandle(_mutex_handle)
            _mutex_handle = None
            return False
        return True
    except Exception:
        return True


def focus_existing():
    try:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        psapi = ctypes.windll.psapi
        found = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def cb(hwnd, _):
            if not user32.IsWindowVisible(hwnd):
                return True
            pid = ctypes.c_ulong()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value == os.getpid():
                return True
            hproc = kernel32.OpenProcess(0x1000, False, pid.value)
            if hproc:
                buf = ctypes.create_unicode_buffer(260)
                if psapi.GetModuleBaseNameW(hproc, None, buf, 260) and \
                        buf.value.lower().startswith("officesetupwizard"):
                    found.append(hwnd)
                kernel32.CloseHandle(hproc)
            return True

        user32.EnumWindows(cb, None)
        if found:
            user32.ShowWindow(found[0], 9)
            user32.SetForegroundWindow(found[0])
            return True
    except Exception:
        pass
    return False


class SHELLEXECUTEINFOW(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("fMask", ctypes.c_ulong),
        ("hwnd", ctypes.c_void_p),
        ("lpVerb", ctypes.c_wchar_p),
        ("lpFile", ctypes.c_wchar_p),
        ("lpParameters", ctypes.c_wchar_p),
        ("lpDirectory", ctypes.c_wchar_p),
        ("nShow", ctypes.c_int),
        ("hInstApp", ctypes.c_void_p),
        ("lpIDList", ctypes.c_void_p),
        ("lpClass", ctypes.c_wchar_p),
        ("hwndClass", ctypes.c_void_p),
        ("hKey", ctypes.c_void_p),
        ("dwHotKey", ctypes.c_ulong),
        ("hIconOrMonitor", ctypes.c_void_p),
        ("hProcess", ctypes.c_void_p),
    ]


def spawn_elevated(params):
    try:
        dlog("spawn_elevated: file=%s params=%s" % (sys.executable, params[:200]))
        sei = SHELLEXECUTEINFOW()
        sei.cbSize = ctypes.sizeof(SHELLEXECUTEINFOW)
        sei.fMask = SEE_MASK_NOCLOSEPROCESS
        sei.lpVerb = "runas"
        sei.lpFile = sys.executable
        sei.lpParameters = params
        sei.nShow = 0
        if not ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(sei)):
            err = ctypes.windll.kernel32.GetLastError()
            dlog("spawn_elevated: ShellExecuteExW failed, error=%d" % err)
            return None
        dlog("spawn_elevated: success, hProcess=%s" % sei.hProcess)
        return sei.hProcess
    except Exception as exc:
        dlog("spawn_elevated: exception=%s" % exc)
        return None


def relaunch_elevated():
    try:
        sei = SHELLEXECUTEINFOW()
        sei.cbSize = ctypes.sizeof(SHELLEXECUTEINFOW)
        sei.lpVerb = "runas"
        sei.lpFile = sys.executable
        sei.lpParameters = "--no-elevate"
        sei.nShow = 5
        if not ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(sei)):
            return False
        ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(sei.hProcess))
        return True
    except Exception:
        return False


def is_frozen():
    return bool(getattr(sys, "frozen", False))


UNINSTALL_XML = (
    '<Configuration>\n'
    '  <Remove All="TRUE" />\n'
    '  <Display Level="None" AcceptEULA="TRUE" />\n'
    '</Configuration>\n')


def build_config_xml(settings):
    lang = settings["lang"]
    edition = settings["edition"]
    channel = settings["channel"]
    excluded = [aid for aid, inc in settings["apps"].items() if not inc]

    def product(pid, pkey):
        lines = ['  <Product ID="%s" PIDKEY="%s">' % (pid, pkey)]
        lines.append('    <Language ID="%s" />' % lang)
        for aid in excluded:
            lines.append('    <ExcludeApp ID="%s" />' % aid)
        lines.append('  </Product>')
        return lines

    out = ['<Configuration ID="%s">' % uuid.uuid4()]
    out.append('  <Add OfficeClientEdition="%s" Channel="%s">' % (edition, channel))
    out += product("ProPlus2024Volume", "XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB")
    if settings["include_visio"]:
        out += product("VisioPro2024Volume", "B7TN8-FJ8V3-7QYCP-HQPMV-YY89G")
    if settings["include_project"]:
        out += product("ProjectPro2024Volume", "FQQ23-N4YCY-73HQ3-FM9WC-76HF4")
    out.append('  </Add>')
    out.append('  <Property Name="SharedComputerLicensing" Value="%d" />'
               % (1 if settings["shared"] else 0))
    out.append('  <Property Name="FORCEAPPSHUTDOWN" Value="%s" />'
               % ("TRUE" if settings["force"] else "FALSE"))
    out.append('  <Property Name="DeviceBasedLicensing" Value="%d" />'
               % (1 if settings["device"] else 0))
    out.append('  <Property Name="SCLCacheOverride" Value="%d" />'
               % (1 if settings["scl"] else 0))
    out.append('  <Property Name="AUTOACTIVATE" Value="%d" />'
               % (1 if settings["autoactivate"] else 0))
    out.append('  <Property Name="PINNEDICONS" Value="%s" />'
               % ("TRUE" if settings.get("pin") else "FALSE"))
    out.append('  <Property Name="MigrateArch" Value="%s" />'
               % ("TRUE" if settings.get("migrate") else "FALSE"))
    out.append('  <Updates Enabled="%s" />' % ("TRUE" if settings["updates"] else "FALSE"))
    if settings["removemsi"]:
        out.append('  <RemoveMSI />')
    if settings["silent"]:
        out.append('  <Display Level="None" AcceptEULA="TRUE" />')
    out.append('  <AppSettings>')
    out.append('    <User Key="software\\microsoft\\office\\16.0\\excel\\options" '
               'Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" '
               'Id="L_SaveExcelfilesas" />')
    out.append('    <User Key="software\\microsoft\\office\\16.0\\powerpoint\\options" '
               'Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" '
               'Id="L_SavePowerPointfilesas" />')
    out.append('    <User Key="software\\microsoft\\office\\16.0\\word\\options" '
               'Name="defaultformat" Value="" Type="REG_SZ" App="word16" '
               'Id="L_SaveWordfilesas" />')
    out.append('  </AppSettings>')
    out.append('</Configuration>')
    return "\n".join(out) + "\n"


def create_lnk(target, lnk):
    try:
        ps = ("$w = New-Object -ComObject WScript.Shell; "
              "$s = $w.CreateShortcut('%s'); "
              "$s.TargetPath = '%s'; "
              "$s.WorkingDirectory = '%s'; "
              "$s.IconLocation = '%s'; "
              "$s.Save()"
              % (lnk.replace("'", "''"), target.replace("'", "''"),
                 os.path.dirname(target).replace("'", "''"),
                 target.replace("'", "''")))
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       creationflags=CREATE_NO_WINDOW,
                       capture_output=True, timeout=60)
        return os.path.isfile(lnk)
    except Exception:
        return False


def create_office_shortcuts(log=lambda s: None):
    try:
        office_dirs = [r"C:\Program Files\Microsoft Office\root\Office16",
                       r"C:\Program Files (x86)\Microsoft Office\root\Office16"]
        app_map = [("WINWORD.EXE", "Word"), ("EXCEL.EXE", "Excel"),
                   ("POWERPNT.EXE", "PowerPoint"), ("OUTLOOK.EXE", "Outlook"),
                   ("ONENOTE.EXE", "OneNote"), ("ACCESS.EXE", "Access"),
                   ("MSPUB.EXE", "Publisher"), ("LYNC.EXE", "Skype for Business"),
                   ("VISIO.EXE", "Visio"), ("WINPROJ.EXE", "Project")]
        desktop = os.path.join(os.environ.get("PUBLIC", r"C:\Users\Public"),
                               "Desktop")
        if not os.path.isdir(desktop):
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        created = []
        for exe, name in app_map:
            target = None
            for d in office_dirs:
                p = os.path.join(d, exe)
                if os.path.isfile(p):
                    target = p
                    break
            if not target:
                continue
            if create_lnk(target, os.path.join(desktop, name + ".lnk")):
                created.append(name)
        if created:
            log("Desktop shortcuts created: " + ", ".join(created))
    except Exception as exc:
        log("Shortcut error: %s" % exc)


class ConfigEditor:
    def __init__(self, app):
        self.app = app
        self.t = app.t
        self.pal = app.pal
        self.win = tk.Toplevel(app.root)
        self.win.title(self.t("config_title"))
        self.win.configure(bg=self.pal["bg"])
        self.win.resizable(False, False)
        self.win.transient(app.root)
        try:
            self.win.iconbitmap(resource_path("office_setup_wizard.ico"))
        except Exception:
            pass

        mode = app.mode.get()
        default_lang = lang_label_for_code(app.selected_install_lang())
        self.lang_var = tk.StringVar(value=default_lang)
        self.edition_var = tk.StringVar(value="64")
        self.channel_var = tk.StringVar(value="PerpetualVL2024")
        self.app_vars = {}
        for human, aid in OFFICE_APPS:
            self.app_vars[aid] = tk.BooleanVar(value=(aid in DEFAULT_APPS))
        self.visio_var = tk.BooleanVar(value=True)
        self.project_var = tk.BooleanVar(value=True)
        self.updates_var = tk.BooleanVar(value=False)
        self.autoact_var = tk.BooleanVar(value=True)
        self.removemsi_var = tk.BooleanVar(value=True)
        self.silent_var = tk.BooleanVar(value=True)
        self.force_var = tk.BooleanVar(value=False)
        self.shared_var = tk.BooleanVar(value=False)
        self.device_var = tk.BooleanVar(value=False)
        self.scl_var = tk.BooleanVar(value=False)
        self.pin_var = tk.BooleanVar(value=False)
        self.migrate_var = tk.BooleanVar(value=False)
        self.shortcuts_var = tk.BooleanVar(
            value=getattr(app, "shortcuts_enabled", True))

        self._build()
        self._load_config()
        self.win.update_idletasks()
        w = max(480, self.win.winfo_reqwidth())
        h = self.win.winfo_reqheight()
        x = app.root.winfo_rootx() + (app.root.winfo_width() - w) // 2
        y = max(10, app.root.winfo_rooty() + 30)
        self.win.geometry("%dx%d+%d+%d" % (w, h, x, y))
        apply_dark_titlebar(self.win.winfo_id(), app.theme == "dark")
        self.win.grab_set()

    def _load_config(self):
        path = None
        is_master = False
        if self.app.edited_config_path and os.path.isfile(self.app.edited_config_path):
            path = self.app.edited_config_path
        elif self.app.mode.get() == "custom":
            p = self.app.custom_path.get().strip()
            if p and os.path.isfile(p):
                path = p
        else:
            path = config_res_path("Configuration/Configuration.xml")
            is_master = True
        if not path or not os.path.isfile(path):
            return
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                root = ET.fromstring(f.read())
        except Exception:
            return
        add = root.find("Add")
        if add is None:
            return
        self.edition_var.set(add.get("OfficeClientEdition", "64"))
        self.channel_var.set(add.get("Channel", "PerpetualVL2024"))
        self.visio_var.set(False)
        self.project_var.set(False)
        for prod in add.findall("Product"):
            pid = prod.get("ID", "")
            if "Visio" in pid:
                self.visio_var.set(True)
            elif "Project" in pid:
                self.project_var.set(True)
            for lang in prod.findall("Language"):
                lid = lang.get("ID")
                if lid:
                    self.lang_var.set(lang_label_for_code(lid))
                    break
        excluded = set()
        for prod in add.findall("Product"):
            for ex in prod.findall("ExcludeApp"):
                excluded.add(ex.get("ID"))
        for human, aid in OFFICE_APPS:
            self.app_vars[aid].set(aid not in excluded)
        props = {}
        for prop in root.findall("Property"):
            props[prop.get("Name")] = prop.get("Value")
        self.shared_var.set(props.get("SharedComputerLicensing", "0") == "1")
        self.force_var.set(props.get("FORCEAPPSHUTDOWN", "FALSE").upper() == "TRUE")
        self.device_var.set(props.get("DeviceBasedLicensing", "0") == "1")
        self.scl_var.set(props.get("SCLCacheOverride", "0") == "1")
        self.autoact_var.set(props.get("AUTOACTIVATE", "1") == "1")
        self.pin_var.set(props.get("PINNEDICONS", "FALSE").upper() == "TRUE")
        self.migrate_var.set(props.get("MigrateArch", "FALSE").upper() == "TRUE")
        up = root.find("Updates")
        if up is not None:
            self.updates_var.set(str(up.get("Enabled", "TRUE")).upper() == "TRUE")
        self.removemsi_var.set(root.find("RemoveMSI") is not None)
        dsp = root.find("Display")
        self.silent_var.set(dsp is not None and
                            str(dsp.get("Level", "")).upper() == "NONE")
        if is_master:
            self.lang_var.set(lang_label_for_code(self.app.selected_install_lang()))

    def _build(self):
        outer = ttk.Frame(self.win, padding=12)
        outer.pack(fill="both", expand=True)

        row1 = ttk.Frame(outer)
        row1.pack(fill="x", pady=(0, 8))
        ttk.Label(row1, text=self.t("config_lang")).pack(side="left")
        self.lang_cb = ttk.Combobox(
            row1, textvariable=self.lang_var, state="readonly", width=30,
            values=[lang_label(i) for i in OFFICE_LANGS])
        self.lang_cb.pack(side="left", padx=(8, 0))

        row2 = ttk.Frame(outer)
        row2.pack(fill="x", pady=(0, 8))
        ttk.Label(row2, text=self.t("config_edition")).pack(side="left")
        for label, val in (("64-bit", "64"), ("32-bit", "32")):
            ttk.Radiobutton(row2, text=label, value=val,
                            variable=self.edition_var).pack(side="left", padx=(8, 0))
        ttk.Label(row2, text=self.t("config_channel")).pack(side="left", padx=(16, 0))
        ttk.Combobox(row2, textvariable=self.channel_var, state="readonly",
                     width=20, values=OFFICE_CHANNELS).pack(side="left", padx=(8, 0))

        af = ttk.LabelFrame(outer, text=self.t("config_apps"), padding=8)
        af.pack(fill="x", pady=(0, 8))
        for i, (human, aid) in enumerate(OFFICE_APPS):
            ttk.Checkbutton(af, text=human,
                            variable=self.app_vars[aid]).grid(
                row=i // 3, column=i % 3, sticky="w", padx=(0, 12), pady=2)
        ttk.Checkbutton(af, text=self.t("config_visio"),
                        variable=self.visio_var).grid(row=3, column=0, sticky="w", pady=2)
        ttk.Checkbutton(af, text=self.t("config_project"),
                        variable=self.project_var).grid(row=3, column=1, sticky="w", pady=2)

        mf = ttk.LabelFrame(outer, text=self.t("config_more"), padding=8)
        mf.pack(fill="x", pady=(0, 8))
        more = [
            ("updates_var", self.t("config_updates")),
            ("autoact_var", self.t("config_autoactivate")),
            ("removemsi_var", self.t("config_removemsi")),
            ("silent_var", self.t("config_silent")),
            ("force_var", self.t("config_force")),
            ("shared_var", self.t("config_shared")),
            ("device_var", self.t("config_device")),
            ("scl_var", self.t("config_scl")),
            ("pin_var", self.t("config_pin")),
            ("migrate_var", self.t("config_migrate")),
            ("shortcuts_var", self.t("config_shortcuts")),
        ]
        for i, (attr, txt) in enumerate(more):
            ttk.Checkbutton(mf, text=txt,
                            variable=getattr(self, attr)).grid(
                row=i // 2, column=i % 2, sticky="w", padx=(0, 12), pady=2)

        bf = ttk.Frame(outer)
        bf.pack(fill="x", pady=(8, 0))
        ttk.Button(bf, text=self.t("config_save"),
                   command=self._save).pack(side="left", padx=(0, 8))
        ttk.Button(bf, text=self.t("config_save_as"),
                   command=self._save_as).pack(side="left", padx=(0, 8))
        ttk.Button(bf, text=self.t("config_cancel"),
                   command=self.win.destroy).pack(side="left")

    def _settings(self):
        apps = {}
        for human, aid in OFFICE_APPS:
            apps[aid] = self.app_vars[aid].get()
        return {
            "lang": lang_code_from_label(self.lang_var.get()) or "en-us",
            "edition": self.edition_var.get(),
            "channel": self.channel_var.get() or "PerpetualVL2024",
            "apps": apps,
            "include_visio": self.visio_var.get(),
            "include_project": self.project_var.get(),
            "updates": self.updates_var.get(),
            "autoactivate": self.autoact_var.get(),
            "removemsi": self.removemsi_var.get(),
            "silent": self.silent_var.get(),
            "force": self.force_var.get(),
            "shared": self.shared_var.get(),
            "device": self.device_var.get(),
            "scl": self.scl_var.get(),
            "pin": self.pin_var.get(),
            "migrate": self.migrate_var.get(),
        }

    def _save(self):
        path = os.path.join(tempfile.gettempdir(), "office_config.xml")
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(build_config_xml(self._settings()))
        except Exception as exc:
            messagebox.showerror(self.t("config_title"), "%s" % exc, parent=self.win)
            return
        self.app.edited_config_path = path
        self.app.shortcuts_enabled = self.shortcuts_var.get()
        self.app.status_var.set(self.t("config_saved"))
        self.win.destroy()

    def _save_as(self):
        path = filedialog.asksaveasfilename(
            parent=self.win, title=self.t("config_title"),
            defaultextension=".xml",
            filetypes=[("XML dosyası", "*.xml"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(build_config_xml(self._settings()))
        except Exception as exc:
            messagebox.showerror(self.t("config_title"), "%s" % exc, parent=self.win)
            return
        self.app.edited_config_path = path
        self.app.shortcuts_enabled = self.shortcuts_var.get()
        self.app.status_var.set(self.t("config_saved_as"))
        self.win.destroy()


def ps_argv(cmd, *args):
    all_args = (cmd,) + args
    cmd_str = " ".join('"%s"' % a.replace('"', '`"') for a in all_args)
    return ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-Command", "& { %s }; exit $LASTEXITCODE" % cmd_str]


class Installer:
    def __init__(self, post, log, t):
        self.post = post
        self.log = log
        self.t = t

    def run(self, mode, custom_path, decision, config_path=None, shortcuts=True,
            lang=None, resource_dir=None):
        workdir = tempfile.mkdtemp(prefix="OfficeInstaller_")
        try:
            self._run_steps(mode, custom_path, decision, config_path, workdir,
                            shortcuts, lang, resource_dir)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)

    def _run_steps(self, mode, custom_path, decision, config_path, workdir,
                   shortcuts=True, lang=None, resource_dir=None):
        dlog("steps begin, decision=%s" % decision)
        self.post(self.t("st_prep"), 2)
        self.log(self.t("st_prep"))

        setup_src = find_resource("setup.exe", resource_dir)
        if not setup_src:
            raise RuntimeError(
                "setup.exe could not be found. Restore the Office Deployment "
                "Tool files (see the recovery dialog) and try again.")
        self.post(self.t("st_extract"), 5)
        self.log(self.t("st_extract"))
        setup_dst = os.path.join(workdir, "setup.exe")
        shutil.copyfile(setup_src, setup_dst)

        cfg_path = None
        if decision != "skip":
            self.post(self.t("st_config"), 8)
            self.log(self.t("st_config"))
            if mode == "custom":
                with open(custom_path, "r", encoding="utf-8-sig") as f:
                    xml = f.read()
            elif config_path and os.path.isfile(config_path):
                with open(config_path, "r", encoding="utf-8-sig") as f:
                    xml = f.read()
            else:
                with open(config_res_path("Configuration/Configuration.xml",
                                          resource_dir),
                          "r", encoding="utf-8-sig") as f:
                    xml = f.read()
                lang_sel = lang or ("tr-tr" if detect_language() == "tr"
                                    else "en-us")
                xml = re.sub(r'<Language ID="[^"]*"\s*/>',
                             '<Language ID="%s" />' % lang_sel, xml)

            if not validate_config_xml(xml):
                raise RuntimeError(self.t("custom_invalid"))

            if "<Display " not in xml:
                xml = re.sub(
                    r"(</Configuration>\s*)$",
                    '  <Display Level="None" AcceptEULA="TRUE" />\n</Configuration>',
                    xml, count=1)

            cfg_path = os.path.join(workdir, "configuration.xml")
            try:
                with open(cfg_path, "w", encoding="utf-8") as f:
                    f.write(xml)
                if not os.path.isfile(cfg_path) or os.path.getsize(cfg_path) == 0:
                    raise RuntimeError("Failed to write configuration.xml")
                self.log("> Configuration written to: %s (%d bytes)" % (cfg_path, os.path.getsize(cfg_path)))
            except Exception as exc:
                raise RuntimeError("Failed to create configuration.xml: %s" % exc)

        if decision == "reinstall":
            uninst = os.path.join(workdir, "uninstall.xml")
            with open(uninst, "w", encoding="utf-8") as f:
                f.write(UNINSTALL_XML)
            self.log(self.t("st_uninst")[0])
            code = self._run_ticked(
                ps_argv(setup_dst, "/configure", "uninstall.xml"), workdir,
                15, 40, self.t("st_uninst"))
            if code != 0:
                raise RuntimeError("Office removal failed (code: %d)." % code)

        if decision != "skip":
            inst_start = 45 if decision == "reinstall" else 12
            self.post(self.t("st_inst_start"), inst_start)
            self.log(self.t("st_inst_start"))
            self.log("> setup.exe: %s" % setup_dst)
            if not os.path.isfile(setup_dst):
                raise RuntimeError("setup.exe not found at: %s" % setup_dst)
            if not cfg_path or not os.path.isfile(cfg_path):
                raise RuntimeError("configuration.xml not found at: %s" % (cfg_path or "None"))
            self.log("> configuration.xml: %s" % cfg_path)
            code = self._run_ticked(
                ps_argv(setup_dst, "/configure", cfg_path),
                workdir, inst_start, 78, self.t("st_inst"))
            if code != 0:
                hints = []
                if code == 4294967295:
                    hints.append("This error often means the setup process could not start.")
                    hints.append("Check if antivirus is blocking the installation.")
                    hints.append("Try running the application as administrator.")
                elif code == 3010:
                    hints.append("Installation completed but requires a restart.")
                elif code == 1706:
                    hints.append("No network connection available. Check your internet.")
                hint_text = "\n".join(hints) if hints else ""
                raise RuntimeError(
                    "Office install failed (code: %d).\n\nPossible causes:\n"
                    "- No internet connection\n"
                    "- Antivirus blocking the installation\n"
                    "- Insufficient disk space\n"
                    "%s\n"
                    "Check your internet connection and try again." % (code, hint_text))

        self.post(self.t("st_act_dl"), 85)
        self._activate()
        self.post(self.t("st_act_done"), 96)
        if shortcuts:
            self.post(self.t("st_shortcuts"), 97)
            create_office_shortcuts(self.log)
        self.post(self.t("st_cleanup"), 98)

    def _run_ticked(self, args, cwd, start, cap, msgs, env=None):
        msg = msgs[0] if isinstance(msgs, list) else str(msgs)
        self.post(msg, start)
        self.log("> %s" % " ".join(args))
        self.log("> cwd: %s" % cwd)
        try:
            proc = subprocess.Popen(
                args, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                env=env, creationflags=CREATE_NO_WINDOW)
        except OSError as exc:
            self.log("> failed to start process: %s" % exc)
            raise RuntimeError("Could not start process: %s" % exc)
        stop = threading.Event()

        expected = 120 if "uninstall" in " ".join(args).lower() else 600
        begun = time.time()
        shown = [start]

        def scale(pct):
            frac = max(0.0, min(1.0, pct / 100.0))
            return min(cap, max(start, start + int((cap - start) * frac)))

        def emit(pct):
            if pct > shown[0]:
                shown[0] = pct
                self.post(msg, pct)

        def reader():
            last_tick = [0.0]
            for raw in iter(lambda: proc.stdout.readline(), b""):
                if stop.is_set():
                    break
                line = oem_decode(raw).rstrip("\r\n")
                if line.strip():
                    self.log(line)
                    m = re.search(r"(?:^|[^0-9])(\d{1,3})\s*%", line)
                    if m:
                        emit(scale(int(m.group(1))))
                now = time.time()
                if now - last_tick[0] >= 1.0:
                    last_tick[0] = now
                    frac = min(1.0, (now - begun) / expected)
                    target = start + int((cap - start - 2) * frac)
                    emit(target)

        threading.Thread(target=reader, daemon=True).start()
        try:
            code = proc.wait()
            shown[0] = cap
            self.post(msg, cap)
            self.log("> exit code: %d" % code)
            return code
        finally:
            stop.set()

    def _activate(self):
        self.post(self.t("st_act_dl"), 87)
        try:
            data = download_with_retry(MAS_URL, max_retries=3, timeout=60, delay=5)
        except Exception as exc:
            raise RuntimeError(
                "Activation script could not be downloaded: %s\n\nOffice is "
                "installed but not activated. You can activate it manually:\n%s"
                % (exc, self.t("err_installed_hint")))

        digest = hashlib.sha256(data).hexdigest().upper()
        if digest != MAS_HASH:
            raise RuntimeError("Activation script failed verification (hash mismatch).")

        script = os.path.join(tempfile.gettempdir(), "MAS_%s.cmd" % uuid.uuid4().hex)
        with open(script, "wb") as f:
            f.write(data)
        self.log("> %s (%d bytes)" % (os.path.basename(script), len(data)))
        try:
            self.post(self.t("st_act"), 90)
            env = dict(os.environ)
            for key in ("WT_SESSION", "WT_PROFILE_ID", "TERM_PROGRAM",
                        "TERM_PROGRAM_VERSION", "TERMINAL_EMULATOR"):
                env.pop(key, None)
            inner = "cmd /c '%s' -el -qedit /Ohook" % script.replace("'", "''")
            cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                   "-Command", "& " + inner + "; exit $LASTEXITCODE"]
            proc = subprocess.Popen(
                cmd, cwd=tempfile.gettempdir(),
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                env=env, creationflags=CREATE_NO_WINDOW)

            def reader():
                for raw in iter(lambda: proc.stdout.readline(), b""):
                    line = oem_decode(raw).rstrip("\r\n")
                    if line.strip():
                        self.log(line)

            threading.Thread(target=reader, daemon=True).start()
            try:
                proc.wait(timeout=900)
                self.log("> activation exit code: %d" % proc.returncode)
            except subprocess.TimeoutExpired:
                proc.kill()
                raise RuntimeError("Activation timed out.")
        finally:
            try:
                os.remove(script)
            except OSError:
                pass


def _write_json(logfile, obj):
    with open(logfile, "a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")


def run_helper_mode():
    try:
        args = sys.argv[sys.argv.index("--helper") + 1:]
        mode, custom, decision, config_path, shortcuts, lang, resource_dir, \
            logfile = (args[0], args[1], args[2], args[3], args[4], args[5],
                       args[6], args[7])
        sys_lang = detect_language()

        def t(key):
            return I18N[sys_lang].get(key, key)

        def post(msg, pct):
            _write_json(logfile, {"t": "status", "m": msg, "p": int(pct)})

        def log(line):
            _write_json(logfile, {"t": "log", "m": line})

        installer = Installer(post, log, t)
        installer.run(mode, custom, decision, config_path,
                      shortcuts.lower() in ("1", "true", "yes"), lang or None,
                      resource_dir or None)
        _write_json(logfile, {"t": "done", "ok": True,
                              "m": I18N[sys_lang]["done_ok"]})
    except Exception as exc:
        traceback.print_exc()
        try:
            lang = detect_language()
            _write_json(logfile, {"t": "done", "ok": False,
                                  "m": I18N[lang]["done_err"] + "\n\n%s" % exc})
        except Exception:
            pass


class App:
    def __init__(self, root):
        dlog("App init")
        self.root = root
        self.root.title("%s v%s" % (APP_TITLE, APP_VERSION))
        self.root.resizable(False, False)
        try:
            self.root.iconbitmap(resource_path("office_setup_wizard.ico"))
        except Exception:
            pass

        self.queue = queue.Queue()
        self.mode = tk.StringVar(value="auto")
        self.custom_path = tk.StringVar()
        self.edited_config_path = None
        self.shortcuts_enabled = True
        self.resource_dir = None

        self.lang = detect_language()
        self.theme_mode = "auto"
        self.theme = detect_theme()
        self.running = False
        self.details_on = False
        self.container = None
        self.log_text = None

        self.status_var = tk.StringVar(value="")
        self.pct_var = tk.StringVar(value="%0")

        self.install_lang_var = tk.StringVar(
            value="tr-tr" if self.lang == "tr" else "en-us")

        ensure_temp_store()

        self._rebuild()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(120, self._poll_queue)
        dlog("App ready, lang=%s theme=%s" % (self.lang, self.theme))

    def t(self, key):
        return I18N[self.lang].get(key, key)

    @property
    def pal(self):
        return PALETTES[self.theme]

    def _resolve_install_lang(self):
        return resolve_lang_code(self.install_lang_var.get())

    def selected_install_lang(self):
        return self._resolve_install_lang() or \
            self.install_lang_var.get().strip() or "en-us"

    def _filter_install_langs(self):
        txt = self.install_lang_var.get().strip().lower()
        if txt:
            def rank(item):
                d, c = item
                dl, cl = d.lower(), c.lower()
                if cl.startswith(txt):
                    return 0
                if dl.startswith(txt):
                    return 1
                if txt in cl:
                    return 2
                if txt in dl:
                    return 3
                return 4

            items = sorted(OFFICE_LANGS, key=rank)
        else:
            items = OFFICE_LANGS
        try:
            self.install_lang_cb.configure(values=[lang_label(i) for i in items])
        except Exception:
            pass

    def _on_install_lang_key(self, event):
        self._filter_install_langs()

    def _on_install_lang_select(self, event):
        code = lang_code_from_label(self.install_lang_cb.get())
        if code:
            self.install_lang_var.set(code)

    def _rebuild(self):
        pos = None
        if self.container is not None:
            try:
                pos = (self.root.winfo_x(), self.root.winfo_y())
            except Exception:
                pos = None
            self.container.destroy()
        self.container = ttk.Frame(self.root, padding=14)
        self.container.pack(fill="both", expand=True)
        self._build_header(self.container)
        self._build_body(self.container)
        self._apply_theme()
        self._layout()
        if pos is None:
            self._center()
        else:
            self.root.update_idletasks()
            w = max(440, self.root.winfo_reqwidth())
            h = self.root.winfo_reqheight()
            self.root.geometry("%dx%d+%d+%d" % (w, h, pos[0], pos[1]))
        self.root.update_idletasks()
        self._apply_titlebar()

    def _apply_titlebar(self):
        apply_dark_titlebar(self.root.winfo_id(), self.theme == "dark")
        self.root.after(100,
                        lambda: apply_dark_titlebar(self.root.winfo_id(),
                                                    self.theme == "dark"))

    def _load_icon(self, variant, name):
        try:
            return tk.PhotoImage(
                file=resource_path(os.path.join("icons", variant, name + ".png")))
        except Exception:
            return None

    def _build_header(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(0, 6))
        tk.Label(row, text=self.t("banner"),
                 font=("Segoe UI", 14, "bold"),
                 bg=self.pal["bg"], fg=self.pal["banner"]).pack(side="left")

        ctrl = ttk.Frame(row)
        ctrl.pack(side="right")

        self.icons = {}
        for variant in ("light", "dark"):
            self.icons[variant] = {}
            for name in ("moon", "sun", "auto"):
                self.icons[variant][name] = self._load_icon(variant, name)

        self.theme_menu = tk.Menu(ctrl, tearoff=0)
        for i, (mode, key) in enumerate((("dark", "theme_dark"),
                                         ("light", "theme_light"),
                                         ("auto", "theme_auto"))):
            icon_name = {"dark": "moon", "light": "sun", "auto": "auto"}[mode]
            self.theme_menu.add_command(
                label=self.t(key), image=self.icons[self.theme][icon_name],
                compound="left",
                command=lambda m=mode: self.root.after(
                    0, lambda: self._set_theme(m)))
            if mode == self.theme_mode:
                self.theme_menu.entryconfig(i, foreground=self.pal["accent"])

        icon_name = {"dark": "moon", "light": "sun", "auto": "auto"}[self.theme_mode]
        self.theme_btn = tk.Button(
            ctrl, image=self.icons[self.theme][icon_name], bd=1, relief="raised",
            cursor="hand2", width=36, height=30, command=self._pop_theme_menu)
        self.theme_btn.pack(side="left", padx=(0, 8))

        target_lang = "en" if self.lang == "tr" else "tr"
        self.lang_btn = ttk.Button(ctrl, text=target_lang.upper(), width=5,
                                    command=self._toggle_lang)
        self.lang_btn.pack(side="left", padx=(0, 8))

        about_btn = ttk.Button(ctrl, text="?", width=3,
                                command=self._show_about)
        about_btn.pack(side="left")

    def _pop_theme_menu(self):
        try:
            x = self.theme_btn.winfo_rootx()
            y = self.theme_btn.winfo_rooty() + self.theme_btn.winfo_height()
            self.theme_menu.tk_popup(x, y)
        finally:
            try:
                self.theme_menu.grab_release()
            except Exception:
                pass

    def _build_body(self, parent):
        mode_frame = ttk.LabelFrame(parent, text=self.t("mode_frame"), padding=10)
        mode_frame.pack(fill="x", pady=(0, 10))
        ttk.Radiobutton(mode_frame, text=self.t("mode_auto"),
                        variable=self.mode, value="auto",
                        command=self._on_mode_change).pack(anchor="w")

        lang_row = ttk.Frame(mode_frame)
        lang_row.pack(fill="x", anchor="w", pady=(4, 0), padx=(22, 0))
        ttk.Label(lang_row, text=self.t("install_lang")).pack(side="left")
        self.install_lang_cb = ttk.Combobox(
            lang_row, textvariable=self.install_lang_var, width=32,
            values=[lang_label(i) for i in OFFICE_LANGS])
        self.install_lang_cb.pack(side="left", padx=(8, 0))
        self.install_lang_cb.configure(postcommand=self._filter_install_langs)
        self.install_lang_cb.bind("<KeyRelease>", self._on_install_lang_key)
        self.install_lang_cb.bind("<<ComboboxSelected>>",
                                  self._on_install_lang_select)
        tk.Label(mode_frame, text=self.t("install_lang_hint"),
                 font=("Segoe UI", 8), fg=self.pal["hint"], bg=self.pal["bg"]
                 ).pack(anchor="w", padx=(22, 0), pady=(2, 0))

        ttk.Radiobutton(mode_frame, text=self.t("mode_custom"),
                        variable=self.mode, value="custom",
                        command=self._on_mode_change).pack(anchor="w", pady=(8, 0))

        self.custom_frame = ttk.Frame(parent)
        tk.Label(self.custom_frame, text=self.t("custom_hint1"),
                 font=("Segoe UI", 8), fg=self.pal["hint"], bg=self.pal["bg"]
                 ).pack(anchor="w", pady=(2, 2))
        link = tk.Label(self.custom_frame, text=CONFIG_URL,
                        font=("Segoe UI", 8, "underline"), cursor="hand2",
                        fg=self.pal["link"], bg=self.pal["bg"])
        link.pack(anchor="w", pady=(0, 2))
        link.bind("<Button-1>", lambda e: open_url(CONFIG_URL))
        tk.Label(self.custom_frame, text=self.t("custom_hint2"),
                 font=("Segoe UI", 8), fg=self.pal["hint"], bg=self.pal["bg"]
                 ).pack(anchor="w")
        tk.Label(self.custom_frame, text=self.t("custom_hint3"),
                 font=("Segoe UI", 8), fg=self.pal["hint"], bg=self.pal["bg"]
                 ).pack(anchor="w")
        tk.Label(self.custom_frame, text=self.t("custom_hint4"),
                 font=("Segoe UI", 8), fg=self.pal["hint"], bg=self.pal["bg"]
                 ).pack(anchor="w", pady=(0, 6))

        self.custom_row = ttk.Frame(self.custom_frame)
        self.custom_row.pack(fill="x", pady=(0, 10))
        self.entry = ttk.Entry(self.custom_row, textvariable=self.custom_path)
        self.entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.browse_btn = ttk.Button(
            self.custom_row, text=self.t("browse"), command=self._browse)
        self.browse_btn.pack(side="left")

        self.separator = ttk.Separator(parent)
        self.separator.pack(fill="x", pady=(0, 12))

        self.config_btn = ttk.Button(parent, text=self.t("config_edit"),
                                     command=self._open_config_editor)
        self.config_btn.pack(fill="x", pady=(0, 8))

        self.start_btn = ttk.Button(parent, text=self.t("start"), command=self._start)
        self.start_btn.pack(fill="x", pady=(0, 12))

        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill="x")
        self.progress = ttk.Progressbar(progress_frame, mode="determinate",
                                        maximum=100, value=0)
        self.progress.pack(side="left", fill="x", expand=True)
        tk.Label(progress_frame, textvariable=self.pct_var,
                 font=("Segoe UI", 9, "bold"), width=5,
                 bg=self.pal["bg"], fg=self.pal["fg"]).pack(side="left", padx=(8, 0))

        self.status_lbl = tk.Label(parent, textvariable=self.status_var,
                                   font=("Segoe UI", 8), fg=self.pal["status"],
                                   bg=self.pal["bg"], justify="left",
                                   wraplength=520, anchor="w")
        self.status_lbl.pack(fill="x", pady=(8, 0))

        self.details_btn = ttk.Button(
            parent, text=self.t("details_show"), command=self._toggle_details)
        self.details_btn.pack(anchor="w", pady=(6, 0))

        self.log_frame = ttk.Frame(parent)
        self.log_text = tk.Text(self.log_frame, height=12, wrap="word",
                                state="disabled",
                                bg=self.pal["entry_bg"], fg=self.pal["entry_fg"],
                                insertbackground=self.pal["entry_fg"],
                                selectbackground=self.pal["accent"],
                                selectforeground=self.pal["btn_fg"],
                                highlightbackground=self.pal["border"],
                                highlightcolor=self.pal["border"],
                                relief="solid", borderwidth=1,
                                font=("Consolas", 8))
        scroll = ttk.Scrollbar(self.log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)

    def _apply_theme(self):
        self.root.configure(bg=self.pal["bg"])
        try:
            self.root.option_clear()
        except Exception:
            pass
        self.root.option_add("*TCombobox*Listbox.background", self.pal["entry_bg"])
        self.root.option_add("*TCombobox*Listbox.foreground", self.pal["entry_fg"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", self.pal["accent"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", self.pal["btn_fg"])
        self.root.option_add("*TEntry*Listbox.background", self.pal["entry_bg"])

        style = ttk.Style(self.root)
        style.theme_use("clam")
        p = self.pal
        style.configure(".", background=p["bg"], foreground=p["fg"])
        style.configure("TFrame", background=p["bg"])
        style.configure("TLabel", background=p["bg"], foreground=p["fg"])
        style.configure("TLabelframe", background=p["bg"], foreground=p["fg"],
                        bordercolor=p["border"])
        style.configure("TLabelframe.Label", background=p["bg"], foreground=p["fg"])
        style.configure("TButton", background=p["btn"], foreground=p["btn_fg"],
                        bordercolor=p["border"], lightcolor=p["btn"],
                        darkcolor=p["btn"], padding=(10, 4), focuscolor=p["accent"])
        style.map("TButton",
                  background=[("pressed", p["btn_pressed"]),
                              ("active", p["btn_hover"]),
                              ("disabled", p["btn_disabled"])],
                  foreground=[("disabled", p["btn_fg_disabled"])])
        style.configure("TRadiobutton", background=p["bg"], foreground=p["fg"],
                        indicatorcolor=p["entry_bg"], bordercolor=p["border"],
                        lightcolor=p["bg"], darkcolor=p["bg"])
        style.map("TRadiobutton",
                  background=[("active", p["bg"])],
                  foreground=[("active", p["fg"])],
                  indicatorcolor=[("selected", p["accent"]),
                                  ("active", p["entry_bg"])])
        style.configure("TCheckbutton", background=p["bg"], foreground=p["fg"],
                        indicatorcolor=p["entry_bg"], bordercolor=p["border"],
                        lightcolor=p["bg"], darkcolor=p["bg"])
        style.map("TCheckbutton",
                  background=[("active", p["bg"])],
                  foreground=[("active", p["fg"])],
                  indicatorcolor=[("selected", p["accent"]),
                                  ("active", p["entry_bg"])])
        style.configure("TCombobox", fieldbackground=p["entry_bg"],
                        background=p["btn"], foreground=p["entry_fg"],
                        arrowcolor=p["btn_fg"], bordercolor=p["border"],
                        lightcolor=p["btn"], darkcolor=p["btn"])
        style.map("TCombobox",
                  fieldbackground=[("readonly", p["entry_bg"]),
                                   ("disabled", p["bg"])],
                  foreground=[("readonly", p["entry_fg"]),
                              ("disabled", p["btn_fg_disabled"])],
                  selectbackground=[("readonly", p["entry_bg"])],
                  selectforeground=[("readonly", p["entry_fg"])],
                  background=[("active", p["btn_hover"])],
                  arrowcolor=[("disabled", p["btn_fg_disabled"])])
        style.configure("TEntry", fieldbackground=p["entry_bg"],
                        foreground=p["entry_fg"], insertcolor=p["entry_fg"],
                        bordercolor=p["border"], lightcolor=p["border"],
                        darkcolor=p["border"])
        style.map("TEntry",
                  fieldbackground=[("disabled", p["bg"])],
                  foreground=[("disabled", p["fg"])])
        style.configure("Horizontal.TProgressbar",
                        troughcolor=p["trough"], background=p["accent"],
                        bordercolor=p["bg"], lightcolor=p["accent"],
                        darkcolor=p["accent"])
        style.configure("TSeparator", background=p["border"])
        style.configure("TScrollbar", background=p["btn"],
                        troughcolor=p["trough"], bordercolor=p["bg"],
                        lightcolor=p["btn"], darkcolor=p["btn"],
                        arrowcolor=p["btn_fg"], relief="flat")
        style.map("TScrollbar",
                  background=[("active", p["btn_hover"]),
                              ("pressed", p["btn_pressed"])],
                  arrowcolor=[("disabled", p["btn_fg_disabled"])])
        if hasattr(self, "theme_btn"):
            self.theme_btn.configure(
                bg=p["btn"], fg=p["btn_fg"],
                activebackground=p["btn_hover"], activeforeground=p["btn_fg"],
                highlightthickness=0)
        if hasattr(self, "theme_menu"):
            self.theme_menu.configure(
                bg=p["entry_bg"], fg=p["fg"],
                activebackground=p["accent"], activeforeground=p["btn_fg"],
                borderwidth=1, relief="flat")

    def _set_theme(self, mode):
        self.theme_mode = mode
        self.theme = detect_theme() if mode == "auto" else mode
        self._rebuild()

    def _toggle_lang(self):
        self.lang = "en" if self.lang == "tr" else "tr"
        self._rebuild()

    def _toggle_details(self):
        if self.log_text is None:
            return
        self.details_on = not self.details_on
        if self.details_on:
            self.log_frame.pack(fill="both", expand=True, pady=(6, 0))
            self.details_btn.configure(text=self.t("details_hide"))
        else:
            self.log_frame.pack_forget()
            self.details_btn.configure(text=self.t("details_show"))
        self._center()

    def _on_mode_change(self):
        self._layout()
        self._center()

    def _open_config_editor(self):
        if self.running:
            return
        ConfigEditor(self)

    def _layout(self):
        if self.mode.get() == "custom":
            self.custom_frame.pack(fill="x", pady=(0, 10),
                                   before=self.separator)
            self.entry.configure(state="normal")
            self.browse_btn.configure(state="normal")
        else:
            self.custom_frame.pack_forget()
        if hasattr(self, "install_lang_cb"):
            self.install_lang_cb.configure(
                state="disabled" if self.mode.get() == "custom" else "normal")

    def _browse(self):
        path = filedialog.askopenfilename(
            title=self.t("mode_custom"),
            filetypes=[("XML dosyası", "*.xml"), ("All files", "*.*")])
        if path:
            self.custom_path.set(path)

    def _center(self):
        self.root.update_idletasks()
        w = max(440, self.root.winfo_reqwidth())
        h = self.root.winfo_reqheight()
        x = (self.root.winfo_screenwidth() - w) // 2
        y = max(10, (self.root.winfo_screenheight() - h) // 3)
        self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))

    def _ask_office_installed(self):
        win = tk.Toplevel(self.root)
        win.title(self.t("ask_title"))
        win.configure(bg=self.pal["bg"])
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        result = {"v": None}

        tk.Label(win, text=self.t("ask_msg"),
                 bg=self.pal["bg"], fg=self.pal["fg"],
                 justify="left", wraplength=380
                 ).pack(padx=18, pady=(16, 10))

        btns = ttk.Frame(win)
        btns.pack(pady=(4, 16))

        def choose(v):
            result["v"] = v
            win.destroy()

        ttk.Button(btns, text=self.t("ask_reinstall"),
                   command=lambda: choose("reinstall")).pack(side="left", padx=4)
        ttk.Button(btns, text=self.t("ask_skip"),
                   command=lambda: choose("skip")).pack(side="left", padx=4)
        ttk.Button(btns, text=self.t("ask_cancel"),
                   command=lambda: choose("cancel")).pack(side="left", padx=4)

        win.update_idletasks()
        w = max(480, win.winfo_reqwidth())
        h = win.winfo_reqheight()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - w) // 2
        y = self.root.winfo_rooty() + 40
        win.geometry("%dx%d+%d+%d" % (w, h, x, y))
        self.root.wait_window(win)
        return result["v"]

    def _start(self):
        if self.running:
            return
        mode = self.mode.get()
        custom = self.custom_path.get().strip()
        lang = None

        if mode == "custom":
            if not custom or not os.path.isfile(custom):
                messagebox.showerror(APP_TITLE, self.t("custom_invalid"))
                return
            try:
                with open(custom, "r", encoding="utf-8-sig") as f:
                    txt = f.read()
            except Exception as exc:
                messagebox.showerror(APP_TITLE, "%s" % exc)
                return
            if "<Configuration" not in txt:
                messagebox.showerror(APP_TITLE, self.t("custom_invalid"))
                return
        else:
            lang = self._resolve_install_lang()
            if lang is None:
                messagebox.showerror(APP_TITLE, self.t("lang_invalid"))
                return

        self.status_var.set(self.t("net_check"))
        self.root.update_idletasks()
        if not check_internet():
            messagebox.showerror(APP_TITLE, self.t("net_fail"))
            self.status_var.set("")
            return

        if not self._ensure_setup_resources():
            return

        if office_installed():
            decision = self._ask_office_installed()
            if decision == "cancel":
                messagebox.showinfo(APP_TITLE, self.t("cancelled"))
                return
        else:
            decision = "install"

        self.running = True
        self.start_btn.configure(state="disabled")
        self._layout()
        self.status_var.set("")
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.pct_var.set("%0")

        config_path = self.edited_config_path if mode != "custom" else None
        shortcuts = "1" if self.shortcuts_enabled else "0"
        resource_dir = self.resource_dir

        if is_admin():
            threading.Thread(target=self._worker,
                             args=(mode, custom, decision, config_path,
                                   self.shortcuts_enabled, lang, resource_dir),
                             daemon=True).start()
        else:
            self.status_var.set(self.t("st_check"))
            self._start_helper(mode, custom, decision, config_path, shortcuts,
                               lang, resource_dir)

    def _ensure_setup_resources(self):
        if find_resource("setup.exe", self.resource_dir):
            return True
        self.status_var.set(self.t("odt_dl"))
        self.root.update_idletasks()
        dl = download_odt_if_needed(self._append_log)
        if dl:
            self.status_var.set(self.t("odt_dl_ok"))
            return True
        return self._ask_resource_folder()

    def _ask_resource_folder(self):
        result = {"dir": None}
        win = tk.Toplevel(self.root)
        win.title(self.t("res_missing_title"))
        win.configure(bg=self.pal["bg"])
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        path_var = tk.StringVar()
        where = self.t("res_missing_where").replace(
            "{TEMP}", temp_store_dir())

        tk.Label(win, text=self.t("res_missing_msg"), justify="left",
                 wraplength=440, bg=self.pal["bg"], fg=self.pal["fg"]
                 ).pack(padx=18, pady=(14, 8), anchor="w")
        tk.Label(win, text=where, justify="left", wraplength=440,
                 font=("Segoe UI", 8), fg=self.pal["hint"], bg=self.pal["bg"]
                 ).pack(padx=18, pady=(0, 8), anchor="w")

        row = ttk.Frame(win)
        row.pack(fill="x", padx=18)
        entry = ttk.Entry(row, textvariable=path_var)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 6))

        def pick():
            d = filedialog.askdirectory(parent=win,
                                        title=self.t("res_missing_title"))
            if d:
                path_var.set(d)

        ttk.Button(row, text=self.t("browse"), command=pick).pack(side="left")

        def use():
            d = path_var.get().strip()
            if d and os.path.isdir(d) and find_resource("setup.exe", d):
                result["dir"] = d
                win.destroy()
            else:
                messagebox.showerror(APP_TITLE, self.t("res_invalid"),
                                     parent=win)

        btns = ttk.Frame(win)
        btns.pack(pady=(10, 14))
        ttk.Button(btns, text=self.t("res_use"),
                   command=use).pack(side="left", padx=4)
        ttk.Button(btns, text=self.t("ask_cancel"),
                   command=win.destroy).pack(side="left", padx=4)

        win.update_idletasks()
        w = max(480, win.winfo_reqwidth())
        h = win.winfo_reqheight()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - w) // 2
        y = max(10, self.root.winfo_rooty() + 30)
        win.geometry("%dx%d+%d+%d" % (w, h, x, y))
        self.root.wait_window(win)
        if result["dir"]:
            self.resource_dir = result["dir"]
            return True
        return False

    def _start_helper(self, mode, custom, decision, config_path, shortcuts,
                      lang=None, resource_dir=None):
        logfile = os.path.join(tempfile.gettempdir(),
                               "officeinst_%s.log" % uuid.uuid4().hex)
        try:
            open(logfile, "w", encoding="utf-8").close()
        except OSError:
            pass

        if is_frozen():
            params = subprocess.list2cmdline(
                ["--helper", mode, custom or "", decision, config_path or "",
                 shortcuts, lang or "", resource_dir or "", logfile])
        else:
            params = subprocess.list2cmdline(
                [os.path.abspath(__file__), "--helper", mode,
                 custom or "", decision, config_path or "", shortcuts,
                 lang or "", resource_dir or "", logfile])

        hproc = None
        for _ in range(3):
            hproc = spawn_elevated(params)
            if hproc:
                break
            if not messagebox.askyesno(APP_TITLE, self.t("uac_retry")):
                break
        if not hproc:
            self.running = False
            self.start_btn.configure(state="normal")
            self.status_var.set("")
            messagebox.showerror(APP_TITLE, self.t("uac_failed"))
            return

        threading.Thread(target=self._tail_helper,
                         args=(logfile, hproc), daemon=True).start()

    def _tail_helper(self, logfile, hproc):
        last = 0
        start = time.time()
        done = False
        while not done:
            time.sleep(0.25)
            lines = []
            try:
                with open(logfile, "r", encoding="utf-8") as f:
                    f.seek(last)
                    lines = f.readlines()
                    last = f.tell()
            except Exception:
                pass
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                try:
                    obj = json.loads(line)
                except Exception:
                    continue
                kind = obj.get("t")
                if kind == "status":
                    self.queue.put(("status", obj.get("m", ""), obj.get("p", 0)))
                elif kind == "log":
                    self.queue.put(("log", obj.get("m", "")))
                elif kind == "done":
                    ok = bool(obj.get("ok"))
                    self.queue.put(("done",
                                    "info" if ok else "error",
                                    obj.get("m", "")))
                    done = True

            if not done:
                code = ctypes.c_ulong()
                ok = ctypes.windll.kernel32.GetExitCodeProcess(
                    ctypes.c_void_p(hproc), ctypes.byref(code))
                if ok and code.value != STILL_ACTIVE:
                    if time.time() - start < 4:
                        time.sleep(0.4)
                        continue
                    self.queue.put(("done", "error", self.t("helper_failed")))
                    done = True
                elif time.time() - start > 3600:
                    self.queue.put(("done", "error", self.t("helper_failed")))
                    done = True

        ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(hproc))

    def _worker(self, mode, custom, decision, config_path, shortcuts=True,
                lang=None, resource_dir=None):
        dlog("worker start mode=%s decision=%s" % (mode, decision))
        try:
            installer = Installer(self.post,
                                  lambda line: self.queue.put(("log", line)),
                                  self.t)
            installer.run(mode, custom, decision, config_path, shortcuts, lang,
                          resource_dir)
            self.post(self.t("st_final"), 99)
            self.queue.put(("done", "info", self.t("done_ok")))
        except Exception as exc:
            dlog("worker error: %s" % exc)
            traceback.print_exc()
            self.queue.put(("done", "error",
                            self.t("done_err") + "\n\n%s" % exc))

    def post(self, msg, pct):
        self.queue.put(("status", msg, int(pct)))

    def _append_log(self, line):
        if self.log_text is None:
            return
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n")
        try:
            lines = int(self.log_text.index("end-1c").split(".")[0])
            if lines > 3000:
                self.log_text.delete("1.0", "1500.0")
        except Exception:
            pass
        self.log_text.configure(state="disabled")
        self.log_text.see("end")

    def _poll_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                if item[0] == "status":
                    _, msg, pct = item
                    self.status_var.set(msg)
                    if pct == -1:
                        continue
                    self.progress.stop()
                    self.progress.configure(mode="determinate",
                                            maximum=100, value=pct)
                    self.pct_var.set("%d%%" % pct)
                elif item[0] == "log":
                    self._append_log(item[1])
                elif item[0] == "done":
                    _, kind, text = item
                    self.running = False
                    self.start_btn.configure(state="normal")
                    self._layout()
                    if kind == "error":
                        self.progress.stop()
                        self.progress.configure(mode="determinate", value=100)
                        self.pct_var.set("100%")
                        messagebox.showerror(APP_TITLE, text)
                    else:
                        self.status_var.set(text)
                        self.progress.stop()
                        self.progress.configure(mode="determinate", value=100)
                        self.pct_var.set("100%")
                        if hasattr(self, "status_lbl"):
                            self.status_lbl.configure(fg=self.pal["ok"])
                        self._play_done_sound()
                        self.root.after(3000, self.root.destroy)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_queue)

    def _play_done_sound(self):
        try:
            import winsound
            winsound.PlaySound("SystemAsterisk",
                               winsound.SND_ALIAS | winsound.SND_ASYNC)
        except Exception:
            try:
                ctypes.windll.kernel32.Beep(880, 250)
            except Exception:
                pass

    def _show_about(self):
        win = tk.Toplevel(self.root)
        win.title(self.t("about_title"))
        win.configure(bg=self.pal["bg"])
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text=APP_TITLE,
                 font=("Segoe UI", 14, "bold"),
                 bg=self.pal["bg"], fg=self.pal["banner"]).pack(padx=20, pady=(16, 4))
        tk.Label(win, text=self.t("about_ver") % APP_VERSION,
                 font=("Segoe UI", 10),
                 bg=self.pal["bg"], fg=self.pal["fg"]).pack(padx=20)
        tk.Label(win, text=self.t("about_info"),
                 font=("Segoe UI", 9), wraplength=360, justify="center",
                 bg=self.pal["bg"], fg=self.pal["sub"]).pack(padx=20, pady=(8, 4))
        tk.Label(win, text=self.t("about_mas"),
                 font=("Segoe UI", 9), wraplength=360, justify="center",
                 bg=self.pal["bg"], fg=self.pal["sub"]).pack(padx=20, pady=(2, 2))
        tk.Label(win, text=self.t("about_odt") % ODT_VERSION,
                 font=("Segoe UI", 9), wraplength=360, justify="center",
                 bg=self.pal["bg"], fg=self.pal["sub"]).pack(padx=20, pady=(2, 4))

        link = tk.Label(win, text="github.com/TARIKELER-TARNAK/OfficeSetupWizard",
                        font=("Segoe UI", 8, "underline"), cursor="hand2",
                        fg=self.pal["link"], bg=self.pal["bg"])
        link.pack(padx=20, pady=(4, 8))
        link.bind("<Button-1>", lambda e: open_url(
            "https://github.com/TARIKELER-TARNAK/OfficeSetupWizard"))

        ttk.Button(win, text="OK", command=win.destroy).pack(pady=(0, 12))

        win.update_idletasks()
        w = max(420, win.winfo_reqwidth())
        h = win.winfo_reqheight()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - w) // 2
        y = max(10, self.root.winfo_rooty() + 40)
        win.geometry("%dx%d+%d+%d" % (w, h, x, y))
        apply_dark_titlebar(win.winfo_id(), self.theme == "dark")

    def _on_close(self):
        if self.running:
            messagebox.showinfo(APP_TITLE, self.t("close_blocked"))
            return
        self.root.destroy()


def main():
    if "--helper" in sys.argv:
        run_helper_mode()
        return
    dlog("main start")
    set_preferred_app_mode(True)
    if not is_admin() and "--no-elevate" not in sys.argv:
        dlog("requesting elevation")
        if relaunch_elevated():
            dlog("elevation requested, exiting")
            sys.exit(0)
        dlog("elevation declined, continuing non-elevated")
    if not acquire_single_instance():
        dlog("another instance running, focusing existing window")
        focus_existing()
        sys.exit(0)
    root = tk.Tk()
    App(root)
    dlog("mainloop enter")
    root.mainloop()


if __name__ == "__main__":
    main()
