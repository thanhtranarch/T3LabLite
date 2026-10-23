# -*- coding: utf-8 -*-
"""Split button nhớ nút dùng gần nhất qua các lần mở Revit — phần IronPython.

Một `.splitbutton` của pyRevit đặt IsSynchronizedWithCurrentItem = True: nút
đứng trên cùng là nút vừa bấm. Nhưng Revit chỉ nhớ trong MỘT phiên — mở lại
Revit, pyRevit dựng lại ribbon và nút đầu tiên trong `layout:` lại đứng trên
cùng. Ví dụ: `PDFMarkup.splitbutton` (Bluebeam Status · T3Lab PDF Markup).

Module này giữ lựa chọn đó qua các phiên, cho MỌI `.splitbutton` của extension
(`.splitpushbutton` thì cố ý đứng yên nút đầu, không đụng tới):

  1. Revit đóng (ApplicationClosing) -> remember(): ghi tên nút đang đứng trên
     cùng của từng split button vào %APPDATA%\\T3LabAI\\ribbon_state.json.
  2. Revit mở -> startup.py -> install(). startup.py chạy TRƯỚC khi pyRevit
     dựng ribbon (pyRevit 6.5.5: chạy startup script rồi mới tạo UI), nên việc
     đặt lại nút hoãn tới các lần Idling đầu tiên -> restore().
  3. pyRevit Reload: ribbon cũ vẫn còn và đang giữ lựa chọn của phiên này, nên
     install() ghi nó xuống trước, rồi Idling đặt lại đúng như vậy.

Revit crash thì ApplicationClosing không chạy: lần mở sau dùng lựa chọn đã ghi
ở lần đóng bình thường gần nhất.

Chạy trên IronPython 2.7 / 3.4 (startup.py) và CPython (test): không f-string,
không cú pháp chỉ có ở Python 3. Không hàm nào ném lỗi ra ngoài: startup.py
không được làm Revit khởi động thất bại.
"""
from __future__ import unicode_literals

import io
import json
import os
import re

STATE_FILE_NAME = "ribbon_state.json"
SPLIT_SUFFIX = ".splitbutton"
# AppDomain sống suốt phiên Revit, qua cả pyRevit Reload: ApplicationClosing
# chỉ được nối một lần dù startup.py chạy lại mỗi lần Reload.
CLOSING_HOOK_KEY = "T3Lab.RibbonMemory.ClosingHooked"
# Idling đầu tiên thường đã có ribbon; nếu chưa thì thử thêm vài lượt rồi thôi.
RESTORE_MAX_TRIES = 20

# Split button đã gặp trên ribbon trong phiên này. Tool Ribbon Names đổi
# Title của tab giữa phiên, và GetRibbonPanels(tên cũ) có thể không còn thấy
# tab: lúc Revit đóng vẫn đọc được lựa chọn từ chính object đã giữ.
_LIVE = {}

# install() states
INSTALLED = "installed"
SKIPPED = "skipped"
FAILED = "failed"


def _err(exc):
    try:
        return "%s" % (getattr(exc, "Message", None) or exc)
    except Exception:
        return "<unprintable error>"


def _lib_dir():
    return os.path.dirname(os.path.abspath(__file__))


def _extension_dir():
    return os.path.dirname(_lib_dir())


def _tab_dirs(extension_dir=None):
    """Mọi thư mục `*.tab`, theo thứ tự ribbon (core.extension_paths)."""
    import sys
    if _lib_dir() not in sys.path:
        sys.path.insert(0, _lib_dir())
    try:
        from core.extension_paths import tab_dirs
        return tab_dirs(extension_dir or _extension_dir())
    except Exception:
        return []


# ── Trạng thái trên đĩa ──────────────────────────────────────────────────


def state_path():
    """%APPDATA%\\T3LabAI\\ribbon_state.json — trạng thái của người dùng, không
    bao giờ nằm trong thư mục extension (clone do git quản lý)."""
    try:
        import sys
        if _lib_dir() not in sys.path:
            sys.path.insert(0, _lib_dir())
        from core.paths import user_data_path
        return user_data_path(STATE_FILE_NAME)
    except Exception:
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "T3LabAI", STATE_FILE_NAME)


def load_state(path=None):
    """{tên split button: tên nút đứng trên cùng}. {} khi chưa có / file hỏng."""
    path = path or state_path()
    try:
        with io.open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        return {}
    splits = data.get("split_current") if isinstance(data, dict) else None
    if not isinstance(splits, dict):
        return {}
    return dict((k, v) for k, v in splits.items() if k and v)


def save_state(choices, path=None):
    """Gộp `choices` vào file (split không có mặt lần này vẫn giữ). True khi ghi được."""
    if not choices:
        return False
    path = path or state_path()
    merged = load_state(path)
    merged.update(choices)
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        text = json.dumps({"split_current": merged}, indent=2, sort_keys=True)
        with io.open(path, "w", encoding="utf-8") as fh:
            fh.write("%s\n" % text)
        return True
    except Exception:
        return False


# ── Ribbon trên đĩa ──────────────────────────────────────────────────────


def split_names(extension_dir=None):
    """Tên (folder bỏ đuôi) của mọi `.splitbutton` trong các tab của extension.

    pyRevit đặt Name của SplitButton = tên folder bỏ đuôi; tên bundle là duy
    nhất trong cả extension (dev/audit_ribbon.py P7), nên tên là đủ làm khoá.
    """
    names = []
    for tab in _tab_dirs(extension_dir):
        for root, dirs, _files in os.walk(tab):
            dirs.sort()
            base = os.path.basename(root)
            if base.endswith(SPLIT_SUFFIX):
                names.append(base[:-len(SPLIT_SUFFIX)])
    return names


def _bundle_title(folder):
    try:
        with io.open(os.path.join(folder, "bundle.yaml"), "r", encoding="utf-8") as fh:
            for line in fh:
                m = re.match(r"title\s*:\s*(.+?)\s*$", line)
                if m:
                    return m.group(1).strip("\"'")
    except Exception:
        pass
    return None


def tab_names(extension_dir=None):
    """Tên tab Revit có thể mang: `title:` của tab (pyRevit ưu tiên) rồi tên folder."""
    names = []
    for tab in _tab_dirs(extension_dir):
        for name in (_bundle_title(tab), os.path.basename(tab)[:-len(".tab")]):
            if name and name not in names:
                names.append(name)
    return names


# ── Ribbon trong Revit ───────────────────────────────────────────────────


def find_splits(uiapp, names, tabs):
    """{tên: SplitButton} cho các split button đang có trên ribbon. Never raises.

    RibbonPanel.GetItems() trả cả nút nằm trong stack (pyRevit dùng đúng cách
    này để tìm lại nút của stack khi Reload).
    """
    found = {}
    if uiapp is None or not names:
        return found
    for tab in tabs:
        try:
            panels = uiapp.GetRibbonPanels(tab)
        except Exception:
            continue            # tab chưa có (ribbon chưa dựng) hoặc khác tên
        for panel in panels or []:
            try:
                items = panel.GetItems()
            except Exception:
                continue
            for item in items or []:
                try:
                    name = item.Name
                except Exception:
                    continue
                if name in names and name not in found and hasattr(item, "CurrentButton"):
                    found[name] = item
    return found


def current_choices(splits):
    """{tên split: tên nút đang đứng trên cùng}."""
    out = {}
    for name, split in splits.items():
        try:
            current = split.CurrentButton
            if current is not None and current.Name:
                out[name] = current.Name
        except Exception:
            continue
    return out


def apply_choices(splits, choices):
    """Đặt nút trên cùng theo `choices`. Trả về số split đã đặt.

    Nút đã bị xoá / đổi tên thì bỏ qua: split giữ nút đầu như pyRevit dựng.
    """
    done = 0
    for name, split in splits.items():
        want = choices.get(name)
        if not want:
            continue
        try:
            for button in split.GetItems():
                if button.Name == want:
                    if split.CurrentButton is None or split.CurrentButton.Name != want:
                        split.CurrentButton = button
                    done += 1
                    break
        except Exception:
            continue
    return done


def _live_splits(uiapp, names, tabs):
    """find_splits() + split đã gặp trước đó mà lần tìm này không thấy."""
    splits = find_splits(uiapp, names, tabs)
    _LIVE.update(splits)
    for name, split in _LIVE.items():
        if name in names and name not in splits:
            splits[name] = split
    return splits


def remember(uiapp, path=None, names=None, tabs=None):
    """Ghi lựa chọn hiện tại của mọi split button xuống đĩa. Trả về số split đã ghi."""
    try:
        names = names if names is not None else split_names()
        tabs = tabs if tabs is not None else tab_names()
        choices = current_choices(_live_splits(uiapp, names, tabs))
        return len(choices) if save_state(choices, path) else 0
    except Exception:
        return 0


def restore(uiapp, path=None, names=None, tabs=None):
    """(số split tìm thấy trên ribbon, số split đã đặt lại theo file)."""
    try:
        names = names if names is not None else split_names()
        tabs = tabs if tabs is not None else tab_names()
        splits = find_splits(uiapp, names, tabs)
        if not splits:
            return 0, 0
        _LIVE.update(splits)
        return len(splits), apply_choices(splits, load_state(path))
    except Exception:
        return 0, 0


# ── Nối vào Revit (chỉ trong Revit) ──────────────────────────────────────


def _uiapp(revit_handle):
    """Object có GetRibbonPanels + Idling: UIApplication hoặc UIControlledApplication."""
    candidates = [revit_handle]
    try:
        from pyrevit import HOST_APP
        candidates.append(getattr(HOST_APP, "uiapp", None))
    except Exception:
        pass
    # Có bản pyRevit đưa __revit__ là DB.Application; UIApplication nhận nó
    # qua constructor (cùng cách lib/assistant_pane.py làm).
    if revit_handle is not None and not hasattr(revit_handle, "GetRibbonPanels"):
        try:
            from Autodesk.Revit.UI import UIApplication
            candidates.append(UIApplication(revit_handle))
        except Exception:
            pass
    for app in candidates:
        if app is not None and all(hasattr(app, a) for a in
                                   ("GetRibbonPanels", "Idling", "ApplicationClosing")):
            return app
    return None


def _restore_when_idle(uiapp, names, tabs):
    """restore() ở các lần Idling tới, tới khi gặp split button trên ribbon."""
    tries = [0]

    def on_idling(sender, args):
        tries[0] += 1
        found, _done = restore(uiapp, names=names, tabs=tabs)
        if found or tries[0] >= RESTORE_MAX_TRIES:
            try:
                uiapp.Idling -= on_idling
            except Exception:
                pass

    uiapp.Idling += on_idling


def _hook_closing_once(uiapp, names, tabs):
    """remember() lúc Revit đóng — nối một lần cho cả phiên Revit."""
    from System import AppDomain
    if AppDomain.CurrentDomain.GetData(CLOSING_HOOK_KEY):
        return False

    def on_closing(sender, args):
        remember(uiapp, names=names, tabs=tabs)

    uiapp.ApplicationClosing += on_closing
    AppDomain.CurrentDomain.SetData(CLOSING_HOOK_KEY, True)
    return True


def install(revit_handle):
    """Nối việc nhớ / đặt lại split button. Returns (state, detail); never raises."""
    try:
        names = split_names()
        if not names:
            return SKIPPED, "no .splitbutton on the ribbon"
        tabs = tab_names()
        uiapp = _uiapp(revit_handle)
        if uiapp is None:
            return FAILED, "no UIApplication / UIControlledApplication at startup"
        # Reload: ribbon cũ vẫn giữ lựa chọn của phiên này. Lần khởi động đầu
        # ribbon chưa có nên không ghi gì, file của lần đóng trước còn nguyên.
        remember(uiapp, names=names, tabs=tabs)
        _restore_when_idle(uiapp, names, tabs)
        _hook_closing_once(uiapp, names, tabs)
        return INSTALLED, ", ".join(names)
    except Exception as exc:
        return FAILED, _err(exc)
