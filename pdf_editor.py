# -*- coding: utf-8 -*-
"""
JP-PDF EDITOR  v1.2  -  설치 없이 쓰는 오프라인 PDF 편집기 (레이어 방식)

 v1.2 변경
 * 화면 캡처(Win+Shift+S, PrtSc) 후 Ctrl+V -> 이미지 레이어로 붙여넣기
   (탐색기에서 복사한 그림 파일도 가능, 모서리 핸들은 비율 유지, 작업 파일에 이미지 함께 저장)

 v1.1 (R01) 변경
 * 텍스트 레이어 상자가 글자 크기·내용에 맞춰 자동으로 맞춰짐 (크기 핸들을 끌면 글자 크기가 바뀜)
 * 텍스트 굵기 지정 (보통 / 약간 굵게 / 굵게 / 아주 굵게)
 * 텍스트 추가·수정을 팝업창 없이 페이지 위에서 바로 입력
 * '도형 그리기' 도구 아래에서 도형 종류 선택 (사각형·둥근 사각형·원·삼각형·직선·화살표·양쪽 화살표)

 * 텍스트 / 형광펜 / 도형은 PDF에 바로 새겨지지 않고 '레이어 개체'로 얹힙니다.
   -> 선택, 이동, 크기 조절, 속성(글꼴·색·두께…) 변경, 텍스트 내용 수정, 삭제, 앞/뒤 순서 변경 가능
 * 'PDF로 저장' 할 때만 개체가 PDF에 합쳐집니다.
 * 나중에 다시 수정하려면 '작업 파일(.pdfedit)'로 저장하세요 (원본 PDF + 레이어 정보).
 * 페이지: 분할 / 병합 / 삭제 / 추가(빈 페이지·PDF·이미지) / 순서 변경 / 회전

설치:   pip install pymupdf fonttools
실행:   python pdf_editor.py  [파일.pdf 또는 파일.pdfedit]
exe화:  build_exe.bat 실행  (dist 폴더에 "JP-PDF EDITOR.exe" 생성 - 설치 없이 실행되는 단일 파일)
"""
import os
import sys
import math
import json
import copy
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser
from tkinter import font as tkfont

try:
    import pymupdf as fitz
except ImportError:  # 구버전 호환
    import fitz

APP_TITLE = "JP-PDF EDITOR"
APP_VERSION = "1.2"
PROJ_EXT = ".pdfedit"
THUMB_W, THUMB_H = 110, 140
PAD = 12
UNDO_LIMIT = 40
TEXT_PAD = 2
HANDLE_R = 5


def resource_path(name):
    """exe(PyInstaller)로 묶였을 때와 일반 실행 모두에서 데이터 파일 경로 찾기"""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


# ============================================================
#  글꼴
# ============================================================
def _font_dirs():
    dirs = []
    if sys.platform.startswith("win"):
        win = os.environ.get("WINDIR", r"C:\Windows")
        dirs.append(os.path.join(win, "Fonts"))
        la = os.environ.get("LOCALAPPDATA")
        if la:
            dirs.append(os.path.join(la, "Microsoft", "Windows", "Fonts"))
    elif sys.platform == "darwin":
        dirs += ["/System/Library/Fonts", "/Library/Fonts",
                 os.path.expanduser("~/Library/Fonts")]
    else:
        dirs += ["/usr/share/fonts", os.path.expanduser("~/.fonts")]
    return dirs


FONT_CANDIDATES = [
    ("맑은 고딕", "malgun.ttf"), ("맑은 고딕 Bold", "malgunbd.ttf"),
    ("바탕", "batang.ttc"), ("굴림", "gulim.ttc"), ("돋움", "dotum.ttc"),
    ("궁서", "gungsuh.ttc"),
    ("나눔고딕", "NanumGothic.ttf"), ("나눔고딕 Bold", "NanumGothicBold.ttf"),
    ("나눔바른고딕", "NanumBarunGothic.ttf"), ("나눔명조", "NanumMyeongjo.ttf"),
    ("Arial", "arial.ttf"), ("Arial Bold", "arialbd.ttf"),
    ("Times New Roman", "times.ttf"), ("Courier New", "cour.ttf"),
    ("Consolas", "consola.ttf"), ("Calibri", "calibri.ttf"),
]
KOREAN_FALLBACKS = ["맑은 고딕", "나눔고딕", "굴림", "바탕", "돋움"]


class FontSpec:
    """kind='file' (TTF/TTC/OTF 파일) 또는 'builtin' (helv/tiro/cour)"""

    def __init__(self, label, kind, value):
        self.label, self.kind, self.value = label, kind, value
        self._font = None

    def font(self):
        if self._font is None:
            self._font = (fitz.Font(fontfile=self.value) if self.kind == "file"
                          else fitz.Font(self.value))
        return self._font

    def supports(self, text):
        f = self.font()
        return all(f.has_glyph(ord(c)) for c in text if not c.isspace())

    def kwargs(self):
        if self.kind == "file":
            return {"fontname": "F%d" % (abs(hash(self.value)) % 10 ** 7),
                    "fontfile": self.value}
        return {"fontname": self.value}


def find_fonts():
    fonts = {}
    for label, fname in FONT_CANDIDATES:
        for d in _font_dirs():
            p = os.path.join(d, fname)
            if os.path.isfile(p):
                fonts[label] = FontSpec(label, "file", p)
                break
    fonts["Helvetica (기본·영문만)"] = FontSpec("Helvetica (기본·영문만)", "builtin", "helv")
    fonts["Times (기본·영문만)"] = FontSpec("Times (기본·영문만)", "builtin", "tiro")
    fonts["Courier (기본·영문만)"] = FontSpec("Courier (기본·영문만)", "builtin", "cour")
    return fonts


FONTS = find_fonts()
_SPEC_CACHE = {}


def default_spec():
    for n in KOREAN_FALLBACKS:
        if n in FONTS:
            return FONTS[n]
    return next(iter(FONTS.values()))


def korean_fallback():
    for n in KOREAN_FALLBACKS:
        if n in FONTS:
            return FONTS[n]
    return None


def spec_for(o):
    """개체에 저장된 글꼴 정보 -> FontSpec (파일이 없으면 기본 한글 글꼴로 대체)"""
    key = (o.font_kind, o.font_value)
    if key in _SPEC_CACHE:
        return _SPEC_CACHE[key]
    spec = None
    if o.font_kind == "file" and os.path.isfile(o.font_value):
        spec = FontSpec(o.font_label or os.path.basename(o.font_value), "file", o.font_value)
    elif o.font_kind == "builtin":
        spec = FontSpec(o.font_label, "builtin", o.font_value)
    if spec is None:
        spec = FONTS.get(o.font_label) or default_spec()
    _SPEC_CACHE[key] = spec
    return spec


# ============================================================
#  유틸
# ============================================================
def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def parse_ranges(text, n):
    """'1-3, 5, 7-' -> [('1-3',[0,1,2]), ('5',[4]), ('7-n',[6..n-1])]  (1-based 입력)"""
    out = []
    for part in text.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a = int(a) if a else 1
            b = int(b) if b else n
        else:
            a = b = int(part)
        if a < 1 or b > n or a > b:
            raise ValueError("범위 오류: %s  (1~%d 사이여야 합니다)" % (part, n))
        out.append(("%d-%d" % (a, b) if a != b else str(a), list(range(a - 1, b))))
    if not out:
        raise ValueError("범위를 입력하세요.")
    return out


def group_contiguous(indices):
    groups = []
    for i in indices:
        if groups and i == groups[-1][1] + 1:
            groups[-1][1] = i
        else:
            groups.append([i, i])
    return groups


def unique_path(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    k = 1
    while os.path.exists("%s_%d%s" % (base, k, ext)):
        k += 1
    return "%s_%d%s" % (base, k, ext)


def _silent_remove(p):
    try:
        os.remove(p)
    except OSError:
        pass


def _atomic_write(path, data):
    """임시파일에 쓴 뒤 교체. 한글 경로 / 다른 프로그램이 파일을 열고 있는 경우(윈도우) 대응"""
    tmp_path = path + ".tmp"
    try:
        with open(tmp_path, "wb") as f:
            f.write(data)
        os.replace(tmp_path, path)
    except PermissionError:
        _silent_remove(tmp_path)
        raise PermissionError("'%s' 파일이 다른 프로그램(PDF 뷰어 등)에서 열려 있어 저장할 수 없습니다.\n"
                              "그 프로그램을 닫거나 다른 이름으로 저장하세요." % os.path.basename(path))
    except Exception:
        _silent_remove(tmp_path)
        raise


def save_doc(doc, path, subset=False, inplace=False):
    """PDF 저장. subset=True 이면 삽입한 글꼴을 필요한 글자만 남겨 용량 축소(fonttools 필요)"""
    if subset:
        target = doc if inplace else fitz.open("pdf", doc.tobytes())
        try:
            target.subset_fonts()
        except Exception:
            pass
        data = target.tobytes(garbage=3, deflate=True)
        if not inplace:
            target.close()
    else:
        data = doc.tobytes(garbage=3, deflate=True)
    _atomic_write(path, data)


# ============================================================
#  페이지 좌표계
#   개체 좌표는 '회전 전(page space, P)' 기준으로 저장 -> 페이지를 돌려도 개체가 같이 돈다.
#   화면 좌표(V)는 회전이 반영된 page.rect 기준.  V = P * rotm,  P = V * derotm
# ============================================================
class PageGeo:
    def __init__(self, page):
        self.rot = page.rotation
        self.rotm = fitz.Matrix(page.rotation_matrix)
        self.derotm = fitz.Matrix(page.derotation_matrix)
        r = page.rect
        self.vw, self.vh = r.width, r.height
        u = (r * self.derotm).normalize()
        self.uw, self.uh = u.width, u.height


def rect_v2p(r, geo):
    return (fitz.Rect(r) * geo.derotm).normalize()


def rect_p2v(r, geo):
    return (fitz.Rect(r) * geo.rotm).normalize()


# ============================================================
#  레이어 개체
# ============================================================
BOX_KINDS = ("text", "rect", "rrect", "oval", "triangle", "image")
LINE_KINDS = ("line", "arrow", "darrow")
ARROW_KINDS = ("arrow", "darrow")
KIND_LABEL = {"text": "텍스트", "rect": "사각형", "rrect": "둥근 사각형", "oval": "원", "triangle": "삼각형",
              "line": "직선", "arrow": "화살표", "darrow": "양쪽 화살표", "highlight": "형광펜", "image": "이미지"}
KIND_ICON = {"text": "T", "rect": "▭", "rrect": "▢", "oval": "◯", "triangle": "△", "line": "╱",
             "arrow": "➜", "darrow": "↔", "highlight": "▬", "image": "▣"}
# '도형 그리기' 도구 아래에서 고르는 도형 (순서 = 화면 표시 순서)
SHAPE_TYPES = [("rect", "▭ 사각형"), ("rrect", "▢ 둥근 사각형"), ("oval", "◯ 원 / 타원"),
               ("triangle", "△ 삼각형"), ("line", "╱ 직선"), ("arrow", "➜ 화살표"),
               ("darrow", "↔ 양쪽 화살표")]
RRECT_RADIUS = 0.2          # 둥근 사각형 모서리 반경 (짧은 변 대비 비율)
# 텍스트 굵기: 글자 외곽선 두께 (글자 크기 대비 비율) - 어떤 글꼴이든 굵게 만들 수 있다
WEIGHTS = [("보통", 0.0), ("약간 굵게", 0.025), ("굵게", 0.05), ("아주 굵게", 0.09)]
WEIGHT_NAMES = [n for n, _ in WEIGHTS]


# ---- 이미지 레이어: 그림 데이터는 개체가 아니라 여기에 한 번만 보관 (키 = 내용 해시)
#   개체에는 키만 들어 있어 실행 취소 기록·복사가 가볍다. 작업 파일에는 images/<키>.png 로 저장.
IMAGE_STORE = {}


def store_image(png):
    import hashlib
    key = hashlib.sha1(png).hexdigest()[:20]
    IMAGE_STORE[key] = png
    return key


def to_png(data):
    """이미지 바이트(PNG/JPG/BMP 등) -> (PNG 바이트, 폭px, 높이px). 읽을 수 없으면 예외"""
    pm = fitz.Pixmap(data)
    if pm.colorspace and pm.colorspace.n not in (1, 3):    # CMYK 등 -> RGB
        pm = fitz.Pixmap(fitz.csRGB, pm)
    if pm.alpha:
        a = bytes(pm.samples_mv[pm.n - 1::pm.n])
        if not a or a.count(a[0]) == len(a):               # 캡처(32비트)처럼 알파가 전부 같으면 -> 불투명
            pm = fitz.Pixmap(pm, 0)
    return pm.tobytes("png"), pm.width, pm.height


# ---- 윈도우 클립보드에서 이미지 읽기 (추가 라이브러리 없이 ctypes)
CF_DIB, CF_HDROP = 8, 15
IMG_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff")


def _clip_api():
    import ctypes
    from ctypes import wintypes as wt
    u, k, s = ctypes.windll.user32, ctypes.windll.kernel32, ctypes.windll.shell32
    u.OpenClipboard.argtypes = [wt.HWND]
    u.GetClipboardData.argtypes = [wt.UINT]
    u.GetClipboardData.restype = wt.HANDLE
    u.IsClipboardFormatAvailable.argtypes = [wt.UINT]
    u.RegisterClipboardFormatW.argtypes = [wt.LPCWSTR]
    u.RegisterClipboardFormatW.restype = wt.UINT
    k.GlobalLock.argtypes = [wt.HGLOBAL]
    k.GlobalLock.restype = ctypes.c_void_p
    k.GlobalUnlock.argtypes = [wt.HGLOBAL]
    k.GlobalSize.argtypes = [wt.HGLOBAL]
    k.GlobalSize.restype = ctypes.c_size_t
    s.DragQueryFileW.argtypes = [wt.HANDLE, wt.UINT, wt.LPWSTR, wt.UINT]
    s.DragQueryFileW.restype = wt.UINT
    return ctypes, u, k, s


def clipboard_seq():
    try:
        import ctypes
        return ctypes.windll.user32.GetClipboardSequenceNumber()
    except Exception:
        return None


def dib_to_bmp(dib):
    """클립보드 CF_DIB(BITMAPINFO + 픽셀) 앞에 BMP 파일 머리를 붙인다"""
    import struct
    hsize, = struct.unpack_from("<I", dib, 0)
    bpp, comp = struct.unpack_from("<HI", dib, 14)
    used, = struct.unpack_from("<I", dib, 32)
    off = hsize
    if comp == 3 and hsize == 40:                 # BI_BITFIELDS: 색 마스크 3개
        off += 12
    if bpp <= 8:
        off += 4 * (used or (1 << bpp))
    return struct.pack("<2sIHHI", b"BM", 14 + len(dib), 0, 0, 14 + off) + dib


def clipboard_image():
    """클립보드의 이미지 -> 이미지 바이트. 캡처(PNG/비트맵) 또는 탐색기에서 복사한 그림 파일. 없으면 None"""
    if not sys.platform.startswith("win"):
        return None
    try:
        ctypes, u, k, s = _clip_api()
    except Exception:
        return None
    opened = False
    for _ in range(10):                           # 다른 프로그램이 잠깐 잡고 있을 수 있다
        if u.OpenClipboard(None):
            opened = True
            break
        import time
        time.sleep(0.05)
    if not opened:
        return None

    def read(fmt):
        h = u.GetClipboardData(fmt)
        if not h:
            return None
        p = k.GlobalLock(h)
        if not p:
            return None
        try:
            return ctypes.string_at(p, k.GlobalSize(h))
        finally:
            k.GlobalUnlock(h)

    try:
        png_fmt = u.RegisterClipboardFormatW("PNG")
        if png_fmt and u.IsClipboardFormatAvailable(png_fmt):
            data = read(png_fmt)
            if data:
                return data
        if u.IsClipboardFormatAvailable(CF_DIB):
            data = read(CF_DIB)
            if data:
                return dib_to_bmp(data)
        if u.IsClipboardFormatAvailable(CF_HDROP):
            h = u.GetClipboardData(CF_HDROP)
            n = s.DragQueryFileW(h, 0xFFFFFFFF, None, 0) if h else 0
            for i in range(n):
                buf = ctypes.create_unicode_buffer(1024)
                s.DragQueryFileW(h, i, buf, 1024)
                if buf.value.lower().endswith(IMG_EXTS) and os.path.isfile(buf.value):
                    with open(buf.value, "rb") as f:
                        return f.read()
    finally:
        u.CloseClipboard()
    return None


def weight_ratio(o):
    try:
        return WEIGHTS[int(o.weight)][1]
    except (ValueError, TypeError, IndexError):
        return 0.0


class Obj:
    _counter = 0
    FIELDS = dict(
        kind="rect", visible=True, name="",
        rect=None, p0=None, p1=None, rects=None,                 # 기하 (page space)
        text="", font_kind="builtin", font_value="helv", font_label="",
        size=14.0, align=0, color="#000000", bg=None, auto_w=False,   # 텍스트/형광펜
        weight=0,                                                     # 텍스트 굵기 (WEIGHTS 번호)
        img="", img_rot=0,                                            # 이미지: IMAGE_STORE 키, 붙일 때 페이지 회전
        stroke="#FF0000", fill=None, width=2.0, opacity=1.0)          # 도형

    def __init__(self, kind, **kw):
        Obj._counter += 1
        self.id = Obj._counter
        for k, v in self.FIELDS.items():
            setattr(self, k, copy.deepcopy(v))
        self.kind = kind
        for k, v in kw.items():
            setattr(self, k, v)
        self._img = None
        self._img_key = None
        self._lay_key = None
        self._lay_val = None

    # ---- 직렬화
    def to_dict(self):
        d = {"id": self.id}
        for k in self.FIELDS:
            d[k] = copy.deepcopy(getattr(self, k))
        return d

    @classmethod
    def from_dict(cls, d):
        o = cls(d.get("kind", "rect"))
        for k in cls.FIELDS:
            if k in d:
                setattr(o, k, copy.deepcopy(d[k]))
        o.id = d.get("id", o.id)
        cls._counter = max(cls._counter, o.id)
        if o.rect is not None:
            o.rect = tuple(o.rect)
        if o.p0 is not None:
            o.p0 = tuple(o.p0)
        if o.p1 is not None:
            o.p1 = tuple(o.p1)
        if o.rects is not None:
            o.rects = [tuple(r) for r in o.rects]
        return o

    def clone(self, new_id=True):
        o = Obj.from_dict(self.to_dict())
        if new_id:
            Obj._counter += 1
            o.id = Obj._counter
        return o

    def sig(self):
        return repr(tuple(getattr(self, k) for k in self.FIELDS if k not in ("visible", "name")))

    # ---- 이동
    def shift(self, dx, dy):
        if self.rect:
            x0, y0, x1, y1 = self.rect
            self.rect = (x0 + dx, y0 + dy, x1 + dx, y1 + dy)
        if self.p0:
            self.p0 = (self.p0[0] + dx, self.p0[1] + dy)
        if self.p1:
            self.p1 = (self.p1[0] + dx, self.p1[1] + dy)
        if self.rects:
            self.rects = [(a + dx, b + dy, c + dx, d + dy) for a, b, c, d in self.rects]

    def label(self):
        if self.name:
            return self.name
        if self.kind == "text":
            t = " ".join(self.text.split())
            return t[:14] + ("…" if len(t) > 14 else "")
        return KIND_LABEL[self.kind]


# ---- 텍스트 배치 (줄바꿈 / 자동 높이)
def measure_text(spec, text, size, width, align):
    """주어진 폭에서 줄바꿈했을 때 필요한 높이와 (필요 시 넓힌) 폭"""
    tmp = fitz.open()
    try:
        pg = tmp.new_page(width=4000, height=4000)
        w = max(width, 10.0)
        for _ in range(8):
            rc = pg.insert_textbox(fitz.Rect(0, 0, w, 3990), text, fontsize=size,
                                   align=align, **spec.kwargs())
            if rc >= 0:
                return 3990 - rc, w
            w = min(w * 1.3, 3900)
        return size * 1.3 * (text.count("\n") + 1), w
    finally:
        tmp.close()


def text_layout(o, geo):
    """텍스트 개체가 실제로 차지하는 사각형(화면 기준 V).
    상자는 글자 크기·내용에 딱 맞춘다 (auto_w=True: 폭도 글에 맞춤, False: 지정 폭에서 줄바꿈)."""
    key = (o.text, o.font_kind, o.font_value, round(o.size, 3), o.align, o.weight,
           tuple(round(v, 3) for v in o.rect), o.auto_w, geo.rot, round(geo.vw, 2))
    if o._lay_key == key and o._lay_val is not None:
        return fitz.Rect(o._lay_val)
    spec = spec_for(o)
    pad = TEXT_PAD + o.size * weight_ratio(o) / 2      # 굵게(외곽선)만큼 여유
    rv = rect_p2v(o.rect, geo)
    if o.auto_w:
        fobj = spec.font()
        natural = max(fobj.text_length(l, fontsize=o.size) for l in o.text.split("\n")) + 2 * pad + 2
        w = min(natural, geo.vw - 4)
        rv = fitz.Rect(rv.x0, rv.y0, rv.x0 + w, rv.y0 + o.size * 1.3)
    used, inner_w = measure_text(spec, o.text, o.size, rv.width - 2 * pad, o.align)
    h = used + 2 * pad
    rv = fitz.Rect(rv.x0, rv.y0, rv.x0 + inner_w + 2 * pad, rv.y0 + h)
    o._lay_key, o._lay_val = key, tuple(rv)
    return rv


def bbox_p(o, geo):
    """개체의 (선택용) 외곽 사각형 - page space"""
    if o.kind == "text":
        return rect_v2p(text_layout(o, geo), geo)
    if o.kind == "highlight":
        r = fitz.Rect(o.rects[0])
        for q in o.rects[1:]:
            r |= fitz.Rect(q)
        return r
    if o.kind in LINE_KINDS:
        return fitz.Rect(fitz.Point(o.p0), fitz.Point(o.p1)).normalize()
    return fitz.Rect(o.rect).normalize()


def render_bbox_p(o, geo):
    """그릴 때 필요한 여백 포함 사각형"""
    r = fitz.Rect(bbox_p(o, geo))
    if o.kind == "text":
        m = 3
    elif o.kind in ("highlight", "image"):
        m = 2
    else:
        m = o.width / 2 + 2 + (max(10.0, o.width * 5) if o.kind in ARROW_KINDS else 0)
    return fitz.Rect(r.x0 - m, r.y0 - m, r.x1 + m, r.y1 + m)


# ---- 그리기 (page space, 회전 전 좌표)
def draw_obj(page, o, geo):
    k = o.kind
    if k == "text":
        spec = spec_for(o)
        rv = text_layout(o, geo)
        bw = weight_ratio(o)
        pad = TEXT_PAD + o.size * bw / 2
        if o.bg:
            page.draw_rect(rect_v2p(rv, geo), color=None, fill=hex_to_rgb(o.bg), overlay=True)
        inner = fitz.Rect(rv.x0 + pad, rv.y0 + pad, rv.x1 - pad, rv.y1 - pad)
        col = hex_to_rgb(o.color)
        bold = dict(render_mode=2, border_width=bw, fill=col) if bw > 0 else {}
        page.insert_textbox(rect_v2p(inner, geo), o.text, fontsize=o.size,
                            color=col, align=o.align,
                            rotate=geo.rot, **bold, **spec.kwargs())
    elif k == "image":
        data = IMAGE_STORE.get(o.img)
        if data:     # 붙일 때의 페이지 회전만큼 돌려 그 화면에서 바로 보이게. 나중에 페이지를 돌리면 함께 돈다
            page.insert_image(fitz.Rect(o.rect).normalize(), stream=data, keep_proportion=False,
                              rotate=o.img_rot, overlay=True)
    elif k == "highlight":
        for r in o.rects:
            page.draw_rect(fitz.Rect(r).normalize(), color=None, fill=hex_to_rgb(o.color),
                           fill_opacity=o.opacity, overlay=True)
    else:
        stroke = hex_to_rgb(o.stroke)
        fill = hex_to_rgb(o.fill) if o.fill else None
        common = dict(color=stroke, width=o.width, stroke_opacity=o.opacity, overlay=True)
        if k == "rect":
            page.draw_rect(fitz.Rect(o.rect).normalize(), fill=fill, fill_opacity=o.opacity, **common)
        elif k == "rrect":
            page.draw_rect(fitz.Rect(o.rect).normalize(), fill=fill, fill_opacity=o.opacity,
                           radius=RRECT_RADIUS, **common)
        elif k == "oval":
            page.draw_oval(fitz.Rect(o.rect).normalize(), fill=fill, fill_opacity=o.opacity, **common)
        elif k == "triangle":
            page.draw_polyline(triangle_pts(o, geo), fill=fill, fill_opacity=o.opacity,
                               closePath=True, lineJoin=1, **common)
        else:
            a, b = fitz.Point(o.p0), fitz.Point(o.p1)
            page.draw_line(a, b, **common)
            heads = [(a, b)] if k == "arrow" else [(a, b), (b, a)] if k == "darrow" else []
            for s, e in heads:
                ang = math.atan2(e.y - s.y, e.x - s.x)
                L = max(10.0, o.width * 5)
                h1 = fitz.Point(e.x - L * math.cos(ang - 0.45), e.y - L * math.sin(ang - 0.45))
                h2 = fitz.Point(e.x - L * math.cos(ang + 0.45), e.y - L * math.sin(ang + 0.45))
                page.draw_polyline([e, h1, h2], color=stroke, fill=stroke, width=o.width,
                                   closePath=True, stroke_opacity=o.opacity,
                                   fill_opacity=o.opacity, overlay=True)


def triangle_pts(o, geo):
    """삼각형 꼭짓점 (page space). 화면에서 위쪽이 꼭짓점이 되도록 화면 좌표(V)에서 만든 뒤 되돌린다"""
    rv = rect_p2v(o.rect, geo)
    pts_v = [fitz.Point((rv.x0 + rv.x1) / 2, rv.y0), fitz.Point(rv.x1, rv.y1), fitz.Point(rv.x0, rv.y1)]
    return [p * geo.derotm for p in pts_v]


def apply_objects(page, objs):
    """실제 PDF 페이지에 (보이는) 개체들을 순서대로 합친다"""
    vis = [o for o in objs if o.visible]
    if not vis:
        return
    try:
        page.wrap_contents()
    except Exception:
        pass
    geo = PageGeo(page)
    for o in vis:
        draw_obj(page, o, geo)


def render_obj_png(o, geo, zoom):
    """개체 1개를 투명 배경 PNG로 렌더 -> (png bytes, 장치좌표 x, y)"""
    bb = render_bbox_p(o, geo)
    tmp = fitz.open()
    try:
        pg = tmp.new_page(width=max(geo.uw, 1), height=max(geo.uh, 1))
        draw_obj(pg, o, geo)
        if geo.rot:
            pg.set_rotation(geo.rot)
        clip = (bb * geo.rotm).normalize()
        clip.intersect(pg.rect)
        if clip.is_empty or clip.width < 0.2 or clip.height < 0.2:
            return None
        pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=True)
        return pix.tobytes("png"), pix.x, pix.y
    finally:
        tmp.close()


# ---- 맞춤(hit) 검사
def dist_seg(p, a, b):
    ax, ay, bx, by = a[0], a[1], b[0], b[1]
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    t = 0.0 if L2 == 0 else max(0.0, min(1.0, ((p.x - ax) * dx + (p.y - ay) * dy) / L2))
    return math.hypot(p.x - (ax + t * dx), p.y - (ay + t * dy))


def obj_hit(o, pt, tol, geo):
    k = o.kind
    if k in ("text", "image"):
        r = bbox_p(o, geo)
        return fitz.Rect(r.x0 - tol, r.y0 - tol, r.x1 + tol, r.y1 + tol).contains(pt)
    if k == "highlight":
        return any(fitz.Rect(r.x0 - tol, r.y0 - tol, r.x1 + tol, r.y1 + tol).contains(pt)
                   for r in (fitz.Rect(q).normalize() for q in o.rects))
    if k in LINE_KINDS:
        return dist_seg(pt, o.p0, o.p1) <= max(tol, o.width / 2 + 1)
    if k == "triangle":
        pts = triangle_pts(o, geo)
        if o.fill:
            return _in_poly(pt, pts) or any(dist_seg(pt, pts[i], pts[(i + 1) % 3]) <= tol + o.width / 2
                                            for i in range(3))
        return any(dist_seg(pt, pts[i], pts[(i + 1) % 3]) <= tol + o.width / 2 for i in range(3))
    r = fitz.Rect(o.rect).normalize()
    if k in ("rect", "rrect"):
        t = tol + o.width / 2
        outer = (r.x0 - t <= pt.x <= r.x1 + t) and (r.y0 - t <= pt.y <= r.y1 + t)
        if not outer:
            return False
        if o.fill:
            return True
        return not ((r.x0 + t < pt.x < r.x1 - t) and (r.y0 + t < pt.y < r.y1 - t))
    # oval
    cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
    a, b = max(r.width / 2, 0.01), max(r.height / 2, 0.01)
    rr = math.hypot((pt.x - cx) / a, (pt.y - cy) / b)
    m = min(a, b)
    if o.fill:
        return rr <= 1 + (tol + o.width / 2) / m
    return abs(rr - 1) * m <= tol + o.width / 2


def _in_poly(pt, pts):
    inside = False
    n = len(pts)
    for i in range(n):
        a, b = pts[i], pts[(i + 1) % n]
        if (a.y > pt.y) != (b.y > pt.y):
            x = a.x + (pt.y - a.y) * (b.x - a.x) / (b.y - a.y)
            if pt.x < x:
                inside = not inside
    return inside


# ---- 페이지 위 바로 입력(인라인 편집기)용 화면 글꼴
#   PDF 글꼴 이름 -> Tk 글꼴 이름. 입력하는 동안만 쓰고, 확정하면 실제 PDF 글꼴로 레이어를 그린다.
TK_FAMILY = {"맑은 고딕": "맑은 고딕", "맑은 고딕 Bold": "맑은 고딕", "바탕": "바탕", "굴림": "굴림",
             "돋움": "돋움", "궁서": "궁서", "나눔고딕": "나눔고딕", "나눔고딕 Bold": "나눔고딕",
             "나눔바른고딕": "나눔바른고딕", "나눔명조": "나눔명조", "Arial": "Arial", "Arial Bold": "Arial",
             "Times New Roman": "Times New Roman", "Courier New": "Courier New", "Consolas": "Consolas",
             "Calibri": "Calibri", "Helvetica (기본·영문만)": "Arial",
             "Times (기본·영문만)": "Times New Roman", "Courier (기본·영문만)": "Courier New"}
_TK_FAMILIES = None


def tk_family(label):
    global _TK_FAMILIES
    if _TK_FAMILIES is None:
        try:
            _TK_FAMILIES = set(tkfont.families())
        except tk.TclError:
            _TK_FAMILIES = set()
    for fam in (TK_FAMILY.get(label, label), "맑은 고딕", "Malgun Gothic", "Arial"):
        if fam in _TK_FAMILIES:
            return fam
    return "TkDefaultFont"


# ============================================================
#  화면 배율 / 스크롤 영역
# ============================================================
def screen_work_area(root):
    """작업 표시줄을 뺀 화면 영역 (x, y, 폭, 높이)"""
    if sys.platform.startswith("win"):
        try:
            import ctypes
            from ctypes import wintypes
            r = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(0x30, 0, ctypes.byref(r), 0):  # SPI_GETWORKAREA
                return r.left, r.top, r.right - r.left, r.bottom - r.top
        except Exception:
            pass
    return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight() - 60


class ScrollFrame(ttk.Frame):
    """세로 스크롤 영역 (내용은 .body 에 넣는다).
    공간이 넉넉하면 내용이 영역 높이만큼 늘어나고, 모자라면 스크롤바가 나타나 마우스 휠로 내려 볼 수 있다.
    (화면 배율 125%·150%나 세로가 낮은 화면에서 오른쪽 속성 패널 아래쪽이 잘리던 문제)"""

    def __init__(self, master, step=20, measure=None):
        super().__init__(master)
        self.measure = measure          # 필요한 내용 높이를 직접 셀 함수 (없으면 내용의 요청 높이)
        self.canvas = tk.Canvas(self, width=1, height=1, highlightthickness=0, bd=0,
                                yscrollincrement=step, takefocus=0)
        bg = ttk.Style().lookup("TFrame", "background")
        if bg:
            self.canvas.configure(bg=bg)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.vsb.grid_remove()
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.body = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window(0, 0, window=self.body, anchor="nw")
        self._last = None
        self._acc = 0
        self._tag = "ScrollFrame%d" % id(self)
        self.canvas.bind("<Configure>", self._relayout)
        self.body.bind("<Configure>", self._relayout)
        self.bind_class(self._tag, "<MouseWheel>", self._on_wheel)
        self.bind_class(self._tag, "<Button-4>", lambda e: self._on_wheel(e, 120))
        self.bind_class(self._tag, "<Button-5>", lambda e: self._on_wheel(e, -120))

    def content_width(self):
        """내용이 잘리지 않는 폭 (스크롤바 포함)"""
        return self.body.winfo_reqwidth() + self.vsb.winfo_reqwidth()

    def refresh(self):
        self._last = None
        self._relayout()

    def _relayout(self, event=None):
        c = self.canvas
        bw = self.body.winfo_reqwidth()
        bh = self.measure() if self.measure else self.body.winfo_reqheight()
        if int(float(c.cget("width"))) != bw:
            c.configure(width=bw)                       # 패널이 요청하는 폭 = 내용 폭
        cw, ch = c.winfo_width(), c.winfo_height()
        if ch <= 1:
            return
        h = max(ch, bh)
        state = (cw, h, bh > ch)
        if state == self._last:
            return
        self._last = state
        c.itemconfigure(self._win, width=cw, height=h)
        c.configure(scrollregion=(0, 0, cw, h))
        if bh > ch:
            self.vsb.grid()
        else:
            self.vsb.grid_remove()
            c.yview_moveto(0)

    def watch(self):
        """영역 안 위젯 위에서 마우스 휠을 굴리면 영역이 스크롤되게 한다 (위젯을 다 만든 뒤 호출).
        목록(Treeview)과 그 스크롤바는 제 스크롤을 그대로 쓴다."""
        def add(w):
            if not isinstance(w, (ttk.Treeview, ttk.Scrollbar)) or w is self.vsb:
                tags = w.bindtags()
                if self._tag not in tags:
                    w.bindtags((tags[0], self._tag) + tags[1:])     # 클래스 바인딩보다 먼저
            for ch in w.winfo_children():
                add(ch)
        add(self)

    def _on_wheel(self, e, delta=None):
        if not self.vsb.winfo_ismapped():
            return None                                 # 다 보이면 원래 동작 그대로
        w = e.widget
        if isinstance(w, (ttk.Combobox, ttk.Spinbox)) and w.focus_get() is w:
            return None                                 # 입력 중인 값 칸은 휠로 값 변경 (원래 동작)
        self._acc += e.delta if delta is None else delta
        n = int(self._acc / 120)
        if n:
            self._acc -= n * 120
            self.canvas.yview_scroll(-n * 3, "units")
        return "break"


# ============================================================
#  보조 대화상자
# ============================================================

CURRENT_DOC = "\0current"


class MergeDialog(tk.Toplevel):
    """여러 PDF를 순서 지정해서 하나로 병합"""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("PDF 병합")
        self.transient(app)
        self.geometry("620x400")
        self.files = []
        fr = ttk.Frame(self)
        fr.pack(fill="both", expand=True, padx=10, pady=10)
        self.lb = tk.Listbox(fr, selectmode="extended")
        sb = ttk.Scrollbar(fr, command=self.lb.yview)
        self.lb.config(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        col = ttk.Frame(fr)
        col.pack(side="left", padx=8, fill="y")
        for text, cmd in [("파일 추가…", self.add), ("제거", self.remove),
                          ("▲ 위로", lambda: self.move(-1)), ("▼ 아래로", lambda: self.move(1))]:
            ttk.Button(col, text=text, command=cmd, width=12).pack(pady=3)
        bar = ttk.Frame(self)
        bar.pack(pady=(0, 10))
        ttk.Button(bar, text="병합하여 저장…", command=self.merge).pack(side="left", padx=4)
        ttk.Button(bar, text="닫기", command=self.destroy).pack(side="left", padx=4)
        if app.doc is not None:
            self.files.append(CURRENT_DOC)
            self.lb.insert("end", "(현재 편집 중인 문서 - 레이어 포함)")

    def add(self):
        for p in filedialog.askopenfilenames(parent=self, title="병합할 PDF 선택",
                                             filetypes=[("PDF", "*.pdf")]):
            self.files.append(p)
            self.lb.insert("end", p)

    def remove(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.files[i]

    def move(self, d):
        sel = list(self.lb.curselection())
        if len(sel) != 1:
            return
        i, j = sel[0], sel[0] + d
        if not 0 <= j < len(self.files):
            return
        self.files[i], self.files[j] = self.files[j], self.files[i]
        txt_i, txt_j = self.lb.get(i), self.lb.get(j)
        self.lb.delete(i)
        self.lb.insert(i, txt_j)
        self.lb.delete(j)
        self.lb.insert(j, txt_i)
        self.lb.selection_set(j)

    def merge(self):
        if len(self.files) < 2:
            messagebox.showinfo(APP_TITLE, "2개 이상의 PDF를 추가하세요.", parent=self)
            return
        out = filedialog.asksaveasfilename(parent=self, defaultextension=".pdf",
                                           initialfile="병합.pdf", filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            new = fitz.open()
            for p in self.files:
                src = self.app.build_flat_doc() if p == CURRENT_DOC else self.app.open_pdf_file(p)
                if src is None:
                    return
                new.insert_pdf(src)
                src.close()
            save_doc(new, out, subset=True, inplace=True)
            new.close()
        except Exception as e:
            messagebox.showerror(APP_TITLE, "병합 실패:\n%s" % e, parent=self)
            return
        if messagebox.askyesno(APP_TITLE, "병합 완료. 지금 열어볼까요?", parent=self):
            self.destroy()
            self.app.open_path(out)


# ============================================================
#  메인 앱
# ============================================================
class PDFEditor(tk.Tk):
    def __init__(self, initial=None):
        super().__init__()
        self.title(APP_TITLE)
        self.ui = max(1.0, self.winfo_fpixels("1i") / 96.0)     # 화면 배율 (125% -> 1.25)
        self._fit_window()
        try:
            self.iconbitmap(resource_path("jp_pdf.ico"))
        except Exception:
            pass

        self.doc = None
        self.path = None
        self.file_kind = None          # 'pdf' | 'proj'
        self.modified = False
        self.undo_stack = []
        self.layers = []               # 페이지별 개체 리스트 (리스트 앞쪽 = 뒤, 뒤쪽 = 앞)
        self.sel_id = None
        self.geo = None
        self.cur = 0
        self.zoom = 1.0
        self.off = (PAD, PAD)
        self.page_img = None
        self.thumb_imgs = {}
        self._thumb_job = 0
        self._syncing = False
        self._syncing_layers = False
        self._loading = False
        self._prop_job = None
        self._drag = None
        self._clip = None
        self._paste_n = 0
        self._ed = None                # 페이지 위 텍스트 입력 중인 상태 (인라인 편집기)

        self.fonts = FONTS
        self.tool = tk.StringVar(value="select")
        self.shape_var = tk.StringVar(value="rect")      # '도형 그리기' 하위 도형
        self.weight_var = tk.StringVar(value=WEIGHT_NAMES[0])
        self.keep_tool = tk.BooleanVar(value=False)
        default_font = default_spec().label
        self.font_var = tk.StringVar(value=default_font)
        self.size_var = tk.IntVar(value=14)
        self.align_var = tk.StringVar(value="왼쪽")
        self.bg_none = tk.BooleanVar(value=True)
        self.fill_none = tk.BooleanVar(value=True)
        self.width_var = tk.DoubleVar(value=2.0)
        self.opacity_var = tk.IntVar(value=100)
        self.hl_opacity_var = tk.IntVar(value=40)
        self.colors = {"text": "#000000", "text_bg": "#FFFF99", "hl": "#FFFF00",
                       "stroke": "#FF0000", "fill": "#FFFFFF"}
        self.color_btns = {}

        self._build_menu()
        self._build_ui()
        self._bind_keys()
        for v in (self.font_var, self.size_var, self.align_var, self.weight_var, self.bg_none,
                  self.fill_none, self.width_var, self.opacity_var, self.hl_opacity_var):
            v.trace_add("write", self.on_prop_var)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.update_title()
        self.after(300, self._init_layout)
        if initial and os.path.isfile(initial):
            self.after(200, lambda: self.open_path(initial))

    def _fit_window(self):
        """첫 창 크기: 1420x900(배율 100% 기준)에 화면 배율을 곱하되,
        작업 표시줄을 뺀 화면에 다 들어가지 않으면 최대화해서 창 아래쪽이 화면 밖으로 나가지 않게 한다"""
        s = self.ui
        x0, y0, aw, ah = screen_work_area(self)
        fw, fh = int(20 * s), int(70 * s)               # 창 테두리 + 제목 표시줄 + 메뉴 여유
        w, h = int(1420 * s), int(900 * s)
        self.minsize(min(int(1040 * s), aw - fw), min(620, ah - fh))
        if w + fw <= aw and h + fh <= ah:
            self.geometry("%dx%d+%d+%d" % (w, h, x0 + (aw - w - fw) // 2, y0 + (ah - h - fh) // 2))
        else:
            self.geometry("%dx%d+%d+%d" % (min(w, aw - fw), min(h, ah - fh), x0, y0))
            try:
                self.state("zoomed")
            except tk.TclError:
                pass

    def _init_layout(self):
        """창이 뜬 뒤 오른쪽 패널 폭을 내용이 가로로 잘리지 않는 폭(기본 300, 화면 배율 반영) 이상으로 맞춘다.
        (예전에는 sash_place로 옮겼는데, Tk가 곧바로 최소 폭으로 되돌려 오른쪽 글자가 잘렸다)"""
        try:
            s = self.ui
            rw = max(int(300 * s), self.right_panel.content_width())
            if self.paned.winfo_width() - int(250 * s) - rw >= int(420 * s):
                self.paned.paneconfigure(self.right_panel, width=rw)
        except tk.TclError:
            pass

    # ---------------------------------------------------------- UI 구성
    def _build_menu(self):
        mb = tk.Menu(self)
        f = tk.Menu(mb, tearoff=0)
        f.add_command(label="새 PDF (빈 A4)", command=self.new_doc)
        f.add_command(label="열기…  (Ctrl+O)", command=self.open_dialog)
        f.add_separator()
        f.add_command(label="저장  (Ctrl+S)", command=self.save)
        f.add_command(label="PDF로 저장… (레이어 합침)", command=self.save_as_pdf)
        f.add_command(label="작업 파일(.pdfedit)로 저장… (레이어 유지)", command=self.save_project_as)
        f.add_separator()
        f.add_command(label="종료", command=self.on_close)
        mb.add_cascade(label="파일", menu=f)
        e = tk.Menu(mb, tearoff=0)
        e.add_command(label="실행 취소  (Ctrl+Z)", command=self.undo)
        e.add_separator()
        e.add_command(label="선택 개체 복사  (Ctrl+C)", command=self.copy_obj)
        e.add_command(label="붙여넣기 - 개체 또는 캡처 이미지  (Ctrl+V)", command=self.paste_obj)
        e.add_command(label="선택 개체 복제  (Ctrl+D)", command=self.duplicate_obj)
        e.add_command(label="선택 개체 삭제  (Delete)", command=self.delete_obj)
        mb.add_cascade(label="편집", menu=e)
        p = tk.Menu(mb, tearoff=0)
        p.add_command(label="PDF 병합 (여러 파일 → 1개)…", command=lambda: MergeDialog(self))
        p.add_command(label="분할 / 추출…", command=self.split_dialog)
        mb.add_cascade(label="도구", menu=p)
        h = tk.Menu(mb, tearoff=0)
        h.add_command(label="정보", command=self.show_about)
        mb.add_cascade(label="도움말", menu=h)
        self.config(menu=mb)

    def show_about(self):
        messagebox.showinfo(
            "정보", "%s  v%s\n\n설치 없이 실행하는 오프라인 PDF 편집기\n"
                    "· 텍스트(페이지 위 바로 입력·굵기)/형광펜/도형/캡처 이미지 레이어 편집\n· 페이지 분할·병합·삭제·추가·순서 변경\n\n"
                    "PDF 엔진: PyMuPDF" % (APP_TITLE, APP_VERSION), parent=self)

    def _build_ui(self):
        s = self.ui
        main = tk.PanedWindow(self, orient="horizontal", sashwidth=6, sashrelief="raised", bd=0)
        main.pack(fill="both", expand=True)
        self.paned = main

        # ---- 왼쪽: 썸네일 + 페이지 버튼
        left = ttk.Frame(main)
        main.add(left, minsize=int(210 * s), width=int(250 * s), stretch="never")
        style = ttk.Style()
        style.configure("Thumb.Treeview", rowheight=THUMB_H + 12,
                        background="#c9c9c9", fieldbackground="#c9c9c9")
        style.map("Thumb.Treeview", background=[("selected", "#3a6ea5")])
        # 페이지 버튼을 먼저 아래에 붙여야 썸네일 목록(요청 높이가 큼)에 밀려 사라지지 않는다
        pb = ttk.LabelFrame(left, text="페이지 편집 (Ctrl/Shift로 다중 선택)")
        pb.pack(side="bottom", fill="x", padx=4, pady=4)
        tf = ttk.Frame(left)
        tf.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tf, show="tree", selectmode="extended", style="Thumb.Treeview", height=3)
        self.tree.column("#0", width=200)
        tsb = ttk.Scrollbar(tf, command=self.tree.yview)
        self.tree.config(yscrollcommand=tsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tsb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        btns = [("삭제", self.delete_selected), ("회전 90°", self.rotate_selected),
                ("▲ 위로", lambda: self.move_selected(-1)), ("▼ 아래로", lambda: self.move_selected(1)),
                ("위치 이동…", self.move_to), ("분할/추출…", self.split_dialog),
                ("빈 페이지 추가", self.insert_blank), ("이미지 추가", self.insert_images),
                ("PDF 삽입(선택 뒤)", lambda: self.insert_pdfs("after")),
                ("PDF 붙이기(맨 끝)", lambda: self.insert_pdfs("end"))]
        for i, (t, c) in enumerate(btns):
            ttk.Button(pb, text=t, command=c).grid(row=i // 2, column=i % 2, sticky="ew", padx=2, pady=2)
        pb.columnconfigure(0, weight=1)
        pb.columnconfigure(1, weight=1)

        # ---- 가운데: 뷰어
        center = ttk.Frame(main)
        main.add(center, minsize=int(420 * s), stretch="always")
        bar1 = ttk.Frame(center)
        bar1.pack(fill="x", padx=4, pady=(3, 0))
        ttk.Button(bar1, text="열기", command=self.open_dialog).pack(side="left")
        ttk.Button(bar1, text="저장", command=self.save).pack(side="left", padx=2)
        ttk.Button(bar1, text="PDF로 저장 (레이어 합침)", command=self.save_as_pdf).pack(side="left", padx=2)
        ttk.Button(bar1, text="작업파일 저장 (레이어 유지)", command=self.save_project_as).pack(side="left", padx=2)
        ttk.Button(bar1, text="↶ 실행취소", command=self.undo).pack(side="left", padx=(10, 2))
        nav = ttk.Frame(center)
        nav.pack(fill="x", padx=4, pady=3)
        ttk.Button(nav, text="◀", width=3, command=lambda: self.goto(self.cur - 1)).pack(side="left")
        self.page_entry = ttk.Entry(nav, width=5, justify="center")
        self.page_entry.pack(side="left", padx=2)
        self.page_entry.bind("<Return>", self.on_page_entry)
        self.page_total = ttk.Label(nav, text="/ 0")
        self.page_total.pack(side="left")
        ttk.Button(nav, text="▶", width=3, command=lambda: self.goto(self.cur + 1)).pack(side="left", padx=2)
        ttk.Separator(nav, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(nav, text="－", width=3, command=lambda: self.set_zoom(self.zoom / 1.2)).pack(side="left")
        self.zoom_lbl = ttk.Label(nav, text="100%", width=6, anchor="center")
        self.zoom_lbl.pack(side="left")
        ttk.Button(nav, text="＋", width=3, command=lambda: self.set_zoom(self.zoom * 1.2)).pack(side="left")
        ttk.Button(nav, text="너비맞춤", command=self.fit_width).pack(side="left", padx=4)

        cf = ttk.Frame(center)
        cf.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(cf, bg="#6b6b6b", highlightthickness=0)
        vs = ttk.Scrollbar(cf, orient="vertical", command=self.canvas.yview)
        hs = ttk.Scrollbar(cf, orient="horizontal", command=self.canvas.xview)
        self.canvas.config(yscrollcommand=vs.set, xscrollcommand=hs.set)
        hs.pack(side="bottom", fill="x")
        vs.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Double-Button-1>", self.on_double)
        self.canvas.bind("<Motion>", self.on_motion)
        self.canvas.bind("<Configure>", lambda e: self.place_page())
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        self.canvas.bind("<Button-4>", lambda e: self.on_wheel(e, 1))
        self.canvas.bind("<Button-5>", lambda e: self.on_wheel(e, -1))
        self.canvas.bind("<ButtonPress-2>", lambda e: self.canvas.scan_mark(e.x, e.y))
        self.canvas.bind("<B2-Motion>", lambda e: self.canvas.scan_dragto(e.x, e.y, gain=1))
        sf = ttk.Frame(center)
        sf.pack(fill="x")
        self.coord_lbl = ttk.Label(sf, text="", width=22, anchor="e", relief="sunken")
        self.coord_lbl.pack(side="right")
        self.status = ttk.Label(sf, text="PDF를 열어주세요 (Ctrl+O)", anchor="w", relief="sunken")
        self.status.pack(side="left", fill="x", expand=True)

        # ---- 오른쪽: 도구 + (속성 / 레이어) - 화면이 작거나 배율이 크면 세로 스크롤
        self.right_panel = ScrollFrame(main, step=int(20 * s))
        main.add(self.right_panel, minsize=int(280 * s), width=int(300 * s), stretch="never")
        right = self.right_panel.body
        tl = ttk.LabelFrame(right, text="도구")
        tl.pack(fill="x", padx=6, pady=(6, 3))
        for text, val in [("선택 / 편집 (개체 이동·크기조절)", "select"), ("화면 이동 (드래그)", "hand"),
                          ("텍스트 추가", "text"), ("형광펜 (마커)", "highlight"),
                          ("도형 그리기", "shape")]:
            ttk.Radiobutton(tl, text=text, value=val, variable=self.tool,
                            command=self.on_tool_change).pack(anchor="w", padx=6, pady=0)
        sf = ttk.Frame(tl)                                # '도형 그리기' 하위 도형 선택
        sf.pack(fill="x", padx=(24, 4), pady=(0, 2))
        for i, (val, text) in enumerate(SHAPE_TYPES):
            ttk.Radiobutton(sf, text=text, value=val, variable=self.shape_var,
                            command=self.on_shape_pick).grid(row=i // 2, column=i % 2, sticky="w", padx=(0, 6))
        ttk.Checkbutton(tl, text="연속 그리기 (그린 후에도 도구 유지)", variable=self.keep_tool).pack(
            anchor="w", padx=6, pady=(2, 2))

        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True, padx=4, pady=3)
        prop = ttk.Frame(nb)
        lay = ttk.Frame(nb)
        nb.add(prop, text="  속성  ")
        nb.add(lay, text="  레이어  ")
        self._build_props(prop)
        self._build_layer_tab(lay)
        self.right_nb = nb
        self.right_panel.measure = self._right_need_h
        nb.bind("<<NotebookTabChanged>>", lambda e: self.right_panel.refresh(), add="+")
        self.right_panel.watch()

    def _right_need_h(self):
        """오른쪽 패널에 필요한 높이. 탭 묶음은 가장 큰 탭(속성) 대신 지금 보이는 탭 기준으로 센다
        (레이어 탭일 때 버튼이 괜히 스크롤 아래로 밀리지 않게)"""
        nb = self.right_nb
        if not nb.select():
            return self.right_panel.body.winfo_reqheight()
        tabs = [nb.nametowidget(t) for t in nb.tabs()]
        cur = nb.nametowidget(nb.select())
        return (self.right_panel.body.winfo_reqheight()
                - max(t.winfo_reqheight() for t in tabs) + cur.winfo_reqheight())

    def _build_props(self, parent):
        tip = ttk.Label(parent, text="※ 개체를 선택한 상태에서 값을 바꾸면 그 개체에 바로 적용됩니다.",
                        foreground="#0a5", wraplength=int(250 * self.ui), justify="left")
        tip.pack(anchor="w", padx=6, pady=(4, 0))
        tx = ttk.LabelFrame(parent, text="텍스트")
        tx.pack(fill="x", padx=4, pady=3)
        ttk.Label(tx, text="글꼴").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.font_cb = ttk.Combobox(tx, textvariable=self.font_var, state="readonly", width=20,
                                    values=list(self.fonts) + ["글꼴 파일 선택…"])
        self.font_cb.grid(row=0, column=1, columnspan=2, sticky="ew", pady=2)
        self.font_cb.bind("<<ComboboxSelected>>", self.on_font_select)
        ttk.Label(tx, text="크기(pt)").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        ttk.Spinbox(tx, from_=4, to=300, textvariable=self.size_var, width=6).grid(row=1, column=1, sticky="w")
        ttk.Label(tx, text="정렬").grid(row=2, column=0, sticky="w", padx=4, pady=2)
        ttk.Combobox(tx, textvariable=self.align_var, state="readonly", width=8,
                     values=["왼쪽", "가운데", "오른쪽"]).grid(row=2, column=1, sticky="w")
        ttk.Label(tx, text="굵기").grid(row=3, column=0, sticky="w", padx=4, pady=2)
        ttk.Combobox(tx, textvariable=self.weight_var, state="readonly", width=10,
                     values=WEIGHT_NAMES).grid(row=3, column=1, columnspan=2, sticky="w")
        ttk.Label(tx, text="글자색").grid(row=4, column=0, sticky="w", padx=4, pady=2)
        self._color_button(tx, "text", 4, 1)
        ttk.Label(tx, text="배경색").grid(row=5, column=0, sticky="w", padx=4, pady=2)
        self._color_button(tx, "text_bg", 5, 1)
        ttk.Checkbutton(tx, text="배경 없음", variable=self.bg_none).grid(row=5, column=2, sticky="w")

        hl = ttk.LabelFrame(parent, text="형광펜")
        hl.pack(fill="x", padx=4, pady=3)
        ttk.Label(hl, text="색상").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self._color_button(hl, "hl", 0, 1)
        ttk.Label(hl, text="불투명도(%)").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        ttk.Spinbox(hl, from_=10, to=100, increment=10, textvariable=self.hl_opacity_var,
                    width=5).grid(row=1, column=1, sticky="w")

        sh = ttk.LabelFrame(parent, text="도형")
        sh.pack(fill="x", padx=4, pady=3)
        ttk.Label(sh, text="선 색").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self._color_button(sh, "stroke", 0, 1)
        ttk.Label(sh, text="채우기").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        self._color_button(sh, "fill", 1, 1)
        ttk.Checkbutton(sh, text="채우기 없음", variable=self.fill_none).grid(row=1, column=2, sticky="w")
        ttk.Label(sh, text="선 두께").grid(row=2, column=0, sticky="w", padx=4, pady=2)
        ttk.Spinbox(sh, from_=0.5, to=20, increment=0.5, textvariable=self.width_var,
                    width=6).grid(row=2, column=1, sticky="w")
        ttk.Label(sh, text="불투명도(%)").grid(row=3, column=0, sticky="w", padx=4, pady=2)
        ttk.Spinbox(sh, from_=10, to=100, increment=10, textvariable=self.opacity_var,
                    width=6).grid(row=3, column=1, sticky="w")

        hint = ttk.Label(parent, foreground="#555", justify="left", wraplength=int(250 * self.ui),
                         text="[조작]\n· 텍스트: 클릭 후 바로 입력 (Ctrl+Enter·바깥 클릭=확정)\n"
                              "· 텍스트 핸들을 끌면 글자 크기 변경, 더블클릭=수정\n"
                              "· 선택: 드래그=이동, 핸들=크기, Delete=삭제\n"
                              "· 화면 캡처(Win+Shift+S) 후 Ctrl+V=이미지 붙여넣기\n"
                              "· 방향키=미세 이동 (Shift=10pt)")
        hint.pack(anchor="w", padx=6, pady=4)

    def _build_layer_tab(self, parent):
        ttk.Label(parent, text="현재 페이지의 레이어 (위 = 맨 앞)", foreground="#555").pack(
            anchor="w", padx=6, pady=(4, 0))
        bf = ttk.Frame(parent)                    # 버튼을 먼저 아래에 붙여 목록에 밀리지 않게
        bf.pack(side="bottom", fill="x", padx=4, pady=(0, 4))
        lf = ttk.Frame(parent)
        lf.pack(fill="both", expand=True, padx=4, pady=3)
        self.layer_tree = ttk.Treeview(lf, show="tree", selectmode="browse", height=6)
        ls = ttk.Scrollbar(lf, command=self.layer_tree.yview)
        self.layer_tree.config(yscrollcommand=ls.set)
        self.layer_tree.pack(side="left", fill="both", expand=True)
        ls.pack(side="left", fill="y")
        self.layer_tree.bind("<<TreeviewSelect>>", self.on_layer_select)
        btns = [("▲ 앞으로", lambda: self.reorder("up")), ("▼ 뒤로", lambda: self.reorder("down")),
                ("맨 앞으로", lambda: self.reorder("front")), ("맨 뒤로", lambda: self.reorder("back")),
                ("표시 / 숨김", self.toggle_visible), ("이름 변경…", self.rename_obj),
                ("내용 수정…", self.edit_text_selected), ("복제", self.duplicate_obj),
                ("복사", self.copy_obj), ("붙여넣기", self.paste_obj),
                ("삭제", self.delete_obj), ("모든 페이지에 복사", self.copy_to_all_pages)]
        for i, (t, c) in enumerate(btns):
            ttk.Button(bf, text=t, command=c).grid(row=i // 2, column=i % 2, sticky="ew", padx=2, pady=2)
        bf.columnconfigure(0, weight=1)
        bf.columnconfigure(1, weight=1)

    def _color_button(self, parent, key, row, col):
        b = tk.Button(parent, width=4, bg=self.colors[key], activebackground=self.colors[key],
                      relief="groove", command=lambda: self.pick_color(key))
        b.grid(row=row, column=col, sticky="w", pady=2)
        self.color_btns[key] = b

    def pick_color(self, key):
        c = colorchooser.askcolor(color=self.colors[key], parent=self, title="색상 선택")[1]
        if not c:
            return
        self.colors[key] = c
        self.color_btns[key].config(bg=c, activebackground=c)
        if key == "text_bg":
            self.bg_none.set(False)
        if key == "fill":
            self.fill_none.set(False)
        self.on_prop_var()

    def rgb(self, key):
        return hex_to_rgb(self.colors[key])

    def on_font_select(self, _e=None):
        if self.font_var.get() != "글꼴 파일 선택…":
            return
        p = filedialog.askopenfilename(parent=self, title="글꼴 파일 선택",
                                       filetypes=[("글꼴", "*.ttf *.ttc *.otf"), ("모든 파일", "*.*")])
        if p:
            name = os.path.splitext(os.path.basename(p))[0]
            FONTS[name] = FontSpec(name, "file", p)
            self.font_cb.config(values=list(self.fonts) + ["글꼴 파일 선택…"])
            self.font_var.set(name)
        else:
            self.font_var.set(default_spec().label)

    def _bind_keys(self):
        self.bind_all("<Control-o>", lambda e: self.open_dialog())
        self.bind_all("<Control-s>", lambda e: self.save())
        self.bind_all("<Control-z>", self._key_undo)
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.bind("<Prior>", lambda e: self.goto(self.cur - 1))
        self.bind("<Next>", lambda e: self.goto(self.cur + 1))
        c = self.canvas
        c.bind("<Delete>", lambda e: self.delete_obj())
        c.bind("<Escape>", lambda e: self.select_obj(None))
        c.bind("<Control-c>", lambda e: self.copy_obj())
        self.bind_all("<Control-v>", self._key_paste)      # 캡처 후 돌아왔을 때 초점이 어디 있든
        self.bind_all("<Control-V>", self._key_paste)
        c.bind("<Control-d>", lambda e: self.duplicate_obj())
        for key, dx, dy in (("Left", -1, 0), ("Right", 1, 0), ("Up", 0, -1), ("Down", 0, 1)):
            c.bind("<%s>" % key, lambda e, dx=dx, dy=dy: self.nudge(dx, dy, 1))
            c.bind("<Shift-%s>" % key, lambda e, dx=dx, dy=dy: self.nudge(dx, dy, 10))

    def _key_paste(self, e):
        try:
            w = self.focus_get()
        except (KeyError, tk.TclError):
            return
        if isinstance(w, (tk.Text, tk.Entry, ttk.Entry, ttk.Spinbox, ttk.Combobox)):
            return                      # 입력창 안에서는 글자 붙여넣기
        if isinstance(w, tk.Toplevel) or (w is not None and w.winfo_toplevel() is not self):
            return                      # 다른 창(병합·분할 대화상자 등)
        self.paste_obj()
        return "break"

    def _key_undo(self, e):
        w = self.focus_get()
        if isinstance(w, (tk.Text, tk.Entry, ttk.Entry, ttk.Spinbox, ttk.Combobox, tk.Listbox)):
            return                      # 입력창 안에서는 그 입력창의 동작에 맡긴다
        self.undo()

    def set_status(self, msg):
        self.status.config(text=msg)

    def update_title(self):
        name = os.path.basename(self.path) if self.path else "(제목 없음)"
        self.title("%s%s - %s" % ("*" if self.modified else "", name, APP_TITLE))

    # ---------------------------------------------------------- 좌표 변환
    def v2c(self, x, y):
        return self.off[0] + x * self.zoom, self.off[1] + y * self.zoom

    def c2v(self, cx, cy):
        return fitz.Point((cx - self.off[0]) / self.zoom, (cy - self.off[1]) / self.zoom)

    def c2v_clamped(self, cx, cy):
        p = self.c2v(cx, cy)
        return fitz.Point(min(max(p.x, 0), self.geo.vw), min(max(p.y, 0), self.geo.vh))

    def p2c(self, pt):
        v = fitz.Point(pt) * self.geo.rotm
        return self.v2c(v.x, v.y)

    def obj_rect_v(self, o):
        return rect_p2v(bbox_p(o, self.geo), self.geo)

    # ---------------------------------------------------------- 파일
    def confirm_discard(self):
        self.finish_inline_edit()
        if self.doc is not None and self.modified:
            r = messagebox.askyesnocancel(APP_TITLE, "저장하지 않은 변경 사항이 있습니다. 저장할까요?")
            if r is None:
                return False
            if r:
                return self.save()
        return True

    def open_pdf_bytes(self, data, name):
        try:
            d = fitz.open(stream=data, filetype="pdf")
        except Exception as e:
            messagebox.showerror(APP_TITLE, "파일을 열 수 없습니다:\n%s" % e)
            return None
        if d.needs_pass:
            for _ in range(3):
                pw = simpledialog.askstring(APP_TITLE, "%s\n비밀번호를 입력하세요:" % name,
                                            show="*", parent=self)
                if pw is None:
                    return None
                if d.authenticate(pw):
                    break
            else:
                messagebox.showerror(APP_TITLE, "비밀번호가 맞지 않습니다.")
                return None
        return d

    def open_pdf_file(self, path):
        """스트림으로 열어 원본 파일 잠금을 피한다"""
        try:
            with open(path, "rb") as f:
                data = f.read()
        except Exception as e:
            messagebox.showerror(APP_TITLE, "파일을 열 수 없습니다:\n%s" % e)
            return None
        return self.open_pdf_bytes(data, os.path.basename(path))

    def open_dialog(self):
        if not self.confirm_discard():
            return
        p = filedialog.askopenfilename(
            parent=self, title="PDF 또는 작업 파일 열기",
            filetypes=[("PDF / 작업 파일", "*.pdf *.pdfedit"), ("PDF", "*.pdf"),
                       ("작업 파일", "*.pdfedit")])
        if p:
            self.open_path(p, checked=True)

    def open_path(self, path, checked=False):
        if not checked and not self.confirm_discard():
            return
        if path.lower().endswith(PROJ_EXT):
            self.open_project(path)
            return
        d = self.open_pdf_file(path)
        if d is not None:
            self.set_doc(d, path, "pdf", [[] for _ in range(len(d))])

    def open_project(self, path):
        try:
            with zipfile.ZipFile(path) as z:
                base = z.read("base.pdf")
                lay = json.loads(z.read("layers.json").decode("utf-8"))
                for n in z.namelist():                  # 이미지 레이어의 그림
                    if n.startswith("images/") and n.endswith(".png"):
                        IMAGE_STORE[n[7:-4]] = z.read(n)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "작업 파일을 열 수 없습니다:\n%s" % e)
            return
        d = self.open_pdf_bytes(base, os.path.basename(path))
        if d is None:
            return
        layers = [[Obj.from_dict(x) for x in pg] for pg in lay.get("pages", [])]
        self.set_doc(d, path, "proj", layers)

    def new_doc(self):
        if not self.confirm_discard():
            return
        d = fitz.open()
        d.new_page(width=595, height=842)
        self.set_doc(d, None, None, [[]])
        self.modified = True
        self.update_title()

    def set_doc(self, d, path, kind, layers):
        self.doc, self.path, self.file_kind = d, path, kind
        self.layers = layers
        self.ensure_layers()
        self.sel_id = None
        self.modified = False
        self.undo_stack.clear()
        self.cur = 0
        self.refresh_thumbs(select=[0])
        self.after(50, self.fit_width)
        self.update_title()
        n_obj = sum(len(p) for p in self.layers)
        self.set_status("열림: %s  (%d쪽%s)" % (path or "새 문서", len(d),
                                              ", 레이어 개체 %d개" % n_obj if n_obj else ""))

    def ensure_layers(self):
        n = len(self.doc)
        while len(self.layers) < n:
            self.layers.append([])
        del self.layers[n:]

    # ---- 합치기 / 저장
    def has_file_fonts(self, layers=None):
        return any(o.kind == "text" and spec_for(o).kind == "file"
                   for pg in (layers if layers is not None else self.layers) for o in pg)

    def build_flat_doc(self):
        """편집 중인 문서 + 레이어 개체를 합친 새 PDF (원본은 그대로)"""
        out = fitz.open("pdf", self.doc.tobytes())
        for i, objs in enumerate(self.layers):
            apply_objects(out[i], objs)
        return out

    def write_pdf(self, path):
        out = self.build_flat_doc()
        try:
            save_doc(out, path, subset=self.has_file_fonts(), inplace=True)
        finally:
            out.close()

    def write_project(self, path):
        data = self.doc.tobytes(garbage=3, deflate=True)
        lay = {"version": 2, "pages": [[o.to_dict() for o in pg] for pg in self.layers]}
        tmp = path + ".tmp"
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("base.pdf", data)
                z.writestr("layers.json", json.dumps(lay, ensure_ascii=False))
                for key in sorted({o.img for pg in self.layers for o in pg if o.kind == "image"}):
                    if key in IMAGE_STORE:
                        z.writestr("images/%s.png" % key, IMAGE_STORE[key], zipfile.ZIP_STORED)
            os.replace(tmp, path)
        except PermissionError:
            _silent_remove(tmp)
            raise PermissionError("'%s' 파일이 다른 프로그램에서 열려 있어 저장할 수 없습니다." % os.path.basename(path))
        except Exception:
            _silent_remove(tmp)
            raise

    def n_objects(self):
        return sum(len(p) for p in self.layers)

    def save(self):
        self.finish_inline_edit()
        if self.doc is None:
            return False
        if not self.path:
            return self.save_as_pdf()
        try:
            if self.file_kind == "proj":
                self.write_project(self.path)
                msg = "작업 파일 저장 완료: %s  (레이어 유지)" % self.path
            else:
                self.write_pdf(self.path)
                msg = "PDF 저장 완료: %s" % self.path
                if self.n_objects():
                    msg += "  (개체는 PDF에 합쳐짐 · 나중에 수정하려면 '작업파일 저장')"
        except Exception as e:
            messagebox.showerror(APP_TITLE, "저장 실패:\n%s" % e)
            return False
        self.modified = False
        self.update_title()
        self.set_status(msg)
        return True

    def save_as_pdf(self):
        self.finish_inline_edit()
        if self.doc is None:
            return False
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "새문서"
        init = base + ("_편집.pdf" if self.path and self.file_kind == "pdf" else ".pdf")
        p = filedialog.asksaveasfilename(parent=self, defaultextension=".pdf", initialfile=init,
                                         filetypes=[("PDF", "*.pdf")])
        if not p:
            return False
        try:
            self.write_pdf(p)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "저장 실패:\n%s" % e)
            return False
        if self.file_kind == "proj":          # 작업 파일 편집 중이면 '내보내기'로만 동작
            self.set_status("PDF 내보내기 완료: %s  (작업 파일은 그대로)" % p)
            return True
        self.path, self.file_kind = p, "pdf"
        self.modified = False
        self.update_title()
        self.set_status("PDF 저장 완료: %s" % p + ("  (개체는 PDF에 합쳐짐)" if self.n_objects() else ""))
        return True

    def save_project_as(self):
        self.finish_inline_edit()
        if self.doc is None:
            return False
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "작업"
        p = filedialog.asksaveasfilename(parent=self, defaultextension=PROJ_EXT, initialfile=base + PROJ_EXT,
                                         filetypes=[("PDF 편집 작업 파일", "*" + PROJ_EXT)])
        if not p:
            return False
        try:
            self.write_project(p)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "저장 실패:\n%s" % e)
            return False
        self.path, self.file_kind = p, "proj"
        self.modified = False
        self.update_title()
        self.set_status("작업 파일 저장 완료: %s  (레이어 유지)" % p)
        return True

    def on_close(self):
        if self.confirm_discard():
            self.destroy()

    # ---------------------------------------------------------- 실행 취소
    def snapshot(self, with_doc=False):
        return {"doc": self.doc.tobytes() if with_doc else None,
                "layers": [[o.to_dict() for o in pg] for pg in self.layers],
                "cur": self.cur, "sel": self.sel_id}

    def push_undo(self, structural=False):
        """페이지 구조를 바꾸는 작업 직전에 호출 (문서 사본 포함)"""
        self.undo_stack.append(self.snapshot(with_doc=structural))
        self._trim_undo()
        self.modified = True
        self.update_title()

    def _trim_undo(self):
        while len(self.undo_stack) > UNDO_LIMIT:
            self.undo_stack.pop(0)

    def commit(self, pre, msg=None):
        """레이어 개체 변경 작업 확정: pre = 변경 전 snapshot(with_doc=False)"""
        self.undo_stack.append(pre)
        self._trim_undo()
        self.modified = True
        self.update_title()
        self.after_obj_change(msg)

    def after_obj_change(self, msg=None):
        self.update_thumb(self.cur)
        self.draw_overlays()
        self.refresh_layer_list()
        if msg:
            self.set_status(msg)

    def undo(self, quiet=False):
        if self._ed is not None:            # 입력 중이면 입력만 취소
            self.finish_inline_edit(cancel=True)
            return
        if not self.undo_stack or self.doc is None:
            return
        s = self.undo_stack.pop()
        if s["doc"] is not None:
            self.doc = fitz.open("pdf", s["doc"])
        self.layers = [[Obj.from_dict(d) for d in pg] for pg in s["layers"]]
        self.ensure_layers()
        self.cur = min(s["cur"], len(self.doc) - 1)
        self.sel_id = s["sel"]
        self.modified = True
        self.update_title()
        if s["doc"] is not None:
            self.refresh_thumbs(select=[self.cur])
        else:
            self.update_thumb(self.cur)
            self.render()
            self.refresh_layer_list()
        if not quiet:
            self.set_status("실행 취소됨 (남은 취소 가능 횟수: %d)" % len(self.undo_stack))

    # ---------------------------------------------------------- 썸네일
    def refresh_thumbs(self, select=None):
        self._thumb_job += 1
        job = self._thumb_job
        self.tree.delete(*self.tree.get_children())
        self.thumb_imgs = {}
        if self.doc is None:
            self.render()
            self.refresh_layer_list()
            return
        n = len(self.doc)
        for i in range(n):
            self.tree.insert("", "end", iid=str(i), text="  %d" % (i + 1))
        if select:
            select = [i for i in select if 0 <= i < n] or [0]
            self._syncing = True
            self.tree.selection_set([str(i) for i in select])
            self.tree.focus(str(select[-1]))
            self.tree.see(str(select[-1]))
            self._syncing = False
            self.cur = select[-1]
        self.page_total.config(text="/ %d" % n)
        self.render()
        self.refresh_layer_list()
        self.after(10, lambda: self._fill_thumbs(0, job))

    def _fill_thumbs(self, start, job):
        if job != self._thumb_job or self.doc is None:
            return
        n = len(self.doc)
        for i in range(start, min(start + 4, n)):
            self.thumb_imgs[i] = self.make_thumb(i)
            if self.tree.exists(str(i)):
                self.tree.item(str(i), image=self.thumb_imgs[i])
        if start + 4 < n:
            self.after(5, lambda: self._fill_thumbs(start + 4, job))

    def make_thumb(self, i):
        objs = self.layers[i] if i < len(self.layers) else []
        tmp = None
        page = self.doc[i]
        if any(o.visible for o in objs):
            tmp = fitz.open()
            tmp.insert_pdf(self.doc, from_page=i, to_page=i)
            page = tmp[0]
            try:
                apply_objects(page, objs)
            except Exception:
                pass
        r = page.rect
        s = min(THUMB_W / r.width, THUMB_H / r.height)
        pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
        img = tk.PhotoImage(data=pix.tobytes("ppm"))
        if tmp is not None:
            tmp.close()
        return img

    def update_thumb(self, i):
        if self.doc is None or not self.tree.exists(str(i)):
            return
        self.thumb_imgs[i] = self.make_thumb(i)
        self.tree.item(str(i), image=self.thumb_imgs[i])

    def selected_pages(self):
        return sorted(int(i) for i in self.tree.selection())

    def on_tree_select(self, _e=None):
        if self._syncing or self.doc is None:
            return
        sel = self.tree.selection()
        if not sel:
            return
        f = self.tree.focus()
        idx = int(f) if f in sel else int(sel[-1])
        if idx != self.cur:
            self.show_page(idx, sync=False)

    # ---------------------------------------------------------- 뷰어
    def goto(self, i):
        if self.doc is None or not 0 <= i < len(self.doc):
            return
        self.show_page(i, sync=True)

    def on_page_entry(self, _e=None):
        try:
            self.goto(int(self.page_entry.get()) - 1)
        except ValueError:
            pass

    def show_page(self, i, sync=True):
        self.finish_inline_edit()
        self.cur = i
        self.sel_id = None
        if sync:
            self._syncing = True
            self.tree.selection_set(str(i))
            self.tree.focus(str(i))
            self.tree.see(str(i))
            self._syncing = False
        self.render()
        self.refresh_layer_list()
        self.canvas.yview_moveto(0)

    def render(self):
        self.finish_inline_edit()
        self.canvas.delete("all")
        self.page_img = None
        if self.doc is None:
            return
        self.ensure_layers()
        self.page_entry.delete(0, "end")
        self.page_entry.insert(0, str(self.cur + 1))
        page = self.doc[self.cur]
        self.geo = PageGeo(page)
        z = self.zoom
        pix = page.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
        self.page_img = tk.PhotoImage(data=pix.tobytes("ppm"))
        self.img_w, self.img_h = pix.width, pix.height
        self.canvas.create_rectangle(0, 0, 1, 1, tags="shadow", outline="", fill="#333")
        self.img_item = self.canvas.create_image(0, 0, anchor="nw", image=self.page_img, tags="page")
        self.place_page()
        self.zoom_lbl.config(text="%d%%" % round(z * 100))

    def place_page(self):
        if self.doc is None or self.page_img is None:
            return
        cw = self.canvas.winfo_width()
        ox = max(PAD, (cw - self.img_w) // 2)
        oy = PAD
        self.off = (ox, oy)
        self.canvas.coords(self.img_item, ox, oy)
        self.canvas.coords("shadow", ox + 4, oy + 4, ox + self.img_w + 4, oy + self.img_h + 4)
        self.canvas.config(scrollregion=(0, 0, ox + self.img_w + PAD, oy + self.img_h + PAD))
        self.draw_overlays()

    # ---- 레이어 개체 그리기
    def overlay_image(self, o):
        key = (round(self.zoom, 4), self.geo.rot, o.sig())
        if o._img_key == key:
            return o._img
        res = None
        try:
            res = render_obj_png(o, self.geo, self.zoom)
        except Exception as e:
            self.set_status("개체 표시 실패: %s" % e)
        o._img = (tk.PhotoImage(data=res[0]), res[1], res[2]) if res else None
        o._img_key = key
        return o._img

    def draw_overlays(self):
        if self.doc is None or self.page_img is None:
            return
        self.canvas.delete("ov")
        ox, oy = self.off
        for o in self.layers[self.cur]:
            if not o.visible:
                continue
            img = self.overlay_image(o)
            if img:
                self.canvas.create_image(ox + img[1], oy + img[2], anchor="nw", image=img[0],
                                         tags=("ov", "o%d" % o.id))
        self.draw_selection()

    def sel_obj(self):
        if self.sel_id is None or self.doc is None:
            return None
        for o in self.layers[self.cur]:
            if o.id == self.sel_id:
                return o
        return None

    def handles(self, o):
        if o.kind in LINE_KINDS:
            a, b = self.p2c(o.p0), self.p2c(o.p1)
            return [("p0", a[0], a[1]), ("p1", b[0], b[1])]
        r = self.obj_rect_v(o)
        x0, y0 = self.v2c(r.x0, r.y0)
        x1, y1 = self.v2c(r.x1, r.y1)
        xm, ym = (x0 + x1) / 2, (y0 + y1) / 2
        return [("nw", x0, y0), ("n", xm, y0), ("ne", x1, y0), ("e", x1, ym),
                ("se", x1, y1), ("s", xm, y1), ("sw", x0, y1), ("w", x0, ym)]

    def draw_selection(self):
        self.canvas.delete("sel")
        o = self.sel_obj()
        if o is None or not o.visible:
            return
        hs = self.handles(o)
        if o.kind in LINE_KINDS:
            self.canvas.create_line(hs[0][1], hs[0][2], hs[1][1], hs[1][2], fill="#0078d7",
                                    dash=(4, 3), tags="sel")
        else:
            r = self.obj_rect_v(o)
            x0, y0 = self.v2c(r.x0, r.y0)
            x1, y1 = self.v2c(r.x1, r.y1)
            self.canvas.create_rectangle(x0, y0, x1, y1, outline="#0078d7", dash=(4, 3), tags="sel")
        for _n, x, y in hs:
            self.canvas.create_rectangle(x - HANDLE_R, y - HANDLE_R, x + HANDLE_R, y + HANDLE_R,
                                         fill="white", outline="#0078d7", width=2, tags="sel")

    def handle_at(self, cx, cy):
        o = self.sel_obj()
        if o is None or not o.visible:
            return None
        for name, x, y in reversed(self.handles(o)):
            if abs(cx - x) <= HANDLE_R + 3 and abs(cy - y) <= HANDLE_R + 3:
                return name
        return None

    def hit_test(self, cx, cy):
        pt = self.c2v(cx, cy) * self.geo.derotm
        tol = 5.0 / self.zoom
        for o in reversed(self.layers[self.cur]):
            if o.visible and obj_hit(o, pt, tol, self.geo):
                return o
        return None

    def set_zoom(self, z):
        self.zoom = max(0.1, min(8.0, z))
        self.render()

    def fit_width(self):
        if self.doc is None:
            return
        cw = max(self.canvas.winfo_width(), 300)
        self.set_zoom((cw - 2 * PAD - 24) / self.doc[self.cur].rect.width)

    def on_wheel(self, e, direction=None):
        d = direction if direction is not None else (1 if e.delta > 0 else -1)
        if e.state & 0x4:                                # Ctrl + 휠 = 확대/축소
            self.set_zoom(self.zoom * (1.1 if d > 0 else 1 / 1.1))
        elif e.state & 0x1:                              # Shift + 휠 = 좌우
            self.canvas.xview_scroll(-d * 3, "units")
        else:
            self.canvas.yview_scroll(-d * 3, "units")

    def active_tool(self):
        """실제로 쓸 도구. '도형 그리기'면 그 아래에서 고른 도형 종류"""
        t = self.tool.get()
        return self.shape_var.get() if t == "shape" else t

    def on_shape_pick(self):
        self.tool.set("shape")
        self.on_tool_change()

    def on_tool_change(self):
        self.finish_inline_edit()
        t = self.active_tool()
        if t != "select":
            self.select_obj(None)        # 새 개체를 만들 때 이전 선택 개체가 바뀌지 않도록
        self.canvas.config(cursor={"select": "arrow", "hand": "fleur", "text": "xterm"}.get(t, "crosshair"))
        hints = {"select": "개체를 클릭해 선택 · 드래그로 이동 · 네모 핸들로 크기 조절 · 텍스트는 더블클릭으로 수정",
                 "hand": "드래그하여 화면 이동",
                 "text": "글을 넣을 자리를 클릭하고 바로 입력 (Enter=줄바꿈, Ctrl+Enter·바깥 클릭=확정, Esc=취소)",
                 "highlight": "형광펜: 드래그 (글자 위면 글자에 맞춰짐)"}
        self.set_status(hints.get(t) or "도형 그리기 - %s: 드래그하여 그리기" % KIND_LABEL[t])

    def on_motion(self, e):
        if self.doc is None or self.page_img is None:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        p = self.c2v(cx, cy)
        self.coord_lbl.config(text="x=%.0f  y=%.0f  (pt)" % (p.x, p.y))
        if self.tool.get() == "select" and self._drag is None:
            cur = "arrow"
            if self.handle_at(cx, cy):
                cur = "sizing"
            elif self.hit_test(cx, cy):
                cur = "fleur"
            self.canvas.config(cursor=cur)

    # ---------------------------------------------------------- 마우스 편집
    def on_press(self, e):
        if self.doc is None:
            return
        if self._ed is not None:          # 입력 중에 바깥을 클릭 -> 입력 확정만 하고 끝
            self.finish_inline_edit()
            self._drag = None
            self.canvas.focus_set()
            return
        self.canvas.focus_set()
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        self._drag = {"start": (cx, cy), "last": (cx, cy), "mode": None}
        t = self.tool.get()
        if t == "hand":
            self.canvas.scan_mark(e.x, e.y)
            self._drag["mode"] = "pan"
            return
        if t != "select":
            self._drag["mode"] = "create"
            return
        o = self.sel_obj()
        h = self.handle_at(cx, cy) if o else None
        if o and h:
            self._drag.update(mode="resize", handle=h, obj=o, pre=self.snapshot(False),
                              rv0=self.obj_rect_v(o) if o.kind not in LINE_KINDS else None)
            return
        hit = self.hit_test(cx, cy)
        if hit:
            self.select_obj(hit)
            self._drag.update(mode="move", obj=hit, pre=self.snapshot(False))
        else:
            self.select_obj(None)
            self._drag = None

    def on_drag(self, e):
        d = self._drag
        if self.doc is None or d is None:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        m = d["mode"]
        if m == "pan":
            self.canvas.scan_dragto(e.x, e.y, gain=1)
            return
        if m == "move":
            ddx, ddy = cx - d["last"][0], cy - d["last"][1]
            o = d["obj"]
            self.canvas.move("o%d" % o.id, ddx, ddy)
            self.canvas.move("sel", ddx, ddy)
            d["last"] = (cx, cy)
            return
        if m == "resize":
            self.canvas.delete("tmp")
            o, h = d["obj"], d["handle"]
            pv = self.c2v(cx, cy)
            if o.kind in LINE_KINDS:
                other = self.p2c(o.p1 if h == "p0" else o.p0)
                self.canvas.create_line(other[0], other[1], cx, cy, fill="#0078d7", dash=(4, 3), tags="tmp")
            else:
                r = self.resized_rect_v(d["rv0"], h, pv, keep_ratio=(o.kind == "image"))
                x0, y0 = self.v2c(r.x0, r.y0)
                x1, y1 = self.v2c(r.x1, r.y1)
                self.canvas.create_rectangle(x0, y0, x1, y1, outline="#0078d7", dash=(4, 3), tags="tmp")
            return
        if m == "create":
            t = self.active_tool()
            x0, y0 = d["start"]
            self.canvas.delete("tmp")
            if t == "text":
                return
            if t == "oval":
                self.canvas.create_oval(x0, y0, cx, cy, outline="#0078d7", width=2, tags="tmp")
            elif t == "triangle":
                self.canvas.create_polygon((x0 + cx) / 2, min(y0, cy), max(x0, cx), max(y0, cy),
                                           min(x0, cx), max(y0, cy), outline="#0078d7", fill="",
                                           width=2, tags="tmp")
            elif t in LINE_KINDS:
                self.canvas.create_line(x0, y0, cx, cy, fill="#0078d7", width=2, tags="tmp",
                                        arrow={"arrow": "last", "darrow": "both"}.get(t))
            else:
                self.canvas.create_rectangle(x0, y0, cx, cy, outline="#0078d7", width=2,
                                             dash=(4, 3), tags="tmp")

    def on_release(self, e):
        d = self._drag
        self._drag = None
        if self.doc is None or d is None:
            return
        self.canvas.delete("tmp")
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        m = d["mode"]
        if m == "pan" or m is None:
            return
        x0, y0 = d["start"]
        moved = abs(cx - x0) > 3 or abs(cy - y0) > 3
        if m == "move":
            o = d["obj"]
            if abs(cx - x0) < 0.5 and abs(cy - y0) < 0.5:
                self.draw_overlays()
                return
            dp = self.delta_v2p((cx - x0) / self.zoom, (cy - y0) / self.zoom)
            o.shift(dp[0], dp[1])
            self.commit(d["pre"], "개체 이동")
            return
        if m == "resize":
            o = d["obj"]
            self.apply_resize(o, d["handle"], d["rv0"], self.c2v(cx, cy))
            self.commit(d["pre"], "크기 조절")
            return
        # 생성
        t = self.active_tool()
        p0, p1 = self.c2v_clamped(x0, y0), self.c2v_clamped(cx, cy)
        if t == "text":
            self.create_text(p0, p1)
        elif not moved:
            return
        elif t == "highlight":
            self.create_highlight(p0, p1)
        else:
            self.create_shape(t, p0, p1)

    def on_double(self, e):
        if self.doc is None or self.tool.get() != "select":
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        o = self.hit_test(cx, cy)
        if o is not None and o.kind == "text":
            self._drag = None
            self.select_obj(o)
            self.edit_text(o)

    def delta_v2p(self, dvx, dvy):
        a = fitz.Point(0, 0) * self.geo.derotm
        b = fitz.Point(dvx, dvy) * self.geo.derotm
        return (b.x - a.x, b.y - a.y)

    def resized_rect_v(self, rv0, h, pv, keep_ratio=False):
        if keep_ratio and len(h) == 2 and rv0.width > 0 and rv0.height > 0:
            # 모서리 핸들: 반대쪽 모서리를 고정하고 가로세로 비율 유지
            ax = rv0.x1 if "w" in h else rv0.x0
            ay = rv0.y1 if "n" in h else rv0.y0
            s = max(abs(pv.x - ax) / rv0.width, abs(pv.y - ay) / rv0.height, 6 / min(rv0.width, rv0.height))
            w, hh = rv0.width * s, rv0.height * s
            x0 = ax - w if "w" in h else ax
            y0 = ay - hh if "n" in h else ay
            return fitz.Rect(x0, y0, x0 + w, y0 + hh)
        x0, y0, x1, y1 = rv0.x0, rv0.y0, rv0.x1, rv0.y1
        if "w" in h:
            x0 = pv.x
        if "e" in h:
            x1 = pv.x
        if "n" in h:
            y0 = pv.y
        if "s" in h:
            y1 = pv.y
        r = fitz.Rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
        if r.width < 6:
            r.x1 = r.x0 + 6
        if r.height < 6:
            r.y1 = r.y0 + 6
        return r

    def apply_resize(self, o, h, rv0, pv):
        if o.kind in LINE_KINDS:
            p = fitz.Point(pv) * self.geo.derotm
            if h == "p0":
                o.p0 = (p.x, p.y)
            else:
                o.p1 = (p.x, p.y)
            return
        r = self.resized_rect_v(rv0, h, pv, keep_ratio=(o.kind == "image"))
        if o.kind == "text":
            self.scale_text(o, h, rv0, r)
            return
        newp = rect_v2p(r, self.geo)
        if o.kind == "highlight":
            oldp = rect_v2p(rv0, self.geo)
            sx = newp.width / oldp.width if oldp.width else 1.0
            sy = newp.height / oldp.height if oldp.height else 1.0
            o.rects = [(newp.x0 + (a - oldp.x0) * sx, newp.y0 + (b - oldp.y0) * sy,
                        newp.x0 + (c - oldp.x0) * sx, newp.y0 + (d - oldp.y0) * sy)
                       for a, b, c, d in o.rects]
        else:
            o.rect = tuple(newp)

    def scale_text(self, o, h, rv0, r):
        """텍스트 상자 크기 조절 = 글자 크기 조절. 잡은 핸들 반대편 모서리는 제자리에 둔다"""
        sw = r.width / rv0.width if rv0.width else 1.0
        sh = r.height / rv0.height if rv0.height else 1.0
        s = sw if h in ("e", "w") else sh if h in ("n", "s") else (sw + sh) / 2
        o.size = float(min(300, max(4, round(o.size * s))))
        w = rv0.width * s if not o.auto_w else 10.0        # 줄바꿈 폭이 정해진 글은 폭도 같은 비율로
        rv = fitz.Rect(rv0.x0, rv0.y0, rv0.x0 + w, rv0.y0 + 10)
        o.rect = tuple(rect_v2p(rv, self.geo))
        lay = text_layout(o, self.geo)
        dx = rv0.x1 - lay.x1 if "w" in h else 0.0
        dy = rv0.y1 - lay.y1 if "n" in h else 0.0
        o.rect = tuple(rect_v2p(fitz.Rect(rv.x0 + dx, rv.y0 + dy, rv.x1 + dx, rv.y1 + dy), self.geo))
        self.load_props(o)

    def nudge(self, dx, dy, step):
        o = self.sel_obj()
        if o is None:
            return "break"
        pre = self.snapshot(False)
        dp = self.delta_v2p(dx * step, dy * step)
        o.shift(dp[0], dp[1])
        self.commit(pre, None)
        return "break"

    # ---------------------------------------------------------- 개체 생성
    def current_font_spec(self):
        return self.fonts.get(self.font_var.get()) or default_spec()

    @staticmethod
    def set_obj_font(o, spec):
        o.font_kind, o.font_value, o.font_label = spec.kind, spec.value, spec.label

    def resolve_font(self, spec, text):
        """글꼴이 글자를 지원하지 않으면 한글 글꼴로 대체. 실패 시 None"""
        if spec.supports(text):
            return spec
        fb = korean_fallback()
        if fb is not None and fb.supports(text):
            self.set_status("선택한 글꼴에 없는 글자가 있어 '%s'(으)로 대체했습니다." % fb.label)
            return fb
        messagebox.showwarning(APP_TITLE, "선택한 글꼴이 입력한 글자를 지원하지 않습니다.\n"
                               "다른 글꼴(예: 맑은 고딕)을 선택하세요.")
        return None

    def add_obj(self, o, msg):
        pre = self.snapshot(False)
        self.layers[self.cur].append(o)
        self.sel_id = o.id
        self.commit(pre, msg)
        self.load_props(o)
        self.sync_layer_selection()
        if not self.keep_tool.get():
            self.tool.set("select")
            self.on_tool_change()
            self.set_status(msg + " · 선택 도구로 전환됨 (이동/크기조절/수정 가능)")
        self.canvas.focus_set()

    def cur_size(self):
        try:
            return max(4.0, float(self.size_var.get()))
        except (tk.TclError, ValueError):
            return 14.0

    def cur_weight(self):
        n = self.weight_var.get()
        return WEIGHT_NAMES.index(n) if n in WEIGHT_NAMES else 0

    def create_text(self, p0, p1):
        """텍스트 도구로 클릭(또는 드래그)한 자리에서 바로 입력 시작"""
        rv = fitz.Rect(p0, p1).normalize()
        self.start_inline_edit(None, fitz.Point(rv.x0, rv.y0))

    # ---- 페이지 위 바로 입력 (인라인 편집기)
    def _editor_style(self, o=None):
        """입력창에 쓸 (Tk 글꼴, 글자색, 배경색, 정렬). o가 있으면 그 개체, 없으면 속성 패널 값"""
        if o is not None:
            label = o.font_label if o.font_kind == "file" else (spec_for(o).label or o.font_value)
            size, wt, color, bg, align = o.size, o.weight, o.color, o.bg, o.align
        else:
            label = self.current_font_spec().label
            size, wt, color = self.cur_size(), self.cur_weight(), self.colors["text"]
            bg = None if self.bg_none.get() else self.colors["text_bg"]
            align = {"왼쪽": 0, "가운데": 1, "오른쪽": 2}[self.align_var.get()]
        f = tkfont.Font(self, family=tk_family(label), size=-max(4, int(round(size * self.zoom))),
                        weight="bold" if wt >= 2 else "normal")
        return f, color, bg or "#FFFFFF", ("left", "center", "right")[align]

    def start_inline_edit(self, o, pv):
        """o=None 이면 새 글(pv = 왼쪽 위, 화면 좌표 V), 아니면 기존 텍스트 개체를 그 자리에서 수정"""
        self.finish_inline_edit()
        if self.doc is None:
            return
        if o is not None:
            rv = text_layout(o, self.geo)
            pad = TEXT_PAD + o.size * weight_ratio(o) / 2
            pv = fitz.Point(rv.x0 + pad, rv.y0 + pad)
        else:
            pv = fitz.Point(pv.x + TEXT_PAD, pv.y + TEXT_PAD)
        f, fg, bg, just = self._editor_style(o)
        t = tk.Text(self.canvas, wrap="none", undo=True, bd=0, padx=0, pady=0, relief="flat",
                    font=f, fg=fg, bg=bg, insertbackground=fg, insertwidth=2,
                    highlightthickness=1, highlightcolor="#0078d7", highlightbackground="#0078d7",
                    width=1, height=1)
        t.tag_configure("al", justify=just)
        if o is not None:
            t.insert("1.0", o.text)
            t.edit_reset()
            self.canvas.itemconfig("o%d" % o.id, state="hidden")
            self.canvas.delete("sel")
        cx, cy = self.v2c(pv.x, pv.y)
        win = self.canvas.create_window(cx - 1, cy - 1, anchor="nw", window=t, tags="editor")
        self._ed = {"w": t, "win": win, "font": f, "obj": o, "page": self.cur, "pv": pv}
        t.bind("<Control-Return>", lambda e: (self.finish_inline_edit(), "break")[1])
        t.bind("<Escape>", lambda e: (self.finish_inline_edit(cancel=True), "break")[1])
        t.bind("<KeyRelease>", lambda e: self._inline_fit())
        t.bind("<<Paste>>", lambda e: self.after_idle(self._inline_fit))
        t.bind("<FocusOut>", lambda e, w=t: self.after(80, lambda: self._inline_focus_out(w)))
        self._inline_fit()
        t.mark_set("insert", "end-1c")
        t.focus_set()
        self.set_status("입력 중 - Enter=줄바꿈 · Ctrl+Enter 또는 바깥 클릭=확정 · Esc=취소")

    def _inline_fit(self):
        """입력한 글에 맞춰 입력창 크기를 늘리고 줄인다"""
        ed = self._ed
        if ed is None:
            return
        t, f = ed["w"], ed["font"]
        t.tag_add("al", "1.0", "end")
        lines = t.get("1.0", "end-1c").split("\n")
        w = max(f.measure(l) for l in lines) + f.measure("M") + 4
        h = len(lines) * f.metrics("linespace") + 4
        self.canvas.itemconfig(ed["win"], width=max(w, 20), height=h)

    def _inline_restyle(self):
        """입력 중에 속성 패널(글꼴·크기·굵기·색·정렬)을 바꾸면 입력창에도 반영"""
        ed = self._ed
        if ed is None:
            return
        f, fg, bg, just = self._editor_style(None)
        ed["font"] = f
        ed["w"].config(font=f, fg=fg, bg=bg, insertbackground=fg)
        ed["w"].tag_configure("al", justify=just)
        self._inline_fit()

    def _inline_focus_out(self, w):
        ed = self._ed
        if ed is None or ed["w"] is not w:
            return
        try:
            fw = self.focus_get()
        except (KeyError, tk.TclError):
            fw = None
        if fw is None or fw is w:          # 다른 프로그램으로 잠시 전환 -> 계속 입력
            return
        if isinstance(fw, (ttk.Combobox, ttk.Spinbox, tk.Spinbox, ttk.Entry)) and fw is not self.page_entry:
            return                           # 입력 중 속성 패널 조작 -> 계속 입력
        self.finish_inline_edit()

    def finish_inline_edit(self, cancel=False):
        """입력 확정(또는 취소). 새 글이면 레이어 개체를 만들고, 기존 글이면 내용을 바꾼다"""
        ed = self._ed
        if ed is None:
            return
        self._ed = None
        text = ed["w"].get("1.0", "end-1c")
        try:
            self.canvas.delete(ed["win"])
            ed["w"].destroy()
        except tk.TclError:
            pass
        o = ed["obj"]
        if self.doc is None or ed["page"] != self.cur or self.page_img is None:
            return
        if cancel:
            self.draw_overlays()
            self.set_status("입력 취소됨")
            return
        if o is None:
            if text.strip():
                self.make_text_obj(text, ed["pv"])
            else:
                self.draw_overlays()
                self.set_status("빈 글은 추가하지 않았습니다")
            return
        if text == o.text:
            self.draw_overlays()
            return
        if not text.strip():
            pre = self.snapshot(False)
            self.layers[self.cur] = [x for x in self.layers[self.cur] if x.id != o.id]
            self.sel_id = None
            self.commit(pre, "내용을 모두 지워 텍스트 개체를 삭제했습니다 (Ctrl+Z로 복구)")
            return
        spec = self.resolve_font(spec_for(o), text)
        if spec is None:
            self.draw_overlays()
            return
        pre = self.snapshot(False)
        self.set_obj_font(o, spec)
        o.text = text
        self.commit(pre, "텍스트 수정됨")

    def make_text_obj(self, text, pv):
        spec = self.resolve_font(self.current_font_spec(), text)
        if spec is None:
            self.draw_overlays()
            return
        rv = fitz.Rect(pv.x - TEXT_PAD, pv.y - TEXT_PAD, pv.x + 20, pv.y + 20)
        o = Obj("text", rect=tuple(rect_v2p(rv, self.geo)), text=text, size=self.cur_size(),
                align={"왼쪽": 0, "가운데": 1, "오른쪽": 2}[self.align_var.get()], weight=self.cur_weight(),
                color=self.colors["text"], bg=None if self.bg_none.get() else self.colors["text_bg"],
                auto_w=True)
        self.set_obj_font(o, spec)
        pad = o.size * weight_ratio(o) / 2                # 굵은 글은 글자 자리가 입력창과 같도록
        rv = fitz.Rect(rv.x0 - pad, rv.y0 - pad, rv.x1, rv.y1)
        o.rect = tuple(rect_v2p(rv, self.geo))
        lay = text_layout(o, self.geo)                 # 페이지 오른쪽 밖으로 나가면 당겨 넣기
        if lay.x1 > self.geo.vw - 2:
            nx0 = max(2.0, rv.x0 - (lay.x1 - (self.geo.vw - 2)))
            rv = fitz.Rect(nx0, rv.y0, nx0 + rv.width, rv.y1)
            o.rect = tuple(rect_v2p(rv, self.geo))
        self.add_obj(o, "텍스트 추가됨")

    def create_highlight(self, p0, p1):
        page = self.doc[self.cur]
        urect = rect_v2p(fitz.Rect(p0, p1).normalize(), self.geo)
        words = [w for w in page.get_text("words")
                 if urect.contains(fitz.Point((w[0] + w[2]) / 2, (w[1] + w[3]) / 2))]
        if words:                                        # 글자 위 -> 줄 단위로 맞춰 붙기
            groups = {}
            for w in words:
                key = (w[5], w[6])
                r = fitz.Rect(w[:4])
                groups[key] = (groups[key] | r) if key in groups else r
            rects = [tuple(r) for r in groups.values()]
            msg = "글자 마커 추가됨"
        else:
            rects = [tuple(urect)]
            msg = "영역 마커 추가됨"
        o = Obj("highlight", rects=rects, color=self.colors["hl"],
                opacity=min(1.0, max(0.05, self.hl_opacity_var.get() / 100.0)))
        self.add_obj(o, msg)

    def create_shape(self, kind, p0, p1):
        kw = dict(stroke=self.colors["stroke"], fill=None if self.fill_none.get() else self.colors["fill"],
                  width=float(self.width_var.get()), opacity=min(1.0, max(0.05, self.opacity_var.get() / 100.0)))
        if kind in LINE_KINDS:
            a, b = fitz.Point(p0) * self.geo.derotm, fitz.Point(p1) * self.geo.derotm
            o = Obj(kind, p0=(a.x, a.y), p1=(b.x, b.y), **kw)
        else:
            o = Obj(kind, rect=tuple(rect_v2p(fitz.Rect(p0, p1).normalize(), self.geo)), **kw)
        self.add_obj(o, "%s 추가됨" % KIND_LABEL[kind])

    # ---------------------------------------------------------- 선택 / 속성
    def select_obj(self, o):
        self.sel_id = o.id if o else None
        if o:
            self.load_props(o)
        self.draw_selection()
        self.sync_layer_selection()

    def load_props(self, o):
        """선택한 개체의 값을 오른쪽 속성 패널에 표시"""
        self._loading = True
        try:
            if o.kind == "text":
                if o.font_label in self.fonts:
                    self.font_var.set(o.font_label)
                self.size_var.set(int(round(o.size)))
                self.align_var.set({0: "왼쪽", 1: "가운데", 2: "오른쪽"}.get(o.align, "왼쪽"))
                self.weight_var.set(WEIGHT_NAMES[o.weight] if 0 <= o.weight < len(WEIGHTS) else WEIGHT_NAMES[0])
                self._set_color("text", o.color)
                self.bg_none.set(o.bg is None)
                if o.bg:
                    self._set_color("text_bg", o.bg)
            elif o.kind == "highlight":
                self._set_color("hl", o.color)
                self.hl_opacity_var.set(int(round(o.opacity * 100)))
            elif o.kind == "image":
                pass
            else:
                self._set_color("stroke", o.stroke)
                self.fill_none.set(o.fill is None)
                if o.fill:
                    self._set_color("fill", o.fill)
                self.width_var.set(o.width)
                self.opacity_var.set(int(round(o.opacity * 100)))
        finally:
            self._loading = False

    def _set_color(self, key, hexv):
        self.colors[key] = hexv
        self.color_btns[key].config(bg=hexv, activebackground=hexv)

    def on_prop_var(self, *_a):
        if self._loading:
            return
        if self._ed is not None:
            self.after_idle(self._inline_restyle)
        if self.sel_obj() is None:
            return
        if self._prop_job:
            self.after_cancel(self._prop_job)
        self._prop_job = self.after(350, self.apply_props)

    def apply_props(self):
        self._prop_job = None
        o = self.sel_obj()
        if o is None:
            return
        pre = self.snapshot(False)
        if self.props_to_obj(o):
            self.commit(pre, "속성 변경")

    def props_to_obj(self, o):
        before = o.sig()
        try:
            if o.kind == "text":
                spec = self.fonts.get(self.font_var.get())
                size = max(4.0, float(self.size_var.get()))
                align = {"왼쪽": 0, "가운데": 1, "오른쪽": 2}[self.align_var.get()]
                if spec is not None:
                    spec = self.resolve_font(spec, o.text)
                    if spec is not None:
                        self.set_obj_font(o, spec)
                o.size, o.align, o.weight = size, align, self.cur_weight()
                o.color = self.colors["text"]
                o.bg = None if self.bg_none.get() else self.colors["text_bg"]
            elif o.kind == "highlight":
                op = min(1.0, max(0.05, self.hl_opacity_var.get() / 100.0))
                o.color, o.opacity = self.colors["hl"], op
            elif o.kind == "image":
                return False
            else:
                w = max(0.1, float(self.width_var.get()))
                op = min(1.0, max(0.05, self.opacity_var.get() / 100.0))
                o.stroke = self.colors["stroke"]
                o.fill = None if self.fill_none.get() else self.colors["fill"]
                o.width, o.opacity = w, op
        except (tk.TclError, ValueError):
            return False
        return o.sig() != before

    # ---------------------------------------------------------- 레이어 목록 / 조작
    def layer_label(self, o):
        return "%s %s %s" % ("●" if o.visible else "○", KIND_ICON[o.kind], o.label())

    def refresh_layer_list(self):
        self._syncing_layers = True
        try:
            self.layer_tree.delete(*self.layer_tree.get_children())
            if self.doc is not None and self.cur < len(self.layers):
                for o in reversed(self.layers[self.cur]):
                    self.layer_tree.insert("", "end", iid=str(o.id), text=self.layer_label(o))
            if self.sel_id is not None and self.layer_tree.exists(str(self.sel_id)):
                self.layer_tree.selection_set(str(self.sel_id))
        finally:
            self._syncing_layers = False

    def sync_layer_selection(self):
        self._syncing_layers = True
        try:
            if self.sel_id is not None and self.layer_tree.exists(str(self.sel_id)):
                self.layer_tree.selection_set(str(self.sel_id))
                self.layer_tree.see(str(self.sel_id))
            else:
                self.layer_tree.selection_remove(self.layer_tree.selection())
        finally:
            self._syncing_layers = False

    def on_layer_select(self, _e=None):
        if self._syncing_layers or self.doc is None:
            return
        sel = self.layer_tree.selection()
        if not sel:
            return
        oid = int(sel[0])
        if oid == self.sel_id:          # 이미 선택된 개체면 (가상 이벤트 지연 도착) 무시
            return
        for o in self.layers[self.cur]:
            if o.id == oid:
                self.tool.set("select")
                self.canvas.config(cursor="arrow")
                self.sel_id = oid
                self.load_props(o)
                self.draw_selection()
                return

    def need_sel(self):
        self.finish_inline_edit()
        o = self.sel_obj()
        if o is None:
            self.set_status("먼저 개체를 선택하세요 (캔버스에서 클릭하거나 레이어 목록에서 선택)")
        return o

    def delete_obj(self):
        o = self.need_sel()
        if o is None:
            return
        pre = self.snapshot(False)
        self.layers[self.cur] = [x for x in self.layers[self.cur] if x.id != o.id]
        self.sel_id = None
        self.commit(pre, "개체 삭제됨 (Ctrl+Z로 복구)")

    def reorder(self, how):
        o = self.need_sel()
        if o is None:
            return
        lst = self.layers[self.cur]
        i = lst.index(o)
        j = {"up": i + 1, "down": i - 1, "front": len(lst) - 1, "back": 0}[how]
        j = max(0, min(len(lst) - 1, j))
        if j == i:
            return
        pre = self.snapshot(False)
        lst.pop(i)
        lst.insert(j, o)
        self.commit(pre, "레이어 순서 변경")

    def toggle_visible(self):
        o = self.need_sel()
        if o is None:
            return
        pre = self.snapshot(False)
        o.visible = not o.visible
        self.commit(pre, "표시" if o.visible else "숨김 (저장 시 PDF에도 포함되지 않음)")

    def rename_obj(self):
        o = self.need_sel()
        if o is None:
            return
        n = simpledialog.askstring(APP_TITLE, "레이어 이름 (비우면 기본 이름):", initialvalue=o.name, parent=self)
        if n is None:
            return
        pre = self.snapshot(False)
        o.name = n.strip()
        self.commit(pre, None)

    def edit_text_selected(self):
        o = self.need_sel()
        if o is None:
            return
        if o.kind != "text":
            self.set_status("텍스트 개체만 내용을 수정할 수 있습니다.")
            return
        self.edit_text(o)

    def edit_text(self, o):
        """기존 텍스트를 팝업 없이 그 자리에서 수정"""
        if not o.visible:
            self.set_status("숨긴 개체입니다. 먼저 '표시 / 숨김'으로 보이게 하세요.")
            return
        self.start_inline_edit(o, None)

    def duplicate_obj(self):
        o = self.need_sel()
        if o is None:
            return
        pre = self.snapshot(False)
        n = o.clone()
        n.shift(12, 12)
        lst = self.layers[self.cur]
        lst.insert(lst.index(o) + 1, n)
        self.sel_id = n.id
        self.commit(pre, "개체 복제됨")

    def copy_obj(self):
        o = self.need_sel()
        if o is None:
            return
        self._clip = (o.to_dict(), self.cur)
        self._clip_seq = clipboard_seq()
        self._paste_n = 0
        self.set_status("복사됨 - 원하는 페이지에서 붙여넣기(Ctrl+V)")

    def paste_obj(self):
        """Ctrl+V: 개체를 복사한 뒤 새로 캡처(클립보드 변경)했으면 캡처 이미지, 아니면 복사한 개체"""
        if self.doc is None:
            return
        self.finish_inline_edit()
        if not self._clip or clipboard_seq() != getattr(self, "_clip_seq", None):
            data = clipboard_image()
            if data:
                self.paste_image(data)
                return
        if not self._clip:
            self.set_status("붙여 넣을 것이 없습니다 - 화면을 캡처(Win+Shift+S)하거나 개체를 복사(Ctrl+C)하세요")
            return
        d, page = self._clip
        n = Obj.from_dict(d)
        Obj._counter += 1
        n.id = Obj._counter
        self._paste_n += 1
        if page == self.cur:
            n.shift(12 * self._paste_n, 12 * self._paste_n)
        pre = self.snapshot(False)
        self.layers[self.cur].append(n)
        self.sel_id = n.id
        self.commit(pre, "붙여넣기 완료")

    def paste_image(self, data):
        """캡처 이미지를 지금 보이는 화면 가운데에 이미지 레이어로 붙인다.
        크기는 캡처한 화면 크기 그대로(현재 확대 비율 기준), 페이지보다 크면 페이지 안으로 줄인다."""
        try:
            png, pw, ph = to_png(data)
        except Exception as e:
            messagebox.showerror(APP_TITLE, "클립보드 이미지를 읽을 수 없습니다:\n%s" % e)
            return
        key = store_image(png)
        w, h = pw / self.zoom, ph / self.zoom
        s = min(1.0, self.geo.vw * 0.9 / w, self.geo.vh * 0.9 / h)
        w, h = w * s, h * s
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        c = self.c2v(self.canvas.canvasx(cw / 2), self.canvas.canvasy(ch / 2))
        x0 = min(max(c.x - w / 2, 0), self.geo.vw - w)
        y0 = min(max(c.y - h / 2, 0), self.geo.vh - h)
        o = Obj("image", rect=tuple(rect_v2p(fitz.Rect(x0, y0, x0 + w, y0 + h), self.geo)), img=key,
                img_rot=self.geo.rot)
        self.add_obj(o, "이미지 붙여넣기 완료 (%dx%d px)" % (pw, ph))

    def copy_to_all_pages(self):
        o = self.need_sel()
        if o is None:
            return
        if len(self.doc) < 2:
            return
        if not messagebox.askyesno(APP_TITLE, "선택한 개체를 다른 모든 페이지(%d쪽)에 같은 위치로 복사할까요?\n"
                                   "(Ctrl+Z로 취소 가능)" % (len(self.doc) - 1)):
            return
        pre = self.snapshot(False)
        for i in range(len(self.doc)):
            if i != self.cur:
                self.layers[i].append(o.clone())
        self.commit(pre, "모든 페이지에 복사됨")
        for i in range(len(self.doc)):
            if i != self.cur:
                self.update_thumb(i)

    # ---------------------------------------------------------- 페이지 편집
    def need_doc(self):
        self.finish_inline_edit()
        if self.doc is None:
            messagebox.showinfo(APP_TITLE, "먼저 PDF를 열어주세요.")
            return False
        return True

    def delete_selected(self):
        if not self.need_doc():
            return
        sel = self.selected_pages()
        if not sel:
            return
        if len(sel) >= len(self.doc):
            messagebox.showwarning(APP_TITLE, "모든 페이지를 삭제할 수는 없습니다.")
            return
        if not messagebox.askyesno(APP_TITLE, "선택한 %d페이지를 삭제할까요?\n(Ctrl+Z로 취소 가능)" % len(sel)):
            return
        self.push_undo(structural=True)
        for i in reversed(sel):
            self.doc.delete_page(i)
            del self.layers[i]
        self.refresh_thumbs(select=[min(sel[0], len(self.doc) - 1)])
        self.set_status("%d페이지 삭제됨" % len(sel))

    def rotate_selected(self):
        if not self.need_doc():
            return
        sel = self.selected_pages() or [self.cur]
        self.push_undo(structural=True)
        for i in sel:
            p = self.doc[i]
            p.set_rotation((p.rotation + 90) % 360)
        self.refresh_thumbs(select=sel)

    def apply_order(self, order, new_sel):
        self.push_undo(structural=True)
        self.doc.select(order)
        self.layers = [self.layers[o] for o in order]
        self.refresh_thumbs(select=new_sel)

    def move_selected(self, d):
        if not self.need_doc():
            return
        sel = self.selected_pages()
        if not sel:
            return
        n = len(self.doc)
        s = set(sel)
        order = list(range(n))
        if d < 0:
            for i in range(1, n):
                if order[i] in s and order[i - 1] not in s:
                    order[i - 1], order[i] = order[i], order[i - 1]
        else:
            for i in range(n - 2, -1, -1):
                if order[i] in s and order[i + 1] not in s:
                    order[i], order[i + 1] = order[i + 1], order[i]
        if order == list(range(n)):
            return
        self.apply_order(order, [k for k, o in enumerate(order) if o in s])

    def move_to(self):
        if not self.need_doc():
            return
        sel = self.selected_pages()
        if not sel:
            return
        n = len(self.doc)
        t = simpledialog.askinteger(APP_TITLE, "선택한 페이지를 몇 번째 위치로 옮길까요? (1~%d)" % n,
                                    minvalue=1, maxvalue=n, parent=self)
        if not t:
            return
        s = set(sel)
        rest = [i for i in range(n) if i not in s]
        pos = min(t - 1, len(rest))
        order = rest[:pos] + sel + rest[pos:]
        self.apply_order(order, list(range(pos, pos + len(sel))))

    def insert_blank(self):
        if not self.need_doc():
            return
        sel = self.selected_pages()
        at = (max(sel) if sel else self.cur) + 1
        r = self.doc[self.cur].rect
        self.push_undo(structural=True)
        self.doc.new_page(pno=at, width=r.width, height=r.height)
        self.layers.insert(at, [])
        self.refresh_thumbs(select=[at])

    def insert_images(self):
        if not self.need_doc():
            return
        files = filedialog.askopenfilenames(parent=self, title="이미지 선택",
                                            filetypes=[("이미지", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff")])
        if not files:
            return
        sel = self.selected_pages()
        at = (max(sel) if sel else self.cur) + 1
        self.push_undo(structural=True)
        try:
            for k, f in enumerate(files):
                pix = fitz.Pixmap(f)
                w, h = (842, 595) if pix.width > pix.height else (595, 842)
                pg = self.doc.new_page(pno=at + k, width=w, height=h)
                pg.insert_image(pg.rect, filename=f, keep_proportion=True)
                self.layers.insert(at + k, [])
        except Exception as e:
            self.undo(quiet=True)
            messagebox.showerror(APP_TITLE, "이미지 삽입 실패:\n%s" % e)
            return
        self.refresh_thumbs(select=list(range(at, at + len(files))))

    def insert_pdfs(self, where):
        if not self.need_doc():
            return
        files = filedialog.askopenfilenames(parent=self, title="삽입할 PDF 선택",
                                            filetypes=[("PDF", "*.pdf")])
        if not files:
            return
        sel = self.selected_pages()
        at = len(self.doc) if where == "end" else (max(sel) if sel else self.cur) + 1
        first = at
        self.push_undo(structural=True)
        try:
            for f in files:
                src = self.open_pdf_file(f)
                if src is None:
                    continue
                if len(files) == 1 and len(src) > 1:
                    txt = simpledialog.askstring(
                        APP_TITLE, "가져올 페이지 범위 (예: 1-3,5)\n비워두면 전체 %d쪽" % len(src),
                        parent=self)
                    if txt and txt.strip():
                        idx = [i for _, g in parse_ranges(txt, len(src)) for i in g]
                        src.select(idx)
                self.doc.insert_pdf(src, start_at=at)
                self.layers[at:at] = [[] for _ in range(len(src))]
                at += len(src)
                src.close()
        except Exception as e:
            self.undo(quiet=True)
            messagebox.showerror(APP_TITLE, "삽입 실패:\n%s" % e)
            return
        if at == first:
            self.undo_stack.pop()
            return
        self.refresh_thumbs(select=list(range(first, at)))
        self.set_status("%d페이지 삽입됨" % (at - first))

    # ---------------------------------------------------------- 분할 / 추출
    def save_pages(self, indices, path):
        new = fitz.open()
        for a, b in group_contiguous(indices):
            new.insert_pdf(self.doc, from_page=a, to_page=b)
        for k, idx in enumerate(indices):
            apply_objects(new[k], self.layers[idx])
        save_doc(new, path, subset=self.has_file_fonts([self.layers[i] for i in indices]), inplace=True)
        new.close()

    def split_dialog(self):
        if not self.need_doc():
            return
        n = len(self.doc)
        sel = self.selected_pages()
        win = tk.Toplevel(self)
        win.title("PDF 분할 / 추출")
        win.transient(self)
        mode = tk.StringVar(value="sel")
        rng = tk.StringVar(value="1-%d" % max(1, n // 2))
        every = tk.IntVar(value=1)
        fr = ttk.Frame(win, padding=12)
        fr.pack()
        ttk.Radiobutton(fr, text="선택한 %d페이지를 새 PDF 1개로 추출" % len(sel), value="sel",
                        variable=mode).grid(row=0, column=0, columnspan=2, sticky="w", pady=3)
        ttk.Radiobutton(fr, text="페이지 범위별로 분할 (범위마다 파일 1개)", value="ranges",
                        variable=mode).grid(row=1, column=0, columnspan=2, sticky="w", pady=3)
        ttk.Entry(fr, textvariable=rng, width=28).grid(row=2, column=1, sticky="w")
        ttk.Label(fr, text="예: 1-3, 4-6, 7").grid(row=2, column=0, sticky="e", padx=4)
        ttk.Radiobutton(fr, text="N쪽씩 분할", value="every", variable=mode).grid(
            row=3, column=0, sticky="w", pady=3)
        ttk.Spinbox(fr, from_=1, to=max(1, n), textvariable=every, width=6).grid(row=3, column=1, sticky="w")
        ttk.Radiobutton(fr, text="1쪽씩 전부 분할", value="each", variable=mode).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=3)
        ttk.Label(fr, text="※ 레이어 개체는 분할된 PDF에 합쳐져 저장됩니다.", foreground="#555").grid(
            row=5, column=0, columnspan=2, sticky="w")

        def run():
            try:
                base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "문서"
                m = mode.get()
                if m == "sel":
                    if not sel:
                        messagebox.showinfo(APP_TITLE, "왼쪽 목록에서 페이지를 선택하세요.", parent=win)
                        return
                    out = filedialog.asksaveasfilename(parent=win, defaultextension=".pdf",
                                                       initialfile="%s_추출.pdf" % base,
                                                       filetypes=[("PDF", "*.pdf")])
                    if not out:
                        return
                    self.save_pages(sel, out)
                    messagebox.showinfo(APP_TITLE, "저장 완료:\n%s" % out, parent=win)
                    win.destroy()
                    return
                if m == "ranges":
                    parts = parse_ranges(rng.get(), n)
                elif m == "every":
                    k = max(1, int(every.get()))
                    parts = [("%d-%d" % (s + 1, min(s + k, n)), list(range(s, min(s + k, n))))
                             for s in range(0, n, k)]
                else:
                    parts = [(str(i + 1), [i]) for i in range(n)]
                folder = filedialog.askdirectory(parent=win, title="저장할 폴더 선택")
                if not folder:
                    return
                for label, idx in parts:
                    self.save_pages(idx, unique_path(os.path.join(folder, "%s_p%s.pdf" % (base, label))))
                messagebox.showinfo(APP_TITLE, "%d개 파일 저장 완료:\n%s" % (len(parts), folder), parent=win)
                win.destroy()
            except ValueError as e:
                messagebox.showerror(APP_TITLE, str(e), parent=win)
            except Exception as e:
                messagebox.showerror(APP_TITLE, "분할 실패:\n%s" % e, parent=win)

        bar = ttk.Frame(fr)
        bar.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        ttk.Button(bar, text="실행", command=run).pack(side="left", padx=4)
        ttk.Button(bar, text="닫기", command=win.destroy).pack(side="left", padx=4)
        win.grab_set()


def selftest(logpath):
    """exe 안에 필요한 부품(PDF 엔진, 글꼴 축소, 투명 PNG, 작업파일 입출력)이 다 들어 있는지 점검.
    사용: JP-PDF EDITOR.exe --selftest 결과.txt"""
    import tempfile
    lines, state = [], {"ok": True}

    def chk(name, cond, extra=""):
        state["ok"] = state["ok"] and bool(cond)
        lines.append("%s %s %s" % ("OK  " if cond else "FAIL", name, extra))

    try:
        chk("실행 환경", True, "python %s / %s / frozen=%s" % (
            sys.version.split()[0], fitz.__doc__.split(":")[0], bool(getattr(sys, "frozen", False))))
        spec = default_spec()
        chk("글꼴 탐색", len(FONTS) > 0, "%d개, 기본=%s" % (len(FONTS), spec.label))
        tmp = tempfile.mkdtemp()
        doc = fitz.open()
        pg = doc.new_page(width=400, height=300)
        pg.insert_text((40, 60), "Selftest page", fontsize=14)
        objs = [
            Obj("text", rect=(30, 100, 230, 140), text="한글 테스트 텍스트", size=14, color="#0000CC",
                bg="#FFFF99", font_kind=spec.kind, font_value=spec.value, font_label=spec.label),
            Obj("highlight", rects=[(38, 44, 140, 66)], color="#FFFF00", opacity=0.4),
            Obj("rect", rect=(250, 30, 330, 80), stroke="#FF0000", width=2, opacity=1.0),
            Obj("oval", rect=(20, 160, 120, 210), stroke="#008000", fill="#CCFFCC", width=2, opacity=0.7),
            Obj("arrow", p0=(180, 250), p1=(330, 200), stroke="#800080", width=3, opacity=1.0),
            Obj("text", rect=(250, 100, 260, 110), text="굵게 Bold", size=16, weight=2, auto_w=True,
                font_kind=spec.kind, font_value=spec.value, font_label=spec.label),
            Obj("rrect", rect=(140, 160, 220, 200), stroke="#0000FF", fill="#DDEEFF", width=2),
            Obj("triangle", rect=(240, 220, 300, 280), stroke="#FF8800", width=2),
            Obj("darrow", p0=(20, 280), p1=(160, 280), stroke="#000000", width=2)]
        apply_objects(doc[0], objs)
        out = os.path.join(tmp, "selftest.pdf")
        save_doc(doc, out, subset=(spec.kind == "file"), inplace=True)
        size = os.path.getsize(out)
        txt = "".join(fitz.open(out)[0].get_text().split())
        ko_ok = spec.supports("한글")
        chk("레이어 합쳐 PDF 저장", "Selftestpage" in txt and (not ko_ok or "한글테스트텍스트" in txt),
            "%d KB" % (size // 1024))
        if spec.kind == "file":
            chk("글꼴 축소(fontTools)", size < os.path.getsize(spec.value) * 0.5,
                "글꼴 %d KB -> PDF %d KB" % (os.path.getsize(spec.value) // 1024, size // 1024))
        geo = PageGeo(fitz.open(out)[0])
        root = tk.Tk()
        root.withdraw()
        res = render_obj_png(objs[3], geo, 1.5)
        img = tk.PhotoImage(data=res[0])
        chk("투명 PNG 개체 렌더(Tk)", img.width() > 10 and img.height() > 10, "%dx%d" % (img.width(), img.height()))
        root.destroy()
        lay = json.dumps([o.to_dict() for o in objs], ensure_ascii=False)
        back = [Obj.from_dict(d) for d in json.loads(lay)]
        zp = os.path.join(tmp, "t.pdfedit")
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("layers.json", lay)
        with zipfile.ZipFile(zp) as z:
            back2 = [Obj.from_dict(d) for d in json.loads(z.read("layers.json").decode("utf-8"))]
        chk("작업파일(.pdfedit) 입출력", [o.sig() for o in objs] == [o.sig() for o in back2] == [o.sig() for o in back])
        # v1.1: 텍스트 상자가 글자 크기에 맞춰지는지 (크기 2배 -> 상자 높이·폭 약 2배)
        t = Obj("text", rect=(10, 10, 20, 20), text="크기 연동", size=12, auto_w=True,
                font_kind=spec.kind, font_value=spec.value, font_label=spec.label)
        a = text_layout(t, geo)
        t.size = 24
        b = text_layout(t, geo)
        chk("텍스트 상자 = 글자 크기 연동", 1.7 < b.height / a.height < 2.3 and 1.7 < b.width / a.width < 2.3,
            "12pt %.0fx%.0f -> 24pt %.0fx%.0f" % (a.width, a.height, b.width, b.height))
        t.weight = 3
        c = text_layout(t, geo)
        chk("굵기 지정 시 상자 여유", c.width > b.width, "%.1f -> %.1f" % (b.width, c.width))
        old = Obj.from_dict({"kind": "text", "rect": [0, 0, 50, 20], "text": "x"})
        chk("구버전 작업파일 호환 (굵기 없음)", old.weight == 0)
        # v1.2: 캡처 이미지(클립보드 비트맵) -> PNG -> 이미지 레이어
        import struct
        bw_, bh_ = 30, 20
        px = b"".join(bytes([255, 120, 0, 0]) for _ in range(bw_ * bh_))        # 32비트, 알파 0
        dib = struct.pack("<IiiHHIIiiII", 40, bw_, bh_, 1, 32, 0, len(px), 2835, 2835, 0, 0) + px
        png, pw_, ph_ = to_png(dib_to_bmp(dib))
        chk("클립보드 비트맵 -> PNG", (pw_, ph_) == (bw_, bh_) and not fitz.Pixmap(png).alpha,
            "%dx%d" % (pw_, ph_))
        key = store_image(png)
        d3 = fitz.open()
        p3 = d3.new_page(width=200, height=100)
        p3.set_rotation(90)
        apply_objects(p3, [Obj("image", rect=(20, 20, 80, 60), img=key)])
        chk("이미지 레이어 PDF 합치기", len(p3.get_images()) == 1)
        chk("클립보드 읽기 API", clipboard_seq() is not None)
    except Exception as e:
        import traceback
        chk("예외 없이 완료", False, repr(e) + "\n" + traceback.format_exc())
    lines.append("SELFTEST " + ("PASS" if state["ok"] else "FAIL"))
    with open(logpath, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return 0 if state["ok"] else 1


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--selftest":
        sys.exit(selftest(sys.argv[2]))
    try:  # Windows 고해상도 화면에서 선명하게
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    PDFEditor(arg).mainloop()


if __name__ == "__main__":
    main()
