"""
 *by Enes Acar*

Bu dosyada 6 fonksiyon var:

  1) extract_text_from_image   : Görselden/fotoğraftan metin çıkarır
  2) extract_text_from_pdf     : PDF'ten metin çıkarır (taranmış sayfalarda OCR yapar)
  3) create_searchable_pdf     : Taranmış PDF'i aranabilir/seçilebilir PDF yapar
  4) excel_to_csv              : Excel (.xlsx) dosyasını CSV'ye çevirir
  5) csv_to_excel              : CSV dosyasını Excel'e çevirir
  6) create_qr_code            : Metin/URL'den QR kod görseli üretir

KURULUM :
  - Sisteme kur: Tesseract (Turkish dil paketiyle) ve Ghostscript
  - Sonra:  pip install -r requirements-ocr.txt

PENCERE (ARAYÜZ) İLE KULLANIM:
  python ocr_tools.py          (argüman vermeden çalıştırınca pencere açılır)

KOMUT SATIRINDAN KULLANIM:
  python ocr_tools.py gorsel-metin foto.jpg
  python ocr_tools.py pdf-metin belge.pdf
  python ocr_tools.py aranabilir-pdf taranmis.pdf cikti.pdf
  python ocr_tools.py excel-csv tablo.xlsx tablo.csv [sayfa_adi]
  python ocr_tools.py csv-excel tablo.csv tablo.xlsx
  python ocr_tools.py qr "https://ornek.com" qr.png

PYTHON KODU İÇİNDEN KULLANIM:
  from ocr_tools import extract_text_from_image
  print(extract_text_from_image("foto.jpg"))
"""

import argparse
import csv
import io
import os
import re
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl                      # Excel okuma/yazma
import pypdfium2 as pdfium           # PDF'i okur ve sayfayı görsele çevirir (ek program gerektirmez)
import pytesseract                   # Tesseract OCR programına Python köprüsü
import qrcode                        # QR kod üretir
from PIL import Image, ImageOps      # Görsel açma ve yön düzeltme

# Pencere arayüzü için tkinter (Python ile birlikte gelir). Yüklü değilse sadece pencere kapalı kalır.
try:
    import queue
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:
    tk = None

# ----------------------------------------------------------------------
# WINDOWS İÇİN OTOMATİK YOL AYARI
# Tesseract ve Ghostscript PATH'e eklenmemişse varsayılan klasörlerde arar.
# ----------------------------------------------------------------------
# if os.name == "nt":
#     _tess = r"C:\Program Files\Tesseract-OCR"
#     if os.path.isdir(_tess):
#         pytesseract.pytesseract.tesseract_cmd = os.path.join(_tess, "tesseract.exe")
#         os.environ["PATH"] += os.pathsep + _tess
#     _gs_kok = Path(r"C:\Program Files\gs")
#     if _gs_kok.is_dir():
#         for _bin in _gs_kok.glob("gs*/bin"):
#             os.environ["PATH"] += os.pathsep + str(_bin)


# ----------------------------------------------------------------------
# YARDIMCI FONKSİYONLAR (dosyanın içinde kullanılır)
# ----------------------------------------------------------------------
def _dosya_kontrol(yol):
    """Girdi dosyası gerçekten var mı diye bakar; yoksa anlaşılır hata verir."""
    if not os.path.isfile(yol):
        raise FileNotFoundError(f"Dosya bulunamadı: {yol}")


def _klasor_hazirla(cikti_yolu):
    """Çıktı dosyasının klasörü yoksa oluşturur."""
    klasor = os.path.dirname(os.path.abspath(cikti_yolu))
    os.makedirs(klasor, exist_ok=True)


def _yonu_duzelt(img, min_guven=5.0):
    """
    Ters veya yan duran sayfayı Tesseract'ın yön algılamasıyla (OSD) düzeltir.
    Çok az yazı olan sayfalarda güven düşük olur; bu durumda görsele dokunmaz.
    """
    try:
        osd = pytesseract.image_to_osd(img, output_type=pytesseract.Output.DICT)
        if osd.get("rotate", 0) and float(osd.get("orientation_conf", 0)) >= min_guven:
            return img.rotate(-osd["rotate"], expand=True, fillcolor="white")
    except Exception:
        pass  # Yazı çok azsa veya osd dil dosyası yoksa algılama yapılamaz; olduğu gibi bırak
    return img


# ----------------------------------------------------------------------
# 1) GÖRSELDEN METİN
# ----------------------------------------------------------------------
def extract_text_from_image(input_path, language="tur+eng", yon_duzelt=True):
    """
    Görseldeki yazıyı okur ve metin olarak döndürür.

    input_path : görsel dosyası (jpg, png, tif, bmp, webp ...)
    language   : Tesseract dil kodu. "tur+eng" = Türkçe + İngilizce birlikte
    yon_duzelt : True ise ters/yan duran görseli otomatik döndürür

    """
    _dosya_kontrol(input_path)
    with Image.open(input_path) as img:
        # Telefon fotoğraflarında yön bilgisi (EXIF) ayrı tutulur; görsel yan gelmesin diye düzeltiriz.
        img = ImageOps.exif_transpose(img).convert("RGB")
        if yon_duzelt:
            img = _yonu_duzelt(img)  # Ters taranmış görselleri düzelt
        return pytesseract.image_to_string(img, lang=language).strip()


# ----------------------------------------------------------------------
# 2) PDF'TEN METİN
# ----------------------------------------------------------------------
def extract_text_from_pdf(input_path, language="tur+eng", page_based=False,
                          min_chars=20, dpi=300, yon_duzelt=True):
    """
    PDF'ten metin çıkarır. Her sayfa için şöyle çalışır:
      - Sayfada zaten seçilebilir metin varsa onu doğrudan alır (hızlı ve kusursuz).
      - Metin yoksa (taranmış sayfa) sayfayı görsele çevirip OCR yapar.

    input_path : PDF dosyası
    page_based : True ise her sayfanın metnini ayrı liste elemanı olarak döndürür
    min_chars  : bir sayfada bu kadar karakterden az metin varsa "taranmış" sayılır
    dpi        : OCR için sayfa çözünürlüğü (300 önerilir)
    yon_duzelt : True ise ters/yan duran taranmış sayfaları otomatik döndürür
    Döndürür   : page_based=False -> tek metin (str), True -> metin listesi (list[str])
    """
    _dosya_kontrol(input_path)
    pdf = pdfium.PdfDocument(input_path)
    sayfa_metinleri = []
    try:
        for sayfa in pdf:
            metin = sayfa.get_textpage().get_text_range().strip()
            if len(metin) < min_chars:
                # Sayfa taranmış görünüyor: görsele çevir ve OCR uygula.
                # PDF birimi 72 puan/inç olduğundan scale = dpi / 72
                gorsel = sayfa.render(scale=dpi / 72).to_pil().convert("RGB")
                if yon_duzelt:
                    gorsel = _yonu_duzelt(gorsel)
                metin = pytesseract.image_to_string(gorsel, lang=language).strip()
            sayfa_metinleri.append(metin)
    finally:
        pdf.close()  # Dosyayı her durumda kapat

    return sayfa_metinleri if page_based else "\n\n".join(sayfa_metinleri)


# ----------------------------------------------------------------------
# 3) ARANABİLİR PDF OLUŞTUR
# ----------------------------------------------------------------------
def create_searchable_pdf(input_path, output_path, language="tur+eng"):
    """
    Taranmış PDF'e görünmez bir metin katmanı ekler. Sayfa aynı görünür
    ama içindeki yazı seçilebilir, kopyalanabilir ve aranabilir olur.
    Giriş olarak görsel dosyası da verilebilir (jpg/png -> aranabilir PDF).

    input_path  : taranmış PDF (veya görsel)
    output_path : oluşacak PDF'in yolu

    """
    _dosya_kontrol(input_path)
    _klasor_hazirla(output_path)

    komut = [
        sys.executable, "-m", "ocrmypdf",  # OCRmyPDF'i ayrı süreç olarak çalıştır 
        "-l", language,                    # OCR dili
        "--skip-text",                     # Zaten metni olan sayfalara dokunma
        "--rotate-pages",                  # Ters duran sayfaları çevir
        "--rotate-pages-threshold", "8",   # Yön güveni eşiği (varsayılan 14 çok sıkı; ters sayfalar kaçıyordu)
        "--output-type", "pdf",            # PDF/A dönüşümü yapma (daha hızlı, daha küçük)
    ]
    # Görsel girdilerde çözünürlük bilgisi olmayabilir; 300 DPI varsay.
    if Path(input_path).suffix.lower() != ".pdf":
        komut += ["--image-dpi", "300"]
    komut += [input_path, output_path]

    sonuc = subprocess.run(komut, capture_output=True, text=True)
    if sonuc.returncode != 0:
        # Hatanın son kısmını göster (genelde asıl neden orada yazar)
        raise RuntimeError("OCRmyPDF hatası:\n" + (sonuc.stderr or sonuc.stdout)[-800:])
    return output_path


# ----------------------------------------------------------------------
# 4) EXCEL -> CSV
# ----------------------------------------------------------------------
def excel_to_csv(input_path, output_path, sheet_name=None):
    """
    Excel dosyasındaki bir sayfayı CSV'ye çevirir.

    input_path  : .xlsx dosyası
    output_path : oluşacak .csv dosyası
    sheet_name  : sayfa adı. None ise açık olan (ilk) sayfa alınır
   

    Notlar:
      - Formüller yerine son hesaplanan değerler yazılır.
      - Dosya UTF-8 (BOM'lu) yazılır; Excel'de Türkçe karakterler bozulmaz.
    """
    _dosya_kontrol(input_path)
    _klasor_hazirla(output_path)

    # data_only=True: formül yerine sonucunu oku. read_only=True: büyük dosyada hafıza tasarrufu.
    kitap = openpyxl.load_workbook(input_path, data_only=True, read_only=True)
    try:
        if sheet_name is None:
            sayfa = kitap.active
        elif sheet_name in kitap.sheetnames:
            sayfa = kitap[sheet_name]
        else:
            raise ValueError(f"'{sheet_name}' sayfası yok. Mevcut sayfalar: {kitap.sheetnames}")

        with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
            yazici = csv.writer(f)
            bos_satirlar = 0  # Sondaki boş satırları yazmamak için sayarız
            for satir in sayfa.iter_rows(values_only=True):
                degerler = [_hucre_metni(h) for h in satir]
                if all(d == "" for d in degerler):
                    bos_satirlar += 1
                    continue
                # Aradaki boş satırları koru, sondakileri at
                for _ in range(bos_satirlar):
                    yazici.writerow([])
                bos_satirlar = 0
                yazici.writerow(degerler)
    finally:
        kitap.close()
    return output_path


def _hucre_metni(deger):
    """Excel hücre değerini CSV'ye yazılacak metne çevirir."""
    if deger is None:
        return ""
    if isinstance(deger, (datetime, date)):
        return deger.isoformat(sep=" ") if isinstance(deger, datetime) else deger.isoformat()
    if isinstance(deger, float) and deger.is_integer():
        return str(int(deger))  # 5.0 yerine 5 yaz
    return str(deger)


# ----------------------------------------------------------------------
# 5) CSV -> EXCEL
# ----------------------------------------------------------------------
def csv_to_excel(input_path, output_path):
    """
    CSV dosyasını Excel (.xlsx) dosyasına çevirir.

    input_path  : .csv dosyası
    output_path : oluşacak .xlsx dosyası

    Notlar:
      - Ayırıcı (virgül, noktalı virgül, tab) otomatik bulunur.
        Türkçe Excel'in ürettiği ";" ayraçlı dosyalar da açılır.
      - Encoding önce UTF-8, olmazsa Windows-1254 (Türkçe) denenir.
      - Sayılar sayıya çevrilir; 0 ile başlayanlar (telefon, posta kodu) metin kalır.
    """
    _dosya_kontrol(input_path)
    _klasor_hazirla(output_path)

    # Önce UTF-8, olmazsa eski Türkçe Windows kodlaması
    try:
        with open(input_path, encoding="utf-8-sig", newline="") as f:
            icerik = f.read()
    except UnicodeDecodeError:
        with open(input_path, encoding="cp1254", newline="") as f:
            icerik = f.read()

    # Ayırıcıyı ilk bölüme bakarak tahmin et
    try:
        ayirici = csv.Sniffer().sniff(icerik[:4096], delimiters=",;\t|").delimiter
    except csv.Error:
        ayirici = ","

    kitap = openpyxl.Workbook()
    sayfa = kitap.active
    # StringIO + newline='' : tırnak içindeki satır sonları (çok satırlı hücreler) korunur
    for satir in csv.reader(io.StringIO(icerik, newline=""), delimiter=ayirici):
        sayfa.append([_sayiya_cevir(h) for h in satir])
    kitap.save(output_path)
    return output_path


def _sayiya_cevir(metin):
    """'42' -> 42, '3.14' -> 3.14. 0123 gibi baştaki sıfırlı veya çok uzun sayılar metin kalır."""
    m = metin.strip()
    if re.fullmatch(r"-?(0|[1-9]\d{0,14})", m):
        return int(m)
    if re.fullmatch(r"-?(0|[1-9]\d{0,14})\.\d{1,10}", m):
        return float(m)
    return metin


# ----------------------------------------------------------------------
# 6) QR KOD
# ----------------------------------------------------------------------
def create_qr_code(data, output_path):
    """
    Verilen metin veya URL için QR kod görseli üretir.

    data        : QR'a yazılacak metin / web adresi
    output_path : oluşacak görsel (örn. qr.png)
    """
    if not data:
        raise ValueError("QR koda yazılacak veri boş olamaz.")
    _klasor_hazirla(output_path)

    qr = qrcode.QRCode(
        error_correction=qrcode.constants.ERROR_CORRECT_M,  # %15 hasara dayanıklı
        box_size=10,   # Her karenin piksel boyutu
        border=4,      # Kenar boşluğu (standart minimum 4)
    )
    qr.add_data(data)
    qr.make(fit=True)  # Veri uzunluğuna göre boyutu otomatik seç
    qr.make_image(fill_color="black", back_color="white").save(output_path)
    return output_path


