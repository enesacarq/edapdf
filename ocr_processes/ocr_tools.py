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

# ----------------------------------------------------------------------
# PENCERE ARAYÜZÜ (tkinter)
# ----------------------------------------------------------------------
# Dosya seçme pencerelerinde görünecek dosya türleri
_GORSELLER = "*.jpg *.jpeg *.png *.tif *.tiff *.bmp *.webp"
_TIP_GORSEL = [("Görseller", _GORSELLER), ("Tüm dosyalar", "*.*")]
_TIP_PDF = [("PDF", "*.pdf"), ("Tüm dosyalar", "*.*")]
_TIP_PDF_GORSEL = [("PDF ve görseller", "*.pdf " + _GORSELLER), ("Tüm dosyalar", "*.*")]
_TIP_EXCEL = [("Excel", "*.xlsx *.xlsm"), ("Tüm dosyalar", "*.*")]
_TIP_CSV = [("CSV", "*.csv"), ("Tüm dosyalar", "*.*")]

# Menüdeki her işlem için ayarlar:
#   tipler : girdi dosyası seçim penceresinin türleri (None ise girdi metin kutusudur)
#   cikti  : oluşacak dosyanın uzantısı (None ise sonuç ekranda metin olarak gösterilir)
#   dil    : OCR dili seçimi gösterilsin mi
#   sayfa  : Excel sayfa adı kutusu gösterilsin mi
ISLEMLER = {
    "Görselden metin çıkar":   {"tipler": _TIP_GORSEL,     "cikti": None,    "dil": True,  "sayfa": False},
    "PDF'ten metin çıkar":     {"tipler": _TIP_PDF,        "cikti": None,    "dil": True,  "sayfa": False},
    "Aranabilir PDF oluştur":  {"tipler": _TIP_PDF_GORSEL, "cikti": ".pdf",  "dil": True,  "sayfa": False},
    "Excel → CSV":             {"tipler": _TIP_EXCEL,      "cikti": ".csv",  "dil": False, "sayfa": True},
    "CSV → Excel":             {"tipler": _TIP_CSV,        "cikti": ".xlsx", "dil": False, "sayfa": False},
    "QR kod oluştur":          {"tipler": None,            "cikti": ".png",  "dil": False, "sayfa": False},
}


def _klasoru_ac(yol):
    """Verilen dosyanın klasörünü dosya gezgininde açar."""
    klasor = os.path.dirname(os.path.abspath(yol))
    if os.name == "nt":
        os.startfile(klasor)
    elif sys.platform == "darwin":
        subprocess.run(["open", klasor])
    else:
        subprocess.run(["xdg-open", klasor])


class OcrArayuz:
    """Tek pencerelik arayüz: işlem seç, dosya seç, çalıştır."""

    def __init__(self, root):
        self.root = root
        root.title("OCR Araçları")
        root.geometry("760x640")
        root.minsize(660, 540)

        self.kuyruk = queue.Queue()  # Arka plan işinin sonucu buraya gelir
        self.son_dosya = None        # Üretilen son dosya ("Klasörü aç" için)

        # Kutulara bağlı değişkenler
        self.islem = tk.StringVar(value=next(iter(ISLEMLER)))
        self.girdi = tk.StringVar()
        self.cikti = tk.StringVar()
        self.dil = tk.StringVar(value="tur+eng")
        self.sayfa = tk.StringVar()

        ana = ttk.Frame(root, padding=14)
        ana.pack(fill="both", expand=True)
        ana.columnconfigure(1, weight=1)
        ana.rowconfigure(7, weight=1)  # Sonuç kutusu pencereyle birlikte büyüsün

        # Satır 0: işlem seçimi
        ttk.Label(ana, text="İşlem").grid(row=0, column=0, sticky="w", pady=4)
        kutu = ttk.Combobox(ana, textvariable=self.islem, values=list(ISLEMLER),
                            state="readonly")
        kutu.grid(row=0, column=1, columnspan=2, sticky="ew", pady=4)
        kutu.bind("<<ComboboxSelected>>", lambda e: self.islem_degisti())

        # Satır 1: girdi (dosya veya metin)
        self.lbl_girdi = ttk.Label(ana, text="Girdi dosyası")
        self.ent_girdi = ttk.Entry(ana, textvariable=self.girdi)
        self.btn_girdi = ttk.Button(ana, text="Gözat...", command=self.girdi_sec)

        # Satır 2: çıktı dosyası
        self.lbl_cikti = ttk.Label(ana, text="Çıktı dosyası")
        self.ent_cikti = ttk.Entry(ana, textvariable=self.cikti)
        self.btn_cikti = ttk.Button(ana, text="Kaydet yeri...", command=self.cikti_sec)

        # Satır 3: dil
        self.lbl_dil = ttk.Label(ana, text="OCR dili")
        self.cmb_dil = ttk.Combobox(ana, textvariable=self.dil, state="readonly",
                                    values=["tur+eng", "tur", "eng"])

        # Satır 4: Excel sayfa adı
        self.lbl_sayfa = ttk.Label(ana, text="Sayfa adı (boşsa ilk sayfa)")
        self.ent_sayfa = ttk.Entry(ana, textvariable=self.sayfa)

        # Satır 5: çalıştır, satır 6: ilerleme
        self.btn_calistir = ttk.Button(ana, text="Çalıştır", command=self.calistir)
        self.btn_calistir.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(12, 4))
        self.cubuk = ttk.Progressbar(ana, mode="indeterminate")
        self.cubuk.grid(row=6, column=0, columnspan=3, sticky="ew", pady=4)

        # Satır 7: sonuç kutusu (düzenlenebilir: OCR hatalarını elle düzeltebilirsin)
        cerceve = ttk.Frame(ana)
        cerceve.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=4)
        cerceve.columnconfigure(0, weight=1)
        cerceve.rowconfigure(0, weight=1)
        self.sonuc = tk.Text(cerceve, wrap="word", height=10)
        kaydirma = ttk.Scrollbar(cerceve, command=self.sonuc.yview)
        self.sonuc.configure(yscrollcommand=kaydirma.set)
        self.sonuc.grid(row=0, column=0, sticky="nsew")
        kaydirma.grid(row=0, column=1, sticky="ns")

        # Satır 8: sonuç düğmeleri
        alt = ttk.Frame(ana)
        alt.grid(row=8, column=0, columnspan=3, sticky="w", pady=(4, 0))
        self.btn_kopyala = ttk.Button(alt, text="Metni kopyala", command=self.kopyala)
        self.btn_txt = ttk.Button(alt, text="TXT olarak kaydet", command=self.txt_kaydet)
        self.btn_klasor = ttk.Button(alt, text="Klasörü aç", command=self.klasoru_ac)
        for b in (self.btn_kopyala, self.btn_txt, self.btn_klasor):
            b.pack(side="left", padx=(0, 6))

        self.islem_degisti()

    # ---------- Görünümü işleme göre ayarla ----------
    def ayar(self):
        return ISLEMLER[self.islem.get()]

    def satir_goster(self, satir, widgetlar, goster):
        """Bir satırdaki kutuları göster/gizle."""
        for sutun, w in enumerate(widgetlar):
            if goster:
                if len(widgetlar) == 2:  # etiket + tek kutu: kutu iki sütuna yayılsın
                    w.grid(row=satir, column=sutun, columnspan=1 if sutun == 0 else 2,
                           sticky="w" if sutun == 0 else "ew", pady=4, padx=(0, 6))
                else:
                    w.grid(row=satir, column=sutun, sticky="w" if sutun == 0 else "ew",
                           pady=4, padx=(0, 6))
            else:
                w.grid_remove()

    def islem_degisti(self):
        a = self.ayar()
        qr = a["tipler"] is None
        self.lbl_girdi.config(text="QR'a yazılacak metin / URL" if qr else "Girdi dosyası")
        self.satir_goster(1, [self.lbl_girdi, self.ent_girdi, self.btn_girdi], True)
        if qr:
            self.btn_girdi.grid_remove()           # metin girilecek, Gözat gerekmez
            self.ent_girdi.grid_configure(columnspan=2)
        else:
            self.ent_girdi.grid_configure(columnspan=1)
        self.satir_goster(2, [self.lbl_cikti, self.ent_cikti, self.btn_cikti], a["cikti"] is not None)
        self.satir_goster(3, [self.lbl_dil, self.cmb_dil], a["dil"])
        self.satir_goster(4, [self.lbl_sayfa, self.ent_sayfa], a["sayfa"])
        self.girdi.set("")
        self.cikti.set("")
        self.sonuc_temizle()

    def sonuc_temizle(self):
        self.sonuc.delete("1.0", "end")
        self.son_dosya = None
        self.butonlari_ayarla(metin=False, dosya=False)

    def butonlari_ayarla(self, metin, dosya):
        """Sonuç düğmelerini duruma göre aç/kapat."""
        self.btn_kopyala.state(["!disabled"] if metin else ["disabled"])
        self.btn_txt.state(["!disabled"] if metin else ["disabled"])
        self.btn_klasor.state(["!disabled"] if dosya else ["disabled"])

    # ---------- Dosya seçme ----------
    def girdi_sec(self):
        yol = filedialog.askopenfilename(title="Dosya seç", filetypes=self.ayar()["tipler"])
        if yol:
            self.girdi.set(yol)
            self.cikti_oner()

    def cikti_oner(self):
        """Çıktı için otomatik bir isim önerir (girdi ile aynı klasörde)."""
        uzanti = self.ayar()["cikti"]
        if not uzanti or not self.girdi.get():
            return
        p = Path(self.girdi.get())
        if self.islem.get() == "Aranabilir PDF oluştur":
            self.cikti.set(str(p.with_name(p.stem + "_aranabilir.pdf")))
        elif self.ayar()["tipler"] is None:  # QR: girdi metindir, dosya adı sabit
            self.cikti.set(str(Path.cwd() / "qr.png"))
        else:
            self.cikti.set(str(p.with_suffix(uzanti)))

    def cikti_sec(self):
        uzanti = self.ayar()["cikti"]
        yol = filedialog.asksaveasfilename(title="Nereye kaydedilsin?",
                                           defaultextension=uzanti,
                                           initialfile=os.path.basename(self.cikti.get()))
        if yol:
            self.cikti.set(yol)

    # ---------- Çalıştırma ----------
    def calistir(self):
        ad, a = self.islem.get(), self.ayar()
        girdi = self.girdi.get().strip()
        cikti = self.cikti.get().strip()

        # Kutular dolu mu diye kontrol et
        if not girdi:
            messagebox.showwarning("Eksik bilgi", "Önce girdi dosyasını seç (veya metni yaz).")
            return
        if a["tipler"] is not None and not os.path.isfile(girdi):
            messagebox.showwarning("Dosya yok", f"Dosya bulunamadı:\n{girdi}")
            return
        if a["cikti"] and not cikti:
            self.cikti_oner()
            cikti = self.cikti.get().strip()

        self.sonuc_temizle()
        self.btn_calistir.state(["disabled"])
        self.cubuk.start(12)

        dil, sayfa = self.dil.get(), self.sayfa.get().strip() or None

        # Uzun işi ayrı iş parçacığında çalıştır; yoksa pencere donar.
        def is_parcacigi():
            try:
                if ad == "Görselden metin çıkar":
                    sonuc = ("metin", extract_text_from_image(girdi, dil))
                elif ad == "PDF'ten metin çıkar":
                    sonuc = ("metin", extract_text_from_pdf(girdi, dil))
                elif ad == "Aranabilir PDF oluştur":
                    sonuc = ("dosya", create_searchable_pdf(girdi, cikti, dil))
                elif ad == "Excel → CSV":
                    sonuc = ("dosya", excel_to_csv(girdi, cikti, sayfa))
                elif ad == "CSV → Excel":
                    sonuc = ("dosya", csv_to_excel(girdi, cikti))
                else:
                    sonuc = ("dosya", create_qr_code(girdi, cikti))
                self.kuyruk.put(("ok",) + sonuc)
            except Exception as hata:
                self.kuyruk.put(("hata", str(hata), None))

        threading.Thread(target=is_parcacigi, daemon=True).start()
        self.root.after(150, self.sonucu_bekle)

    def sonucu_bekle(self):
        """Arka plan işi bitti mi diye periyodik bakar."""
        try:
            durum, tur_veya_mesaj, deger = self.kuyruk.get_nowait()
        except queue.Empty:
            self.root.after(150, self.sonucu_bekle)
            return

        self.cubuk.stop()
        self.btn_calistir.state(["!disabled"])
        if durum == "hata":
            self.sonuc.insert("1.0", "HATA:\n" + tur_veya_mesaj)
            messagebox.showerror("Hata", tur_veya_mesaj[:600])
        elif tur_veya_mesaj == "metin":
            self.sonuc.insert("1.0", deger if deger else "(Metin bulunamadı)")
            self.butonlari_ayarla(metin=bool(deger), dosya=False)
        else:
            self.son_dosya = deger
            self.sonuc.insert("1.0", f"Tamamlandı.\nOluşturulan dosya:\n{deger}")
            self.butonlari_ayarla(metin=False, dosya=True)

    # ---------- Sonuç düğmeleri ----------
    def kopyala(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.sonuc.get("1.0", "end-1c"))

    def txt_kaydet(self):
        yol = filedialog.asksaveasfilename(defaultextension=".txt",
                                           filetypes=[("Metin dosyası", "*.txt")])
        if yol:
            with open(yol, "w", encoding="utf-8") as f:
                f.write(self.sonuc.get("1.0", "end-1c"))

    def klasoru_ac(self):
        if self.son_dosya:
            _klasoru_ac(self.son_dosya)


def _gui():
    """Pencereyi açar."""
    if tk is None:
        sys.exit("tkinter bulunamadı. Python'u yeniden kurarken 'tcl/tk' seçeneğini işaretle.")
    # Windows'ta yazıların bulanık görünmesini engelle
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    OcrArayuz(root)
    root.mainloop()


# ----------------------------------------------------------------------
# KOMUT SATIRI ARAYÜZÜ
# ----------------------------------------------------------------------
def _main():
    p = argparse.ArgumentParser(description="OCR ve dönüştürme araçları")
    alt = p.add_subparsers(dest="komut", required=True)

    a = alt.add_parser("gorsel-metin");  a.add_argument("girdi")
    a = alt.add_parser("pdf-metin");     a.add_argument("girdi")
    a = alt.add_parser("aranabilir-pdf"); a.add_argument("girdi"); a.add_argument("cikti")
    a = alt.add_parser("excel-csv");     a.add_argument("girdi"); a.add_argument("cikti"); a.add_argument("sayfa", nargs="?")
    a = alt.add_parser("csv-excel");     a.add_argument("girdi"); a.add_argument("cikti")
    a = alt.add_parser("qr");            a.add_argument("veri");  a.add_argument("cikti")
    for ad in ("gorsel-metin", "pdf-metin", "aranabilir-pdf"):
        alt.choices[ad].add_argument("--dil", default="tur+eng", help="OCR dili (varsayılan tur+eng)")

    s = p.parse_args()
    if s.komut == "gorsel-metin":
        print(extract_text_from_image(s.girdi, s.dil))
    elif s.komut == "pdf-metin":
        print(extract_text_from_pdf(s.girdi, s.dil))
    elif s.komut == "aranabilir-pdf":
        print("Oluşturuldu:", create_searchable_pdf(s.girdi, s.cikti, s.dil))
    elif s.komut == "excel-csv":
        print("Oluşturuldu:", excel_to_csv(s.girdi, s.cikti, s.sayfa))
    elif s.komut == "csv-excel":
        print("Oluşturuldu:", csv_to_excel(s.girdi, s.cikti))
    elif s.komut == "qr":
        print("Oluşturuldu:", create_qr_code(s.veri, s.cikti))


if __name__ == "__main__":
    # Argüman verilmişse komut satırı, verilmemişse pencere arayüzü
    if len(sys.argv) > 1:
        _main()
    else:
        _gui()

