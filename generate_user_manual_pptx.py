import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------
# COLOR PALETTE (Matching Sampahahayuk Website & Canva Theme)
# ---------------------------------------------------------
NAVY_DARK     = RGBColor(15, 23, 42)      # #0F172A (Deep Slate Navy)
NAVY_PRIMARY  = RGBColor(30, 58, 95)      # #1E3A5F (Website Sidebar Dark)
BLUE_BRAND    = RGBColor(37, 99, 235)     # #2563EB (Website Primary Dark)
BLUE_LIGHT    = RGBColor(59, 130, 246)    # #3B82F6 (Website Primary)
BG_ICE        = RGBColor(240, 247, 255)   # #F0F7FF (Website Body Background)
CARD_BG       = RGBColor(255, 255, 255)   # #FFFFFF (Pure White)
BORDER_COLOR  = RGBColor(219, 234, 254)   # #DBEAFE (Subtle Light Blue Border)
TEXT_MAIN     = RGBColor(15, 23, 42)      # #0F172A (Main Heading/Body)
TEXT_MUTED    = RGBColor(100, 116, 139)   # #64748B (Secondary / Muted)
GREEN_ACCENT  = RGBColor(16, 185, 129)   # #10B981 (Approved / Success)
AMBER_ACCENT  = RGBColor(245, 158, 11)    # #F59E0B (Pending / Warning)
RED_ACCENT    = RGBColor(239, 68, 68)     # #EF4444 (Rejected / Danger)

SCREENSHOTS_DIR = os.path.join(os.getcwd(), "screenshots_manual")
OUTPUT_PPTX     = os.path.join(os.getcwd(), "User_Manual_Sampahahayuk.pptx")

# ---------------------------------------------------------
# HELPER FUNCTIONS FOR CLEAN STYLING
# ---------------------------------------------------------
def create_base_slide(prs, bg_color=BG_ICE):
    """Creates a slide with a clean solid background."""
    blank_layout = prs.slide_layouts[6] # Blank slide
    slide = prs.slides.add_slide(blank_layout)
    
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = bg_color
    bg.line.fill.background()
    return slide

def add_header(slide, category, title, subtitle=None):
    """Adds a standard modern minimalist header to the slide."""
    header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
    tf = header_box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0)
    tf.margin_top = Inches(0)
    tf.margin_right = Inches(0)
    tf.margin_bottom = Inches(0)

    # Category / Pill Text
    p0 = tf.paragraphs[0]
    p0.text = category.upper()
    p0.font.size = Pt(11)
    p0.font.bold = True
    p0.font.color.rgb = BLUE_BRAND
    p0.space_after = Pt(4)

    # Title
    p1 = tf.add_paragraph()
    p1.text = title
    p1.font.size = Pt(22)
    p1.font.bold = True
    p1.font.color.rgb = NAVY_DARK
    
    # Subtitle (if any)
    if subtitle:
        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(12)
        p2.font.color.rgb = TEXT_MUTED

def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=BORDER_COLOR):
    """Adds a rounded clean white card."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)
    else:
        card.line.fill.background()
    return card

def add_callout_box(slide, left, top, width, height, title, items, is_alert=False):
    """Adds a stylish 'Ketentuan & Aturan' or info box."""
    bg_col = RGBColor(254, 242, 242) if is_alert else RGBColor(239, 246, 255)
    border_col = RGBColor(254, 202, 202) if is_alert else RGBColor(191, 219, 254)
    title_col = RED_ACCENT if is_alert else BLUE_BRAND

    box = add_card(slide, left, top, width, height, bg_color=bg_col, border_color=border_col)
    
    tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), height - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = title_col
    p0.space_after = Pt(6)

    for item in items:
        p = tf.add_paragraph()
        p.text = f"•  {item}"
        p.font.size = Pt(10.5)
        p.font.color.rgb = NAVY_DARK
        p.space_after = Pt(4)

def add_step_badge(slide, left, top, num_str, title, subtitle=None, width=Inches(3.2), height=Inches(1.0)):
    """Creates a step block with a numbered circular badge."""
    card = add_card(slide, left, top, width, height)
    
    # Circle badge
    badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.15), top + Inches(0.15), Inches(0.45), Inches(0.45))
    badge.fill.solid()
    badge.fill.fore_color.rgb = BLUE_BRAND
    badge.line.fill.background()
    
    btf = badge.text_frame
    btf.text = str(num_str)
    btf.paragraphs[0].font.size = Pt(12)
    btf.paragraphs[0].font.bold = True
    btf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    btf.paragraphs[0].alignment = PP_ALIGN.CENTER

    tb = slide.shapes.add_textbox(left + Inches(0.7), top + Inches(0.12), width - Inches(0.8), height - Inches(0.24))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.size = Pt(11.5)
    p0.font.bold = True
    p0.font.color.rgb = NAVY_DARK
    
    if subtitle:
        p1 = tf.add_paragraph()
        p1.text = subtitle
        p1.font.size = Pt(9.5)
        p1.font.color.rgb = TEXT_MUTED
        p1.space_before = Pt(2)

def add_arrow(slide, left, top, width=Inches(0.4), height=Inches(0.25), color=BLUE_LIGHT):
    """Draws a crisp directional right arrow."""
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, left, top, width, height)
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = color
    arrow.line.fill.background()
    return arrow

def add_screenshot_frame(slide, img_path, left, top, width, height):
    """Adds a screenshot with an elegant browser mockup frame and subtle shadow border."""
    # Outer frame
    frame = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    frame.fill.solid()
    frame.fill.fore_color.rgb = RGBColor(241, 245, 249)
    frame.line.color.rgb = BORDER_COLOR
    frame.line.width = Pt(1.5)
    
    # Browser dots
    dot_y = top + Inches(0.1)
    for i, dot_col in enumerate([RGBColor(239, 68, 68), RGBColor(245, 158, 11), RGBColor(16, 185, 129)]):
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, left + Inches(0.15 + i * 0.18), dot_y, Inches(0.1), Inches(0.1))
        dot.fill.solid()
        dot.fill.fore_color.rgb = dot_col
        dot.line.fill.background()

    # Screenshot inside
    img_top = top + Inches(0.26)
    img_height = height - Inches(0.28)
    img_left = left + Inches(0.04)
    img_width = width - Inches(0.08)
    
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, img_left, img_top, width=img_width, height=img_height)

# ---------------------------------------------------------
# BUILD PRESENTATION
# ---------------------------------------------------------
def generate_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    print("Building slides...")

    # =========================================================
    # SLIDE 1: COVER (Modern Navy & Clean Minimalist Canva Look)
    # =========================================================
    s1 = create_base_slide(prs, bg_color=NAVY_DARK)

    # Decorative background shapes
    decor = s1.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(8.5), Inches(0), Inches(4.833), Inches(7.5))
    decor.fill.solid()
    decor.fill.fore_color.rgb = NAVY_PRIMARY
    decor.line.fill.background()
    decor.rotation = 180

    # Content Card Left
    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.3), Inches(6.8), Inches(4.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    
    # Pill Tag
    p_tag = tf1.paragraphs[0]
    p_tag.text = "PANDUAN OPERASIONAL RESMI  •  VERSI 1.0"
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = BLUE_LIGHT
    p_tag.space_after = Pt(16)

    # Big Title
    p_title = tf1.add_paragraph()
    p_title.text = "USER MANUAL\nSISTEM BANK SAMPAH"
    p_title.font.size = Pt(36)
    p_title.font.bold = True
    p_title.font.color.rgb = RGBColor(255, 255, 255)
    p_title.space_after = Pt(16)

    # Subtitle
    p_sub = tf1.add_paragraph()
    p_sub.text = "Panduan Praktis Langkah Demi Langkah Penggunaan Aplikasi Web Sampahahayuk untuk 2 Peran Utama: Pengguna (Nasabah) & Administrator."
    p_sub.font.size = Pt(14)
    p_sub.font.color.rgb = RGBColor(203, 213, 225)
    p_sub.space_after = Pt(28)

    # Feature points with arrows
    p_features = tf1.add_paragraph()
    p_features.text = "➔ Alur Setor Sampah  |  ➔ Konversi Saldo  |  ➔ Verifikasi Laporan  |  ➔ Ketentuan SOP"
    p_features.font.size = Pt(11.5)
    p_features.font.bold = True
    p_features.font.color.rgb = RGBColor(147, 197, 253)

    # Screenshot Mockup Right
    img_cov = os.path.join(SCREENSHOTS_DIR, "03_user_dashboard.png")
    if os.path.exists(img_cov):
        add_screenshot_frame(s1, img_cov, Inches(7.5), Inches(1.2), Inches(4.9), Inches(4.9))

    # =========================================================
    # SLIDE 2: DAFTAR ISI & STRUKTUR 2 PERAN
    # =========================================================
    s2 = create_base_slide(prs)
    add_header(s2, "IKHTISAR SISTEM", "Struktur Manual & Perbandingan 2 Peran Pengguna", 
               "Aplikasi Sampahahayuk mendigitalkan proses bank sampah dari penyetoran nasabah hingga validasi admin.")

    # Top Flow Navigation (Canva arrow style)
    modules = [
        ("01. Akses Akun", "Registrasi & Login"),
        ("02. Role Pengguna", "Setor Sampah & Cek Saldo"),
        ("03. Role Admin", "Validasi & Kelola Master Data"),
        ("04. SOP & Ketentuan", "Aturan & Standar Teknis")
    ]
    for i, (m_title, m_sub) in enumerate(modules):
        x = Inches(0.8 + i * 2.95)
        add_step_badge(s2, x, Inches(1.6), str(i+1), m_title, m_sub, width=Inches(2.5), height=Inches(0.85))
        if i < 3:
            add_arrow(s2, x + Inches(2.55), Inches(1.9), width=Inches(0.35), height=Inches(0.2))

    # Two Role Comparison Cards
    # Card 1: Role Nasabah
    c_user = add_card(s2, Inches(0.8), Inches(2.7), Inches(5.6), Inches(4.3))
    tb_u = s2.shapes.add_textbox(Inches(1.1), Inches(2.9), Inches(5.0), Inches(3.9))
    tf_u = tb_u.text_frame
    tf_u.word_wrap = True
    
    p = tf_u.paragraphs[0]
    p.text = "👤 ROLE 1: PENGGUNA (NASABAH)"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = BLUE_BRAND
    p.space_after = Pt(10)

    nasabah_points = [
        ("Tujuan Utama", "Mendaftarkan diri, menyetor sampah kering/terpilah, dan mengumpulkan saldo tabungan."),
        ("Fitur Akses", "Dashboard pribadi, formulir setor sampah, pemantau estimasi harga, riwayat status."),
        ("Alur Kunci", "Pilih jenis sampah ➔ Masukkan berat ➔ Upload foto bukti ➔ Pantau approval ➔ Saldo masuk."),
        ("Hak Finansial", "Memperoleh saldo rupiah otomatis setelah laporan disetujui oleh tim verifikator admin.")
    ]
    for k, v in nasabah_points:
        p_k = tf_u.add_paragraph()
        p_k.text = f"•  {k}: {v}"
        p_k.font.size = Pt(11)
        p_k.font.color.rgb = NAVY_DARK
        p_k.space_after = Pt(6)

    # Card 2: Role Admin
    c_adm = add_card(s2, Inches(6.9), Inches(2.7), Inches(5.6), Inches(4.3))
    tb_a = s2.shapes.add_textbox(Inches(7.2), Inches(2.9), Inches(5.0), Inches(3.9))
    tf_a = tb_a.text_frame
    tf_a.word_wrap = True
    
    p = tf_a.paragraphs[0]
    p.text = "🛡️ ROLE 2: ADMINISTRATOR (PENGELOLA)"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY
    p.space_after = Pt(10)

    admin_points = [
        ("Tujuan Utama", "Memvalidasi setoran sampah fisik, memantau keuangan saldo, dan mengelola operasional."),
        ("Fitur Akses", "Dashboard statistik, verifikasi laporan pending, master jenis sampah & tarif, manajemen wilayah & user."),
        ("Alur Kunci", "Periksa fisik/foto ➔ Ubah status PENDING ke APPROVED/REJECTED ➔ Saldo user ter-update."),
        ("Wewenang", "Penentu kelayakan setoran sampah, pengatur tarif beli per kg, dan pembuat wilayah operasional.")
    ]
    for k, v in admin_points:
        p_k = tf_a.add_paragraph()
        p_k.text = f"•  {k}: {v}"
        p_k.font.size = Pt(11)
        p_k.font.color.rgb = NAVY_DARK
        p_k.space_after = Pt(6)

    # =========================================================
    # SLIDE 3: MODUL 1 - AKSES SISTEM (REGISTRASI & LOGIN)
    # =========================================================
    s3 = create_base_slide(prs)
    add_header(s3, "MODUL 1: AKSES SISTEM", "Langkah Registrasi Akun & Login Aplikasi", 
               "Langkah awal bagi seluruh pengguna untuk mengakses dashboard sesuai peran yang dimiliki.")

    # Flow Banner
    steps_auth = [
        ("1. Buka Web", "Akses /login"),
        ("2. Pilih Menu", "Klik tombol 'Daftar'"),
        ("3. Input Data", "Nama, Email, HP, Sandi"),
        ("4. Masuk Akun", "Redirect otomatis")
    ]
    for i, (st, ss) in enumerate(steps_auth):
        x = Inches(0.8 + i * 2.95)
        add_step_badge(s3, x, Inches(1.5), str(i+1), st, ss, width=Inches(2.5), height=Inches(0.75))
        if i < 3:
            add_arrow(s3, x + Inches(2.55), Inches(1.8), width=Inches(0.35), height=Inches(0.18))

    # Screenshots: Login & Register
    img_login = os.path.join(SCREENSHOTS_DIR, "01_login.png")
    img_reg   = os.path.join(SCREENSHOTS_DIR, "02_register.png")
    add_screenshot_frame(s3, img_login, Inches(0.8), Inches(2.45), Inches(4.3), Inches(4.6))
    add_screenshot_frame(s3, img_reg, Inches(5.3), Inches(2.45), Inches(4.3), Inches(4.6))

    # Ketentuan Box Right
    ketentuan_auth = [
        "Nomor HP & Email Unik: Satu nomor HP dan email hanya dapat digunakan untuk 1 akun terdaftar.",
        "Keamanan Password: Wajib minimal 6 karakter. Dianjurkan kombinasi huruf dan angka.",
        "Role Default Pendaftar: Setiap pengguna yang mendaftar melalui form registrasi otomatis memiliki role 'USER' (Nasabah).",
        "Akses Khusus Admin: Akun Administrator dibuat melalui seeding database internal demi alasan keamanan data.",
        "Sesi Login Aman: Sistem menggunakan HTTP-only JWT Cookie aktif hingga 7 hari."
    ]
    add_callout_box(s3, Inches(9.8), Inches(2.45), Inches(2.7), Inches(4.6), "📌 KETENTUAN AKSES", ketentuan_auth)

    # =========================================================
    # SLIDE 4: ROLE NASABAH - WORKFLOW UTAMA PENYETORAN SAMPAH
    # =========================================================
    s4 = create_base_slide(prs)
    add_header(s4, "ROLE PENGGUNA (NASABAH)", "Alur Kerja Utama: Dari Sampah Menjadi Saldo", 
               "Diagram alur lengkap proses penyetoran sampah kering bernilai ekonomis hingga pencairan saldo.")

    # 5 Major Steps with Big Arrows
    workflow_steps = [
        ("LANGKAH 1", "Pilah & Timbang", "Pilah sampah sesuai kategori (Plastik, Kertas, Logam) lalu timbang."),
        ("LANGKAH 2", "Buat Laporan", "Buka menu '+ Create Report', pilih jenis sampah dan wilayah jemput."),
        ("LANGKAH 3", "Upload Foto", "Lampirkan foto sampah sebagai bukti fisik lalu submit formulir."),
        ("LANGKAH 4", "Verifikasi Admin", "Status laporan 'PENDING'. Admin memeriksa kondisi sampah."),
        ("LANGKAH 5", "Saldo Cair", "Status jadi 'APPROVED'. Saldo otomatis masuk ke akun nasabah!")
    ]
    for i, (step_num, step_title, step_desc) in enumerate(workflow_steps):
        x = Inches(0.8 + i * 2.4)
        c = add_card(s4, x, Inches(1.6), Inches(2.15), Inches(2.2))
        
        # Step header
        badge = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.15), Inches(1.75), Inches(1.85), Inches(0.35))
        badge.fill.solid()
        badge.fill.fore_color.rgb = BLUE_BRAND if i < 4 else GREEN_ACCENT
        badge.line.fill.background()
        btf = badge.text_frame
        btf.text = step_num
        btf.paragraphs[0].font.size = Pt(10)
        btf.paragraphs[0].font.bold = True
        btf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        btf.paragraphs[0].alignment = PP_ALIGN.CENTER

        tb = s4.shapes.add_textbox(x + Inches(0.15), Inches(2.2), Inches(1.85), Inches(1.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        p0.text = step_title
        p0.font.size = Pt(12)
        p0.font.bold = True
        p0.font.color.rgb = NAVY_DARK
        p0.space_after = Pt(4)

        p1 = tf.add_paragraph()
        p1.text = step_desc
        p1.font.size = Pt(9.5)
        p1.font.color.rgb = TEXT_MUTED

        if i < 4:
            add_arrow(s4, x + Inches(2.18), Inches(2.5), width=Inches(0.2), height=Inches(0.25))

    # Bottom Cards: Ketentuan & Standar Mutu
    ketentuan_sampah = [
        "Kondisi Sampah: Wajib bersih, bebas cairan, dan tidak tercampur sampah basah/organik/B3.",
        "Pemilahan Kategori: Botol plastik harus dikosongkan & dipipihkan; kardus/kertas diikat rapi.",
        "Satuan Berat: Timbangan menggunakan satuan Kilogram (Kg) dengan presisi hingga 1 desimal.",
        "Batas Minimal: Minimal penyetoran adalah 0.1 Kg per transaksi laporan.",
        "Perubahan Saldo: Saldo TIDAK akan bertambah sebelum admin melakukan verifikasi fisik & approval."
    ]
    add_callout_box(s4, Inches(0.8), Inches(4.1), Inches(6.8), Inches(2.9), "⚠️ KETENTUAN MUTLAK PENYETORAN", ketentuan_sampah, is_alert=True)

    aturan_tambahan = [
        "1 Transaksi = 1 Jenis: Untuk mempermudah perhitungan, 1 laporan hanya untuk 1 jenis sampah.",
        "Estimasi vs Riil: Total rupiah yang muncul saat input adalah estimasi sistem (harga x berat).",
        "Hak Admin: Admin berhak menolak (REJECT) jika sampah tidak sesuai foto atau kotor/rusak parah.",
        "Pencatatan Permanen: Seluruh transaksi yang disetujui akan tersimpan selamanya pada riwayat."
    ]
    add_callout_box(s4, Inches(7.8), Inches(4.1), Inches(4.7), Inches(2.9), "📌 STANDAR OPERASIONAL (SOP)", aturan_tambahan)

    # =========================================================
    # SLIDE 5: ROLE NASABAH - MEMAHAMI DASHBOARD PENGGUNA
    # =========================================================
    s5 = create_base_slide(prs)
    add_header(s5, "ROLE PENGGUNA (NASABAH)", "Tampilan Dashboard & Monitoring Saldo Nasabah", 
               "Pusat kendali nasabah untuk memantau tabungan sampah, status laporan terkini, dan navigasi.")

    # Screenshot Dashboard
    img_udash = os.path.join(SCREENSHOTS_DIR, "03_user_dashboard.png")
    add_screenshot_frame(s5, img_udash, Inches(0.8), Inches(1.6), Inches(7.8), Inches(5.3))

    # Callout Indicators Left/Right
    card_info = add_card(s5, Inches(8.8), Inches(1.6), Inches(3.7), Inches(5.3))
    tb_ci = s5.shapes.add_textbox(Inches(9.0), Inches(1.8), Inches(3.3), Inches(4.9))
    tf_ci = tb_ci.text_frame
    tf_ci.word_wrap = True

    p = tf_ci.paragraphs[0]
    p.text = "🔍 KOMPONEN DASHBOARD"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = BLUE_BRAND
    p.space_after = Pt(10)

    dash_components = [
        ("💰 Saldo Simpanan (Rp)", "Menampilkan saldo tabungan bersih yang terkumpul dari sampah yang berhasil diverifikasi."),
        ("➕ Tombol '+ Create Report'", "Jalan pintas cepat untuk membuka formulir penyetoran sampah baru."),
        ("📋 Kartu 'Latest Reports'", "Menampilkan 6 laporan terbaru beserta jenis sampah, wilayah jemput, berat, dan status terkini."),
        ("🏷️ Status Badge (PENDING / APPROVED)", "Warna kuning menandakan menunggu verifikasi, warna biru/hijau menandakan telah disetujui."),
        ("🚪 Tombol Logout & Profil", "Menampilkan nama pengguna aktif dan akses keluar akun dengan aman.")
    ]
    for title_c, desc_c in dash_components:
        p_t = tf_ci.add_paragraph()
        p_t.text = f"➔ {title_c}"
        p_t.font.size = Pt(10.5)
        p_t.font.bold = True
        p_t.font.color.rgb = NAVY_DARK
        
        p_d = tf_ci.add_paragraph()
        p_d.text = desc_c
        p_d.font.size = Pt(9.5)
        p_d.font.color.rgb = TEXT_MUTED
        p_d.space_after = Pt(6)

    # =========================================================
    # SLIDE 6: ROLE NASABAH - FORMULIR SETOR SAMPAH (LANGKAH INPUT)
    # =========================================================
    s6 = create_base_slide(prs)
    add_header(s6, "ROLE PENGGUNA (NASABAH)", "Langkah Pengisian Formulir Setor Sampah", 
               "Panduan mengisi formulir laporan setor sampah secara presisi agar cepat diverifikasi petugas.")

    # Screenshot Form Input
    img_form = os.path.join(SCREENSHOTS_DIR, "04_user_form_input.png")
    add_screenshot_frame(s6, img_form, Inches(0.8), Inches(1.6), Inches(7.5), Inches(5.3))

    # Right side: Form Fields Guide with Arrows & Requirements
    card_form = add_card(s6, Inches(8.5), Inches(1.6), Inches(4.0), Inches(5.3))
    tb_fg = s6.shapes.add_textbox(Inches(8.7), Inches(1.75), Inches(3.6), Inches(5.0))
    tf_fg = tb_fg.text_frame
    tf_fg.word_wrap = True

    p = tf_fg.paragraphs[0]
    p.text = "📝 DETAIL LANGKAH PENGISIAN"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = BLUE_BRAND
    p.space_after = Pt(8)

    fields_guide = [
        ("1. Pilih Jenis Sampah *", "Pilih kategori sampah yang akan disetor. Harga beli per kg akan otomatis tertera di sistem."),
        ("2. Pilih Wilayah Penjemputan *", "Pilih kelurahan/kecamatan titik lokasi keberadaan sampah untuk kemudahan rute armada."),
        ("3. Input Estimasi Berat (Kg) *", "Masukkan berat riil sampah dalam angka desimal (contoh: 2.5 atau 10). Sistem otomatis mengalkulasi total harga."),
        ("4. Upload Foto Bukti Sampah *", "Pilih file foto sampah fisik sebagai bukti autentik bagi tim verifikator admin."),
        ("5. Klik 'Kirim Laporan'", "Data langsung tersimpan di sistem dengan status awal 'PENDING'.")
    ]
    for step_f, desc_f in fields_guide:
        p_sf = tf_fg.add_paragraph()
        p_sf.text = step_f
        p_sf.font.size = Pt(10.5)
        p_sf.font.bold = True
        p_sf.font.color.rgb = NAVY_DARK
        
        p_df = tf_fg.add_paragraph()
        p_df.text = desc_f
        p_df.font.size = Pt(9.2)
        p_df.font.color.rgb = TEXT_MUTED
        p_df.space_after = Pt(4)

    # Mini Alert Box inside
    p_alert = tf_fg.add_paragraph()
    p_alert.text = "📌 KETENTUAN FOTO: Format .JPG / .PNG, pencahayaan jelas, sampah terlihat utuh, maksimal ukuran file 5 MB."
    p_alert.font.size = Pt(9.5)
    p_alert.font.bold = True
    p_alert.font.color.rgb = RED_ACCENT
    p_alert.space_before = Pt(6)

    # =========================================================
    # SLIDE 7: ROLE NASABAH - RIWAYAT TRANSAKSI & STATUS VERIFIKASI
    # =========================================================
    s7 = create_base_slide(prs)
    add_header(s7, "ROLE PENGGUNA (NASABAH)", "Pemantauan Riwayat & Siklus Status Transaksi", 
               "Bagaimana memantau kemajuan laporan penyetoran Anda dan arti masing-masing status verifikasi.")

    # Screenshot Riwayat
    img_riwayat = os.path.join(SCREENSHOTS_DIR, "05_user_riwayat.png")
    add_screenshot_frame(s7, img_riwayat, Inches(0.8), Inches(1.6), Inches(7.5), Inches(5.3))

    # Right side: Status Lifecycle with Arrows
    card_stat = add_card(s7, Inches(8.5), Inches(1.6), Inches(4.0), Inches(5.3))
    tb_st = s7.shapes.add_textbox(Inches(8.7), Inches(1.75), Inches(3.6), Inches(5.0))
    tf_st = tb_st.text_frame
    tf_st.word_wrap = True

    p = tf_st.paragraphs[0]
    p.text = "🔄 SIKLUS STATUS LAPORAN"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = BLUE_BRAND
    p.space_after = Pt(10)

    statuses = [
        ("🟡 STATUS: PENDING (Kuning)", 
         "Laporan baru berhasil dikirim. Sedang menunggu jadwal pengecekan & timbang ulang oleh petugas admin. Saldo belum masuk.", 
         AMBER_ACCENT),
        ("🟢 STATUS: APPROVED (Hijau)", 
         "Laporan telah diverifikasi valid! Bobot sampah telah disetujui, dan saldo rupiah LANGSUNG DITAMBAHKAN ke akun Anda secara otomatis.", 
         GREEN_ACCENT),
        ("🔴 STATUS: REJECTED (Merah)", 
         "Laporan ditolak oleh admin (misal: foto tidak valid, sampah basah/terkontaminasi). Saldo tidak bertambah.", 
         RED_ACCENT)
    ]
    for title_s, desc_s, col_s in statuses:
        p_ts = tf_st.add_paragraph()
        p_ts.text = title_s
        p_ts.font.size = Pt(10.5)
        p_ts.font.bold = True
        p_ts.font.color.rgb = col_s
        
        p_ds = tf_st.add_paragraph()
        p_ds.text = desc_s
        p_ds.font.size = Pt(9.5)
        p_ds.font.color.rgb = TEXT_MUTED
        p_ds.space_after = Pt(8)

    p_dt = tf_st.add_paragraph()
    p_dt.text = "➔ Tombol 'Detail Transaksi': Klik untuk melihat rincian nota digital, tanggal transaksi, dan foto dokumentasi."
    p_dt.font.size = Pt(9.5)
    p_dt.font.bold = True
    p_dt.font.color.rgb = NAVY_DARK

    # =========================================================
    # SLIDE 8: ROLE ADMIN - DASHBOARD MONITORING & STATISTIK
    # =========================================================
    s8 = create_base_slide(prs)
    add_header(s8, "ROLE ADMINISTRATOR", "Dashboard Pengelola & Pemantauan Statistik", 
               "Tinjauan eksekutif untuk mengawasi seluruh aktivitas nasabah, volume sampah masuk, dan status operasional.")

    # Screenshot Admin Dashboard
    img_adash = os.path.join(SCREENSHOTS_DIR, "06_admin_dashboard.png")
    add_screenshot_frame(s8, img_adash, Inches(0.8), Inches(1.6), Inches(7.5), Inches(5.3))

    # Right side: Metric breakdown & Admin SOP
    card_ad = add_card(s8, Inches(8.5), Inches(1.6), Inches(4.0), Inches(5.3))
    tb_ad = s8.shapes.add_textbox(Inches(8.7), Inches(1.75), Inches(3.6), Inches(5.0))
    tf_ad = tb_ad.text_frame
    tf_ad.word_wrap = True

    p = tf_ad.paragraphs[0]
    p.text = "📊 INDIKATOR KUNCI ADMIN"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY
    p.space_after = Pt(8)

    admin_metrics = [
        ("Total User & Transaksi", "Memantau jumlah nasabah terdaftar dan volume transaksi finansial yang telah tereksekusi."),
        ("Statistik Jenis Sampah & Wilayah", "Melihat total variasi sampah aktif dan cakupan kelurahan/kecamatan yang terlayani."),
        ("Antrian Status PENDING", "Angka kritis yang menunjukkan berapa banyak setoran baru yang menanti verifikasi petugas."),
        ("Feed 'Aktivitas Terbaru'", "Log waktu-nyata mengenai penyetoran teranyar lengkap dengan nama nasabah dan berat.")
    ]
    for tm, dm in admin_metrics:
        p_tm = tf_ad.add_paragraph()
        p_tm.text = f"•  {tm}"
        p_tm.font.size = Pt(10.2)
        p_tm.font.bold = True
        p_tm.font.color.rgb = NAVY_DARK
        
        p_dm = tf_ad.add_paragraph()
        p_dm.text = dm
        p_dm.font.size = Pt(9.2)
        p_dm.font.color.rgb = TEXT_MUTED
        p_dm.space_after = Pt(4)

    # Admin Duty Rule
    p_duty = tf_ad.add_paragraph()
    p_duty.text = "📌 KETENTUAN SLA ADMIN: Seluruh laporan berstatus PENDING wajib diverifikasi maksimal 1x24 jam kerja agar nasabah segera mendapatkan kepastian saldo."
    p_duty.font.size = Pt(9.5)
    p_duty.font.bold = True
    p_duty.font.color.rgb = BLUE_BRAND
    p_duty.space_before = Pt(6)

    # =========================================================
    # SLIDE 9: ROLE ADMIN - VERIFIKASI & VALIDASI LAPORAN MASUK
    # =========================================================
    s9 = create_base_slide(prs)
    add_header(s9, "ROLE ADMINISTRATOR", "Verifikasi, Validasi Fisik & Approval Setoran", 
               "Prosedur kerja admin dalam memeriksa kesesuaian fisik sampah dan mengeksekusi penambahan saldo nasabah.")

    # Screenshot Admin Laporan
    img_alap = os.path.join(SCREENSHOTS_DIR, "07_admin_laporan.png")
    add_screenshot_frame(s9, img_alap, Inches(0.8), Inches(1.6), Inches(7.5), Inches(5.3))

    # Right side: Verification Workflow with Arrows
    card_vl = add_card(s9, Inches(8.5), Inches(1.6), Inches(4.0), Inches(5.3))
    tb_vl = s9.shapes.add_textbox(Inches(8.7), Inches(1.75), Inches(3.6), Inches(5.0))
    tf_vl = tb_vl.text_frame
    tf_vl.word_wrap = True

    p = tf_vl.paragraphs[0]
    p.text = "⚡ TATA CARA VALIDASI ADMIN"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY
    p.space_after = Pt(8)

    val_steps = [
        ("1. Buka Menu Laporan", "Admin melihat tabel komprehensif seluruh setoran masuk beserta filter pencarian."),
        ("2. Periksa Foto & Bobot Riil", "Klik thumbnail foto untuk memperbesar gambar sampah dan pastikan bobot sesuai timbangan."),
        ("3. Ubah Status Dropdown", "Pilih 'APPROVED' untuk menyetujui, atau 'REJECTED' untuk menolak laporan."),
        ("4. Eksekusi Otomatis Saldo", "Saat status diubah ke 'APPROVED', sistem secara instan mengkredit saldo ke dompet nasabah.")
    ]
    for sv, dv in val_steps:
        p_sv = tf_vl.add_paragraph()
        p_sv.text = sv
        p_sv.font.size = Pt(10.2)
        p_sv.font.bold = True
        p_sv.font.color.rgb = NAVY_DARK
        
        p_dv = tf_vl.add_paragraph()
        p_dv.text = dv
        p_dv.font.size = Pt(9.2)
        p_dv.font.color.rgb = TEXT_MUTED
        p_dv.space_after = Pt(4)

    # Danger Alert
    p_irrev = tf_vl.add_paragraph()
    p_irrev.text = "⚠️ KETENTUAN PENTING: Tindakan 'APPROVED' memicu mutasi finansial saldo riil. Pastikan berat fisik telah terkonfirmasi sebelum mengubah status."
    p_irrev.font.size = Pt(9.5)
    p_irrev.font.bold = True
    p_irrev.font.color.rgb = RED_ACCENT
    p_irrev.space_before = Pt(6)

    # =========================================================
    # SLIDE 10: ROLE ADMIN - MANAJEMEN JENIS SAMPAH & TARIF
    # =========================================================
    s10 = create_base_slide(prs)
    add_header(s10, "ROLE ADMINISTRATOR", "Manajemen Jenis Sampah & Penentuan Tarif Beli", 
               "Pengelolaan katalog jenis sampah yang diterima bank sampah dan penetapan harga beli rupiah per kilogram.")

    # Screenshot Jenis Sampah
    img_ajs = os.path.join(SCREENSHOTS_DIR, "08_admin_jenis_sampah.png")
    add_screenshot_frame(s10, img_ajs, Inches(0.8), Inches(1.6), Inches(7.5), Inches(5.3))

    # Right side: Pricing Rules & Operations
    card_js = add_card(s10, Inches(8.5), Inches(1.6), Inches(4.0), Inches(5.3))
    tb_js = s10.shapes.add_textbox(Inches(8.7), Inches(1.75), Inches(3.6), Inches(5.0))
    tf_js = tb_js.text_frame
    tf_js.word_wrap = True

    p = tf_js.paragraphs[0]
    p.text = "🏷️ PENGELOLAAN KATALOG & TARIF"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY
    p.space_after = Pt(8)

    js_ops = [
        ("Tambah Jenis Baru", "Klik '+ Tambah Jenis Sampah' untuk mendaftarkan komoditas baru beserta harga beli per kg."),
        ("Edit Tarif Mengikuti Pasar", "Harga beli daur ulang dapat diperbarui sewaktu-waktu sesuai fluktuasi pasar lapak / pengepul."),
        ("Hapus / Nonaktifkan", "Kategori yang sudah tidak lagi diterima dapat dihapus melalui tombol aksi hapus di tabel."),
        ("Otomasi Perhitungan", "Tarif per kg ini otomatis menjadi pengali rumus pada formulir input setor nasabah.")
    ]
    for sj, dj in js_ops:
        p_sj = tf_js.add_paragraph()
        p_sj.text = f"➔ {sj}"
        p_sj.font.size = Pt(10.2)
        p_sj.font.bold = True
        p_sj.font.color.rgb = NAVY_DARK
        
        p_dj = tf_js.add_paragraph()
        p_dj.text = dj
        p_dj.font.size = Pt(9.2)
        p_dj.font.color.rgb = TEXT_MUTED
        p_dj.space_after = Pt(4)

    # Ketentuan Tarif
    p_kt = tf_js.add_paragraph()
    p_kt.text = "📌 KETENTUAN TARIF: Perubahan harga hanya berlaku untuk laporan baru yang dibuat setelah tanggal update. Laporan yang sudah masuk tetap menggunakan tarif saat laporan dibuat."
    p_kt.font.size = Pt(9.5)
    p_kt.font.bold = True
    p_kt.font.color.rgb = BLUE_BRAND
    p_kt.space_before = Pt(6)

    # =========================================================
    # SLIDE 11: ROLE ADMIN - MANAJEMEN WILAYAH & DATA PENGGUNA
    # =========================================================
    s11 = create_base_slide(prs)
    add_header(s11, "ROLE ADMINISTRATOR", "Manajemen Wilayah Operasional & Data Nasabah", 
               "Pemetaan area jangkauan penjemputan sampah dan pengawasan akun nasabah serta saldo tersimpan.")

    # Two Small Screenshots: Wilayah & Users
    img_wil = os.path.join(SCREENSHOTS_DIR, "09_admin_wilayah.png")
    img_usr = os.path.join(SCREENSHOTS_DIR, "10_admin_users.png")
    add_screenshot_frame(s11, img_wil, Inches(0.8), Inches(1.6), Inches(5.6), Inches(2.6))
    add_screenshot_frame(s11, img_usr, Inches(0.8), Inches(4.35), Inches(5.6), Inches(2.55))

    # Right side: Detail Wilayah & User Management
    card_wu = add_card(s11, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.3))
    tb_wu = s11.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.1), Inches(4.9))
    tf_wu = tb_wu.text_frame
    tf_wu.word_wrap = True

    p = tf_wu.paragraphs[0]
    p.text = "📍 WILAYAH & DATA NASABAH"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY
    p.space_after = Pt(8)

    wu_points = [
        ("Manajemen Wilayah Operasional", 
         "Admin mendaftarkan nama kelurahan/kecamatan dan kota administrasi (misal: Gambir - Jakarta Pusat). Wilayah ini langsung muncul di form setor nasabah."),
        ("Ketentuan Wilayah", 
         "Pastikan penamaan wilayah jelas untuk mencegah kesalahan penjemputan armada logistik fisik bank sampah."),
        ("Pengawasan Nasabah Terdaftar", 
         "Tabel user menampilkan nama, email, nomor HP, peran (USER/ADMIN), dan total akumulasi saldo dompet."),
        ("Ketentuan Data Nasabah", 
         "Nomor HP nasabah merupakan jalur komunikasi utama bagi admin untuk koordinasi penjemputan sampah besar.")
    ]
    for tw, dw in wu_points:
        p_tw = tf_wu.add_paragraph()
        p_tw.text = f"•  {tw}"
        p_tw.font.size = Pt(10.5)
        p_tw.font.bold = True
        p_tw.font.color.rgb = NAVY_DARK
        
        p_dw = tf_wu.add_paragraph()
        p_dw.text = dw
        p_dw.font.size = Pt(9.5)
        p_dw.font.color.rgb = TEXT_MUTED
        p_dw.space_after = Pt(6)

    # =========================================================
    # SLIDE 12: RINGKASAN MATRIX HAK AKSES & KETENTUAN 2 ROLE
    # =========================================================
    s12 = create_base_slide(prs)
    add_header(s12, "MATRIKS KOMPARASI", "Perbandingan Hak Akses & Tanggung Jawab 2 Role", 
               "Tabel ringkas pembagian wewenang, alur tanggung jawab, dan batasan sistem antara Nasabah dan Admin.")

    # Two Main Comparison Columns
    col1 = add_card(s12, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.3))
    tb_c1 = s12.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(5.2), Inches(4.9))
    tf_c1 = tb_c1.text_frame
    tf_c1.word_wrap = True

    p = tf_c1.paragraphs[0]
    p.text = "👤 NASABAH (ROLE: USER)"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = BLUE_BRAND
    p.space_after = Pt(10)

    user_matrix = [
        ("Akses Dashboard", "Melihat saldo pribadi dan status setoran sendiri."),
        ("Penyetoran Sampah", "Membuat laporan baru (input jenis, berat, wilayah, foto)."),
        ("Perubahan Status", "TIDAK BISA mengubah status (Hanya bisa memantau)."),
        ("Pengaturan Tarif", "TIDAK BISA mengubah tarif beli sampah."),
        ("Pengelolaan Wilayah", "Hanya bisa memilih dari wilayah yang tersedia."),
        ("Hak Saldo", "Menerima saldo otomatis saat laporan disetujui (Approved)."),
        ("Kewajiban", "Memilah sampah, menimbang presisi, mengunggah foto asli.")
    ]
    for ku, vu in user_matrix:
        pu = tf_c1.add_paragraph()
        pu.text = f"✔  {ku}: {vu}"
        pu.font.size = Pt(10)
        pu.font.color.rgb = NAVY_DARK
        pu.space_after = Pt(4)

    col2 = add_card(s12, Inches(6.9), Inches(1.6), Inches(5.6), Inches(5.3))
    tb_c2 = s12.shapes.add_textbox(Inches(7.1), Inches(1.8), Inches(5.2), Inches(4.9))
    tf_c2 = tb_c2.text_frame
    tf_c2.word_wrap = True

    p = tf_c2.paragraphs[0]
    p.text = "🛡️ PENGELOLA (ROLE: ADMIN)"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = NAVY_PRIMARY
    p.space_after = Pt(10)

    admin_matrix = [
        ("Akses Dashboard", "Melihat metrik global, total pengguna, dan volume sampah."),
        ("Penyetoran Sampah", "Memeriksa dan memvalidasi setoran dari seluruh nasabah."),
        ("Perubahan Status", "MEMILIKI HAK MUTLAK menyetujui (Approved) / menolak (Rejected)."),
        ("Pengaturan Tarif", "BEBAS menambah, mengedit harga beli, dan menghapus jenis sampah."),
        ("Pengelolaan Wilayah", "Bebas menambah dan mengatur cakupan wilayah operasional."),
        ("Hak Saldo", "Mengontrol mutasi penambahan saldo melalui sistem approval."),
        ("Kewajiban", "Memeriksa keaslian fisik sampah dalam batas waktu SLA 1x24 jam.")
    ]
    for ka, va in admin_matrix:
        pa = tf_c2.add_paragraph()
        pa.text = f"★  {ka}: {va}"
        pa.font.size = Pt(10)
        pa.font.color.rgb = NAVY_DARK
        pa.space_after = Pt(4)

    # =========================================================
    # SLIDE 13: SOP PEMECAHAN MASALAH & KETENTUAN UMUM
    # =========================================================
    s13 = create_base_slide(prs)
    add_header(s13, "TROUBLESHOOTING & FAQ", "Panduan Pemecahan Masalah & Pertanyaan Umum", 
               "Solusi cepat saat mengalami kendala teknis pada penggunaan aplikasi Sampahahayuk.")

    faq_cards = [
        ("1. Mengapa Saldo Saya Belum Bertambah Setelah Setor?", 
         "Saldo baru bertambah ketika admin telah menyetujui laporan Anda. Cek halaman 'Riwayat': jika status masih 'PENDING', mohon menunggu proses verifikasi fisik oleh petugas."),
        
        ("2. Apa yang Harus Dilakukan Jika Laporan Ditolak (REJECTED)?", 
         "Laporan ditolak umumnya karena foto tidak jelas, jenis sampah tidak sesuai, atau kondisi basah/kotor. Perbaiki pemilahan sampah Anda dan ajukan laporan setor baru."),
        
        ("3. Foto Gagal Terunggah Saat Submit Laporan", 
         "Pastikan ukuran file foto tidak melebihi 5 MB dan berformat .JPG, .PNG, atau .WEBP. Pastikan juga koneksi internet Anda stabil saat melakukan submit."),

        ("4. Bagaimana Cara Menjadi Admin Baru?", 
         "Pendaftaran via halaman web hanya menghasilkan akun Pengguna (Nasabah). Akun Admin diberikan oleh Super Admin melalui konfigurasi database internal.")
    ]

    for i, (q, a) in enumerate(faq_cards):
        row = i // 2
        col = i % 2
        x = Inches(0.8 + col * 5.9)
        y = Inches(1.6 + row * 2.65)
        
        c = add_card(s13, x, y, Inches(5.6), Inches(2.4))
        tb = s13.shapes.add_textbox(x + Inches(0.2), y + Inches(0.18), Inches(5.2), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p0 = tf.paragraphs[0]
        p0.text = q
        p0.font.size = Pt(11)
        p0.font.bold = True
        p0.font.color.rgb = BLUE_BRAND
        p0.space_after = Pt(6)

        p1 = tf.add_paragraph()
        p1.text = a
        p1.font.size = Pt(9.8)
        p1.font.color.rgb = NAVY_DARK

    # =========================================================
    # SLIDE 14: PENUTUP & INFORMASI KONFIGURASI SISTEM
    # =========================================================
    s14 = create_base_slide(prs, bg_color=NAVY_DARK)

    # Decorative Shape
    dec = s14.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(0), Inches(0), Inches(5.0), Inches(7.5))
    dec.fill.solid()
    dec.fill.fore_color.rgb = NAVY_PRIMARY
    dec.line.fill.background()

    tb_end = s14.shapes.add_textbox(Inches(1.5), Inches(1.5), Inches(10.3), Inches(4.5))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True

    p0 = tf_end.paragraphs[0]
    p0.text = "TERIMA KASIH"
    p0.font.size = Pt(36)
    p0.font.bold = True
    p0.font.color.rgb = RGBColor(255, 255, 255)
    p0.space_after = Pt(10)

    p1 = tf_end.add_paragraph()
    p1.text = "Sistem Bank Sampah Digital • Sampahahayuk Web Application"
    p1.font.size = Pt(16)
    p1.font.bold = True
    p1.font.color.rgb = BLUE_LIGHT
    p1.space_after = Pt(24)

    p2 = tf_end.add_paragraph()
    p2.text = "User Manual ini dirancang untuk memudahkan adopsi sistem oleh warga nasabah dan meningkatkan efisiensi operasional pengelola bank sampah."
    p2.font.size = Pt(12)
    p2.font.color.rgb = RGBColor(203, 213, 225)
    p2.space_after = Pt(28)

    # Info Grid
    p3 = tf_end.add_paragraph()
    p3.text = "KONTAK & INFORMASI TEKNIS SISTEM:"
    p3.font.size = Pt(12)
    p3.font.bold = True
    p3.font.color.rgb = RGBColor(255, 255, 255)
    p3.space_after = Pt(8)

    p4 = tf_end.add_paragraph()
    p4.text = "• Administrator Email: admin@sampahahayuk.com\n• Framework: Next.js 16 (App Router) + React 19\n• Database: PostgreSQL (Neon Cloud) via Prisma ORM\n• Otentikasi: JSON Web Token (JWT) Secured Session"
    p4.font.size = Pt(11)
    p4.font.color.rgb = RGBColor(147, 197, 253)

    prs.save(OUTPUT_PPTX)
    print(f"Presentation saved successfully to: {OUTPUT_PPTX}")

if __name__ == "__main__":
    generate_presentation()
