from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

DARK_BLUE = RGBColor(0x1F, 0x4E, 0x79)
BLUE = RGBColor(0x2E, 0x75, 0xB6)
LIGHT_BLUE = RGBColor(0xD6, 0xE9, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GRAY = RGBColor(0x33, 0x3F, 0x48)
GRAY = RGBColor(0x5B, 0x67, 0x70)

def add_rect(slide, x, y, w, h, color, line=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line:
        shape.line.color.rgb = color
    else:
        shape.line.fill.background()
    shape.shadow.inherit = False
    return shape

def add_text(slide, x, y, w, h, text, size=18, color=DARK_GRAY, bold=False, align=PP_ALIGN.LEFT, font_name="Calibri"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font_name
    return tb

def add_bullet(slide, x, y, w, h, bullets, size=16, color=DARK_GRAY, first_size=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for item in bullets:
        if isinstance(item, tuple):
            lead, rest = item
            if first:
                p = tf.paragraphs[0]
                first = False
            else:
                p = tf.add_paragraph()
            p.space_after = Pt(10)
            r = p.add_run()
            r.text = "▪  " + lead + " "
            r.font.bold = True
            r.font.size = Pt(first_size or size)
            r.font.color.rgb = DARK_BLUE
            r2 = p.add_run()
            r2.text = rest
            r2.font.size = Pt(first_size or size)
            r2.font.color.rgb = color
        else:
            if first:
                p = tf.paragraphs[0]
                first = False
            else:
                p = tf.add_paragraph()
            p.space_after = Pt(8)
            r = p.add_run()
            r.text = "▪  " + item
            r.font.size = Pt(size)
            r.font.color.rgb = color
    return tb

def title_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, prs.slide_width, prs.slide_height, DARK_BLUE)
    # accent bar
    add_rect(slide, 0, Inches(6.9), prs.slide_width, Inches(0.6), BLUE)
    add_text(slide, Inches(1), Inches(2.2), Inches(11.33), Inches(1.5),
             "Community Connect — News & Announcements Portal", size=40, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(1), Inches(3.7), Inches(11.33), Inches(1),
             "WS101 Mini Capstone Project", size=24, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(1), Inches(4.4), Inches(11.33), Inches(1),
             "A Frontend-Only Barangay Community Portal  |  HTML • CSS • Vanilla JavaScript", size=18, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(1), Inches(6.0), Inches(11.33), Inches(0.6),
             "Presented by: [Your Name]", size=16, color=WHITE, align=PP_ALIGN.CENTER)
    return slide

def content_slide(title_text, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.9), BLUE)
    add_text(slide, Inches(0.5), Inches(0.12), Inches(12.33), Inches(0.7),
             title_text, size=28, color=WHITE, bold=True)
    if subtitle:
        add_text(slide, Inches(0.5), Inches(0.95), Inches(12.33), Inches(0.4),
                 subtitle, size=14, color=GRAY)
    return slide

# ---- SLIDE 1: Title ----
title_slide()

# ---- SLIDE 2: Project Overview ----
slide = content_slide("Project Overview", "What is Community Connect?")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(5.2), [
    ("What it is:", "A complete frontend-only community portal for barangay news, announcements, events, activities, galleries, and emergency services."),
    ("Technology:", "Plain HTML, CSS, and vanilla JavaScript — no frameworks, no build step, no backend."),
    ("Data storage:", "All data lives in browser localStorage and IndexedDB (for videos)."),
    ("Offline ready:", "Service worker + PWA manifest for installable app experience."),
    ("User roles:", "Admin, Barangay Official, Resident (registered user), and Guest visitor."),
], size=17)

# ---- SLIDE 3: User Roles ----
slide = content_slide("User Roles & Access", "Role-based dashboard system")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(5.2), [
    ("Admin (admin / admin123):", "Full control — content, users, comments, settings, Go Live streaming, emergency broadcasts."),
    ("Barangay Official (official / official123):", "Review certificate requests, respond to incidents, view community feeds scoped to their barangay."),
    ("Resident:", "Register and log in to save posts, RSVP to events, request certificates, report incidents."),
    ("Guest:", "Read public pages, search, view gallery, comment with a display name — no account needed."),
], size=17)

# ---- SLIDE 4: Public Features ----
slide = content_slide("Public Website Features", "No login required")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(6.1), Inches(5.2), [
    "Home — Hero banner, weather card, latest content, quick links.",
    "News — Facebook-style feed with search, category & date filters.",
    "Announcements — Status badges (Important, Upcoming, Active, Emergency).",
    "Events — List + Calendar view with RSVP (\"I'm Going\").",
    "Activities — Community programs listing.",
    "Gallery — Photo lightbox.",
    "Search — Global keyword search with highlighting.",
    "Hotlines — Emergency directory with one-tap call.",
], size=15)
add_bullet(slide, Inches(6.7), Inches(1.6), Inches(6.1), Inches(5.2), [
    "News AI Assistant — Chatbot to ask questions about posted content.",
    "About — Project showcase, team, mission, vision.",
    "Contact — Contact form.",
    "Mobile — Fixed bottom tab bar (Home · News · Alerts · Events · Menu).",
    "i18n — English / Filipino toggle remembered per browser.",
    "Dark mode — Default dark theme with light toggle.",
    "SEO — Meta tags, Open Graph, canonical URLs, JSON-LD.",
], size=15)

# ---- SLIDE 5: Resident Dashboard ----
slide = content_slide("Resident Dashboard", "user-dashboard.html — after registration & login")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(6.1), Inches(5.2), [
    "Overview — Quick actions grid, saved items, notifications.",
    "Certificate Requests — Indigency, Clearance, Residency, Good Moral, Business Clearance.",
    "Status tracking — Pending → Ready for Pickup → Released.",
    "Incident Reports — Report accidents, fires, floods, crime, medical, utility emergencies.",
    "Status — Reported → Responding → Resolved.",
], size=15)
add_bullet(slide, Inches(6.7), Inches(1.6), Inches(6.1), Inches(5.2), [
    "Community Feed — Lost Item / Found Item / Community Share with photos.",
    "Like, Comment, Share, Save posts.",
    "My Profile — Update name, photo, phone, address, password, barangay.",
    "Emergency Alert Popup — 30-second alarm + AI voice reads instructions.",
    "Live Broadcast Player — View admin's Go Live stream.",
], size=15)

# ---- SLIDE 6: Official Dashboard ----
slide = content_slide("Barangay Official Dashboard", "official-dashboard.html — barangay-scoped duties")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(6.1), Inches(5.2), [
    "On Duty toggle — Availability status in header.",
    "Stats — Pending, Ready to Print, Released, Total requests, Active incidents.",
    "Certificate Requests — Only own barangay's requests.",
    "Review, mark ready, reject with remarks, print A4 certificate.",
    "Bulk mark-ready, export release-log CSV.",
], size=15)
add_bullet(slide, Inches(6.7), Inches(1.6), Inches(6.1), Inches(5.2), [
    "Incident Reports — All incidents (not barangay-limited).",
    "Priority queue, claim, timestamped updates, mark Responding/Resolved.",
    "High-severity alert sound, incidents-by-type chart, CSV export.",
    "Community Feeds — Same Facebook-style feeds as residents.",
    "Notification Bell — Filter tabs: All / Incidents / Certificates.",
], size=15)

# ---- SLIDE 7: Admin Dashboard ----
slide = content_slide("Admin Dashboard", "admin.html — full system administration")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(6.1), Inches(5.2), [
    "Dashboard — Content stats, charts, audit log, Quick Actions.",
    "Manage Content — Add/edit/delete News, Announcements, Events, Activities, Gallery.",
    "Multi-photo (up to 6 collage) + multi-video (up to 3).",
    "Pin/feature news to top of public list.",
    "Manage Users — Create/manage Barangay Officials, view residents.",
], size=15)
add_bullet(slide, Inches(6.7), Inches(1.6), Inches(6.1), Inches(5.2), [
    "Comments — Moderate threads; hide/delete (removed from public site).",
    "Community Posts — Moderate resident Lost & Found posts.",
    "Go Live Studio — Stream camera/mic to viewers in real time.",
    "Emergency Broadcast Studio — Publish alerts to one barangay or all.",
    "Settings — General, Notifications, Security, Data & Backup, Users.",
], size=15)

# ---- SLIDE 8: Go Live ----
slide = content_slide("Go Live — Real-Time Streaming", "Admin broadcasts camera/mic to residents")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(5.2), [
    ("How it works:", "Admin clicks GO LIVE → enables camera/mic → sets title/category → Start Broadcast."),
    ("Transport:", "MediaRecorder captures 1-second chunks; viewers use MediaSource for near-live playback."),
    ("Dual mode:", "WebSocket relay (server.js) for cross-device LAN streaming + BroadcastChannel fallback for same-browser."),
    ("Viewer experience:", "LIVE badge, live viewer count, running duration, stall watchdog overlay."),
    ("Recording:", "Optional recordings saved to IndexedDB under Live Archives."),
    ("Requirement:", "node server.js (zero-dependency Node server using built-in modules only)."),
], size=16)

# ---- SLIDE 9: Emergency Broadcast ----
slide = content_slide("Emergency Broadcast System", "Instant alerts to residents & officials")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(5.2), [
    ("Admin sets:", "Title, type, affected barangay (one or all), location, description, safety instructions, expiry."),
    ("Delivery:", "Instant EMERGENCY ALERT popup on matching dashboards."),
    ("Audio:", "30-second Web Audio alarm + AI voice (SpeechSynthesis) reads the description."),
    ("Controls:", "STOP ALARM (silences own device only); REPEAT ANNOUNCEMENT."),
    ("Persistence:", "Popup stays until admin resolves it or it expires."),
], size=16)

# ---- SLIDE 10: Certificate & Incident Services ----
slide = content_slide("Certificate & Incident Services", "Barangay services for residents and officials")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(6.1), Inches(5.2), [
    "Certificate Types: Indigency, Clearance, Residency, Good Moral, Business Clearance.",
    "Resident selects barangay (only admin-created barangays appear).",
    "Reference number (e.g., CC-2026-12345) generated per request.",
    "Officials print A4 certificate; marks Released and records who printed.",
    "Bulk mark-ready and export release-log CSV.",
], size=15)
add_bullet(slide, Inches(6.7), Inches(1.6), Inches(6.1), Inches(5.2), [
    "Incident Types: Vehicular accident, fire, flood/typhoon, crime, medical, utility, other.",
    "Severity: High / Medium / Low with photo + GPS/landmark support.",
    "Reference number (e.g., INC-2026-12345).",
    "Notifies ALL officials (emergencies are not barangay-limited).",
    "Officials claim, add updates, mark Responding/Resolved; CSV export.",
], size=15)

# ---- SLIDE 11: Notifications ----
slide = content_slide("Notifications System", "Real-time alerts across all dashboards")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(5.2), [
    ("Header bell:", "Unread count badge on admin, resident, and official dashboards."),
    ("Activity:", "Records real user activity — likes, comments, shares, saves, RSVPs, sign-ups, Go Live broadcasts."),
    ("Timestamps:", "Relative times (\"2m\", \"4h\", \"3d\"); clickable rows open related content."),
    ("Sync:", "Real-time across tabs via BroadcastChannel with storage fallback."),
    ("Preferences:", "Full Notifications section: All/Unread filters, mark one/all as read, clear history, per-type toggles."),
    ("Honest:", "Fresh install starts empty — nothing is faked."),
], size=16)

# ---- SLIDE 12: Technical Architecture ----
slide = content_slide("Technical Architecture", "Frontend stack with optional Node relay server")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(6.1), Inches(5.2), [
    "Files: index.html, news.html, admin.html, user-dashboard.html, official-dashboard.html, server.js, sw.js, manifest.json.",
    "Styles: css/style.css (~9,000 lines, numbered sections).",
    "Logic: js/script.js (~14,000 lines, vanilla JavaScript).",
    "Storage: localStorage (text) + IndexedDB (videos, no size limit).",
    "i18n: Built-in English ↔ Filipino translation engine (t()).",
], size=15)
add_bullet(slide, Inches(6.7), Inches(1.6), Inches(6.1), Inches(5.2), [
    "server.js: Zero-dependency Node.js static server + WebSocket relay at /live.",
    "PWA: Installable manifest + service worker precaching (cache v35) + offline.html fallback.",
    "SEO: Meta tags, canonical URLs, Open Graph/Twitter cards, JSON-LD.",
    "Accessibility: Skip link, ARIA labels, focus trap, Escape-to-close, reduced-motion support.",
    "Toasts: Top-right stacked notifications instead of blocking alert() dialogs.",
], size=15)

# ---- SLIDE 13: Testing ----
slide = content_slide("Testing & Quality Assurance", "Verified with automated test suite")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(5.2), [
    ("E2E Suite:", "tests/e2e.html — 706 assertions run in hidden iframes via Chrome DevTools Protocol."),
    ("Result:", "706/706 PASS — all features verified end-to-end."),
    ("Relay Test:", "tests/relay-test.cjs — static serving, path-traversal guard, WebSocket text + binary relay (70KB frame): ALL PASS."),
    ("i18n Check:", "tests/check-i18n-missing.cjs — 0 missing translations."),
    ("Not covered:", "CSV/JSON downloads (headless Chrome blocks), service-worker flow (needs HTTP), real camera capture (needs hardware)."),
], size=16)

# ---- SLIDE 14: How to Run ----
slide = content_slide("How to Run", "Installation & demo access")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(5.2), [
    ("Offline:", "Double-click index.html — works from file://."),
    ("PWA / HTTP:", "python -m http.server 8000 → open http://localhost:8000/."),
    ("Go Live:", "node server.js → Admin: http://localhost:8787/admin.html → Viewers: http://<LAN-IP>:8787/user-dashboard.html."),
    ("Demo Accounts:", "Admin (admin/admin123), Official (official/official123), Resident (register at register.html)."),
], size=16)

# ---- SLIDE 15: Conclusion ----
slide = content_slide("Conclusion", "Why Community Connect?")
add_bullet(slide, Inches(0.5), Inches(1.6), Inches(12.33), Inches(5.2), [
    ("Complete portal:", "News, announcements, events, activities, gallery — all publicly accessible."),
    ("Citizen services:", "Certificate requests + incident reporting with full tracking workflows."),
    ("Real-time:", "Go Live streaming + emergency broadcasts with alarm + AI voice."),
    ("Role-based:", "Dashboards for Admin, Barangay Officials, and Residents."),
    ("Professional UX:", "Glass-morphism UI, dark mode, mobile-first, full accessibility, PWA support."),
    ("Zero dependencies:", "Fully offline-capable, 706/706 tests passing, no backend or build step required."),
], size=16)

# ---- SLIDE 16: Thank You ----
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_rect(slide, 0, 0, prs.slide_width, prs.slide_height, DARK_BLUE)
add_rect(slide, 0, Inches(6.9), prs.slide_width, Inches(0.6), BLUE)
add_text(slide, Inches(1), Inches(2.6), Inches(11.33), Inches(1.2),
         "Thank You — Q&A", size=40, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
add_text(slide, Inches(1), Inches(3.9), Inches(11.33), Inches(0.8),
         "Community Connect — News & Announcements Portal", size=20, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)
add_text(slide, Inches(1), Inches(4.5), Inches(11.33), Inches(0.8),
         "WS101 Mini Capstone Project  |  HTML • CSS • Vanilla JavaScript", size=16, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

prs.save("Community_Connect_Report.pptx")
print("Saved Community_Connect_Report.pptx with", len(prs.slides._sldIdLst), "slides")
