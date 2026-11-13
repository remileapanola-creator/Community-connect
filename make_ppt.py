from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

DARK_BLUE = RGBColor(0x1F, 0x4E, 0x79)
BLUE = RGBColor(0x2E, 0x75, 0xB6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BLUE = RGBColor(0xD6, 0xE9, 0xF8)
DARK_GRAY = RGBColor(0x33, 0x3F, 0x48)

def add_title_slide(title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = DARK_BLUE; bg.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.33), Inches(2))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = title; r.font.size = Pt(44); r.font.bold = True; r.font.color.rgb = WHITE
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = subtitle; r2.font.size = Pt(20); r2.font.color.rgb = LIGHT_BLUE
    return slide

def add_content_slide(title, bullets):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bar = slide.shapes.add_shape(1, 0, 0, prs.slide_width, Inches(0.9))
    bar.fill.solid(); bar.fill.fore_color.rgb = BLUE; bar.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(12.33), Inches(0.7))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; r = p.add_run(); r.text = title; r.font.size = Pt(28); r.font.bold = True; r.font.color.rgb = WHITE
    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(12.33), Inches(5.8))
    btf = body.text_frame; btf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = btf.paragraphs[0]; first = False
        else:
            p = btf.add_paragraph()
        p.space_after = Pt(10)
        if isinstance(b, tuple):
            r = p.add_run(); r.text = b[0] + ' '; r.font.bold = True; r.font.size = Pt(18); r.font.color.rgb = DARK_BLUE
            r2 = p.add_run(); r2.text = b[1]; r2.font.size = Pt(16); r2.font.color.rgb = DARK_GRAY
        else:
            r = p.add_run(); r.text = '-  ' + b; r.font.size = Pt(16); r.font.color.rgb = DARK_GRAY
    return slide

# Slide 1
add_title_slide('Community Connect — News & Announcements Portal',
                'WS101 Mini Capstone Project\nA Frontend-Only Barangay Community Portal\nHTML • CSS • Vanilla JavaScript')

# Slide 2
add_content_slide('Project Overview', [
    ('What it is:', 'A complete community portal for barangay news, announcements, events, activities, galleries, and emergency services.'),
    ('Technology:', 'Plain HTML, CSS, and vanilla JavaScript — no frameworks, no build step, no backend.'),
    ('Data storage:', 'All data lives in browser localStorage and IndexedDB (for videos).'),
    ('Offline ready:', 'Service worker + PWA manifest for installable app experience.'),
    ('User roles:', 'Admin, Barangay Official, Resident (registered user), and Guest visitor.'),
])

# Slide 3
add_content_slide('Target Users & Roles', [
    ('Admin (admin / admin123):', 'Full control — content, users, comments, settings, Go Live streaming, emergency broadcasts.'),
    ('Barangay Official (official / official123):', 'Review certificate requests, respond to incidents, view community feeds scoped to their barangay.'),
    ('Resident:', 'Register and log in to save posts, RSVP to events, request certificates, report incidents.'),
    ('Guest:', 'Read public pages, search, view gallery, comment with a display name — no account needed.'),
])

# Slide 4
add_content_slide('Public Website Features', [
    'Home — Hero banner, weather card, latest content, quick links.',
    'News — Facebook-style feed with search, category & date filters.',
    'Announcements — Important notices with status badges.',
    'Events — List + Calendar view with RSVP ("I am Going").',
    'Activities — Community programs and activities.',
    'Gallery — Photo lightbox.',
    'Search — Global keyword search across all content with highlighting.',
    'Hotlines — Emergency directory (barangay hall, police, health, fire, rescue) with one-tap call.',
    'News AI Assistant — Chatbot to ask questions about posted content.',
    'Mobile — Fixed bottom tab bar (Home · News · Alerts · Events · Menu).',
])

# Slide 5
add_content_slide('Resident Dashboard Features', [
    'Overview — Quick actions, saved items, notifications.',
    'Certificate Requests — File for Indigency, Clearance, Residency, Good Moral, Business Clearance. Status: Pending → Ready → Released.',
    'Incident Reports — Report accidents, fires, floods, crime, medical, utility emergencies. Status: Reported → Responding → Resolved.',
    'Community Feed — Post Lost Item / Found Item / Community Share with photos. Like, Comment, Share, Save.',
    'My Profile — Update name, photo, phone, address, password, and set barangay for emergency alerts.',
    'Emergency Alert Popup — 30-second alarm + AI voice reads instructions. STOP ALARM button.',
])

# Slide 6
add_content_slide('Barangay Official Dashboard Features', [
    'On Duty toggle — Shows availability status.',
    'Certificate Requests — Barangay-scoped: only see requests for your own barangay. Review, mark ready, reject with remarks, print A4 certificate, bulk actions, export CSV.',
    'Incident Reports — Claim reports, add timestamped updates, mark Responding/Resolved. Priority queue (escalated first). High-severity alert sound. Incidents-by-type chart.',
    'Community Feeds — Facebook-style news, announcements, activities + Upcoming Events RSVP + Gallery.',
    'Notification Bell — Filter tabs: All / Incidents / Certificates. Alert sound for high-severity.',
])

# Slide 7
add_content_slide('Admin Dashboard Features', [
    'Dashboard — Content stats, charts (content mix, posts per month, news views per day), Recent Admin Activity audit log.',
    'Manage Content — Add/edit/delete News, Announcements, Events, Activities, Gallery. Multi-photo (up to 6) + multi-video (up to 3). Pin/feature news.',
    'Manage Users — Create Barangay Official accounts (add/edit/deactivate/delete). View registered residents.',
    'Comments — Moderate Facebook-style comment threads. Hide/delete comments (hidden from public site).',
    'Community Posts — Moderate resident Lost & Found posts. Alert Community (banner + notification).',
    'Go Live Studio — Stream camera/mic to viewers in real time. Optional recordings saved to IndexedDB.',
    'Emergency Broadcast Studio — Publish alerts to one barangay or all barangays with expiry.',
    'Settings — General, Notifications, Security, Data & Backup, Users. Export/Import JSON, reset with auto-backup.',
])

# Slide 8
add_content_slide('Go Live — Real-Time Streaming', [
    'Admin clicks GO LIVE → opens studio → enables camera/mic → sets title/category → Start Broadcast.',
    'MediaRecorder captures 1-second chunks; relayed to viewers where MediaSource appends them for near-live playback.',
    'Dual transport: WebSocket relay (server.js) for cross-device LAN streaming + BroadcastChannel fallback for same-browser.',
    'Viewers see: LIVE badge, live viewer count, running duration, stall watchdog overlay.',
    'Optional recordings saved to IndexedDB under Live Archives.',
    'Requires: node server.js (zero-dependency Node server).',
])

# Slide 9
add_content_slide('Emergency Broadcast System', [
    'Admin publishes alert: Title, type, affected barangay (one or all), location, description, safety instructions, expiry.',
    'Instant delivery to matching residents and officials dashboards.',
    'EMERGENCY ALERT popup with 30-second Web Audio alarm.',
    'AI voice (SpeechSynthesis) reads the description aloud.',
    'STOP ALARM (silences own device only); REPEAT ANNOUNCEMENT.',
    'Stays visible until admin resolves it or it expires.',
])

# Slide 10
add_content_slide('Certificate Services Flow', [
    'Resident selects their barangay (only barangays with official accounts appear).',
    'Chooses certificate type + enters purpose/details.',
    'Reference number generated (e.g., CC-2026-12345).',
    'Track status: Pending → Ready for Pickup → Released.',
    'Notification sent to that barangay officials.',
    'Officials see only their own barangay requests; review, mark ready, reject with remarks, print A4 certificate.',
    'Bulk mark-ready; export release-log CSV.',
])

# Slide 11
add_content_slide('Incident Reporting Flow', [
    'Resident reports: Type (accident, fire, flood, crime, medical, utility, other), Severity, Barangay, Location, Description, Photo, GPS/landmark.',
    'Reference number (e.g., INC-2026-12345). Status: Reported → Responding → Resolved.',
    'Notifies ALL barangay officials (emergencies are not barangay-limited).',
    'Officials: Claim incident, assign to self, timestamped response timeline (resident-visible or internal).',
    'Escalation flag if still Reported after 15 minutes.',
    'Response-time KPI, open/escalated counters, incidents-by-type chart, CSV export.',
])

# Slide 12
add_content_slide('Notifications System', [
    'Header bell on all dashboards with unread count badge.',
    'Real user activity only: likes, comments, shares, saves, RSVPs, sign-ups, Go Live broadcasts.',
    'Relative timestamps ("2m", "4h", "3d"); clickable rows open related content.',
    'Real-time across tabs (BroadcastChannel with storage fallback).',
    'Full Notifications section: All/Unread filters, mark one/all as read, clear history.',
    'Per-type settings control what alerts the bell. Fresh install starts empty — nothing is faked.',
])

# Slide 13
add_content_slide('Data Storage & Safety', [
    'Storage keys (all community* prefix): communityNews, communityAnnouncements, communityEvents, communityActivities, communityGallery, communityComments, communityBookmarks, communityRsvps, communityUsers, barangayOfficials, certificateRequests, incidentReports, communityNotifications, communityEmergencyAlerts, communityUserPosts, communityLiveArchive (IndexedDB for videos).',
    'Schema versioning with automatic migrations on startup.',
    'Validated JSON imports (app name, schema version, per-key shape).',
    'Auto-backup snapshot on reset; one-click restore (RESTORE LAST RESET).',
    'Export JSON / Import JSON / Export News CSV tools.',
])

# Slide 14
add_content_slide('Design & UX Highlights', [
    'Premium glass UI: Mesh-gradient background, frosted headers/sidebars, gradient stat cards, glass cards/panels.',
    'Dark mode default with light-theme parity; theme remembered per browser.',
    'Mobile-first: Slide-in drawer navigation, bottom sheets for modals, stacked cards for tables, no horizontal overflow from 360px up.',
    'Accessibility: Skip link, aria-expanded/aria-current, focus trap in modals, Escape-to-close, :focus-visible styles, prefers-reduced-motion support.',
    'SEO: Unique meta descriptions, canonical URLs, Open Graph/Twitter cards, JSON-LD.',
    'Toasts: Top-right stacked notifications (green/red/blue) instead of blocking alert() dialogs.',
])

# Slide 15
add_content_slide('Technical Architecture', [
    'Files: index.html, news.html, announcements.html, events.html, activities.html, gallery.html, search.html, admin.html, user-dashboard.html, official-dashboard.html, server.js, sw.js, manifest.json, css/style.css, js/script.js.',
    'Key modules in js/script.js: I18N_MAP/t() (EN<->FIL translation), loadList/setData (localStorage), runMigrations (schema versioning), notifyAdmin (notifications), initLiveTransport/Studio/Player (Go Live), publishEmergency/resolveEmergency (alerts).',
    'server.js: Zero-dependency Node.js static server + WebSocket relay at /live for Go Live streaming.',
    'PWA: Installable manifest + service worker precaching (cache v35) + offline.html fallback.',
])

# Slide 16
add_content_slide('Testing & Quality Assurance', [
    'E2E suite: tests/e2e.html (706 assertions) run in hidden iframes.',
    'Runner: tests/e2e-run.cjs via Chrome DevTools Protocol; waits for real completion (IndexedDB, async work).',
    'Result: 706/706 PASS.',
    'Relay test: tests/relay-test.cjs — static serving, path-traversal guard, WebSocket text + binary relay (including 70KB frame): ALL PASS.',
    'i18n check: 0 missing translations.',
    'Not covered by suite: CSV/JSON downloads (headless Chrome blocks them), service-worker flow (requires HTTP), real camera capture (needs hardware).',
])

# Slide 17
add_content_slide('How to Run', [
    'Simplest: Double-click index.html (works from file://).',
    'With PWA/offline (HTTP): python -m http.server 8000 → open http://localhost:8000/.',
    'Go Live streaming: node server.js → Admin: http://localhost:8787/admin.html → Viewers: http://<LAN-IP>:8787/user-dashboard.html.',
    'Demo accounts: Admin (admin/admin123), Official (official/official123), Resident (register at register.html).',
])

# Slide 18
add_content_slide('Conclusion', [
    'Complete barangay portal: news, announcements, events, activities, gallery.',
    'Citizen services: certificate requests + incident reporting with full tracking workflow.',
    'Real-time engagement: Go Live streaming + emergency broadcasts with alarm + AI voice.',
    'Role-based dashboards for Admin, Officials, and Residents.',
    'Professional, accessible, mobile-first design with PWA support.',
    'Zero dependencies, fully offline-capable, 706/706 tests passing.',
])

# Slide 19
add_title_slide('Thank You — Q&A', 'Community Connect — News & Announcements Portal\nWS101 Mini Capstone Project\nHTML • CSS • Vanilla JavaScript')

prs.save('Community_Connect_Presentation.pptx')
print('PPTX saved successfully')
