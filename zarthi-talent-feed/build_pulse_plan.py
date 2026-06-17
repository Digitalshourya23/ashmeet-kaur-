# -*- coding: utf-8 -*-
"""Builds Zarthi_Pulse_Plan.xlsx — CEO-ready plan for the employee feed."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Zarthi_Pulse_Plan.xlsx")

# ---- palette ----
BRAND   = "6C5CE7"
BRAND_D = "4A3DB0"
ACCENT  = "00D4B4"
GOLD    = "FFCF5C"
INK     = "1A2236"
LIGHT   = "F4F2FF"
LIGHT2  = "EAF7F4"
GREY    = "6B7280"
WHITE   = "FFFFFF"

thin = Side(style="thin", color="DDDDE6")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def font(sz=11, b=False, color="222222", it=False):
    return Font(name="Calibri", size=sz, bold=b, color=color, italic=it)
def fill(c): return PatternFill("solid", fgColor=c)
def center(wrap=True): return Alignment(horizontal="center", vertical="center", wrap_text=wrap)
def left(wrap=True):   return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

wb = Workbook()

def title_block(ws, title, subtitle, ncols, color=BRAND):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(1, 1, title); c.font = font(18, True, WHITE); c.fill = fill(color)
    c.alignment = left(False); ws.row_dimensions[1].height = 34
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    s = ws.cell(2, 1, subtitle); s.font = font(11, False, WHITE, it=True); s.fill = fill(color)
    s.alignment = left(False); ws.row_dimensions[2].height = 22

def header_row(ws, row, headers, color=INK):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row, i, h); c.font = font(11, True, WHITE)
        c.fill = fill(color); c.alignment = center(); c.border = border
    ws.row_dimensions[row].height = 26

def data_rows(ws, start, rows, widths, alt=True, wrap_cols=None):
    wrap_cols = wrap_cols or []
    for r, rowdata in enumerate(rows):
        rr = start + r
        bg = LIGHT if (alt and r % 2 == 0) else WHITE
        for i, val in enumerate(rowdata, 1):
            c = ws.cell(rr, i, val)
            c.font = font(10.5)
            c.fill = fill(bg); c.border = border
            c.alignment = left() if (i in wrap_cols or i == len(rowdata)) else center()
        ws.row_dimensions[rr].height = 30
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

# ============================================================
# 1) EXECUTIVE SUMMARY
# ============================================================
ws = wb.active; ws.title = "1. Exec Summary"
ws.sheet_view.showGridLines = False
title_block(ws, "Zarthi Pulse — Employee Engagement Feed",
            "Internal social feed + Rewards & Recognition + smart relevance — built into the Zarthi app", 4)

rows = [
    ("What is it?", "An in-app feed (like Instagram / Facebook, but private to Zarthi) where every employee can post, react and get recognised."),
    ("Why build it?", "Boost employee engagement, make wins visible, run Rewards & Recognition (R&R), and build culture across teams — all in one place."),
    ("The big idea", "Everyone posts. The system quietly shows each person the MOST RELEVANT posts for their team and skill (SEO sees SEO, Sales sees Sales)."),
    ("Post formats", "Text, Photo, Carousel, Video and GIF — with Instagram-style font, text colour, background, music, location and #tags."),
    ("R&R engine", "A live Top-5 leaderboard card inside the feed. Give kudos, earn points & badges, crown monthly winners."),
    ("Newsletter", "A subscribe card in the feed delivers the monthly company newsletter to inboxes."),
    ("Admin power", "Search any employee by name/email and see their posts, comments, kudos and activity + the keyword tags driving their feed."),
    ("Who is it for?", "Every Zarthi employee — Marketing (SEO/Paid/SMM), Sales, Technology, HR and more."),
    ("Time to launch", "MVP in ~6-8 weeks. Full feature set across 3 phases in ~4-5 months."),
]
header_row(ws, 4, ["Topic", "In plain language"])
ws.merge_cells("B4:D4")
for r, (k, v) in enumerate(rows):
    rr = 5 + r
    bg = LIGHT if r % 2 == 0 else WHITE
    a = ws.cell(rr, 1, k); a.font = font(11, True, BRAND_D); a.fill = fill(bg); a.alignment = left(); a.border = border
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
    b = ws.cell(rr, 2, v); b.font = font(10.5); b.fill = fill(bg); b.alignment = left(); b.border = border
    for cc in range(3, 5):
        ws.cell(rr, cc).fill = fill(bg); ws.cell(rr, cc).border = border
    ws.row_dimensions[rr].height = 34
ws.column_dimensions["A"].width = 18
for col in "BCD": ws.column_dimensions[col].width = 34

# six pillars strip
prow = 5 + len(rows) + 1
ws.cell(prow, 1, "The 6 building blocks").font = font(13, True, BRAND_D)
ws.row_dimensions[prow].height = 24
pillars = [("📰 The Feed", BRAND), ("🎨 Rich Composer", BRAND), ("🏆 R&R Leaderboard", GOLD),
           ("✉️ Newsletter", ACCENT), ("🎯 Relevance Engine", BRAND_D), ("🛠️ Admin Insights", INK)]
prow += 1
for i, (p, col) in enumerate(pillars):
    c = ws.cell(prow + (i // 3), (i % 3) * 1 + 1 + (i % 3), p)  # not used; simpler below
# simpler: place pillars in a 2x3 grid across A..C
prow2 = prow
for i, (p, col) in enumerate(pillars):
    r = prow2 + (i // 3)
    cidx = (i % 3) + 1
    cell = ws.cell(r, cidx, p)
    cell.font = font(11, True, WHITE); cell.fill = fill(col); cell.alignment = center(); cell.border = border
    ws.row_dimensions[r].height = 30

# ============================================================
# 2) FEATURE LIST
# ============================================================
ws = wb.create_sheet("2. Feature List")
ws.sheet_view.showGridLines = False
title_block(ws, "Feature List — what we are building", "Grouped by module, with priority and the phase it lands in", 5)
header_row(ws, 4, ["Module", "Feature", "What it does (plain language)", "Priority", "Phase"])
feat = [
    ("The Feed", "Scrollable feed", "Central home screen showing posts from colleagues, newest + most relevant first.", "Must", "P1"),
    ("The Feed", "Post formats", "Text, photo, carousel (multi-image), video and GIF posts.", "Must", "P1"),
    ("The Feed", "Reactions", "Like / love and other quick reactions on any post.", "Must", "P1"),
    ("The Feed", "Comments", "Threaded comments so people can reply and discuss.", "Must", "P1"),
    ("The Feed", "Share & Save", "Re-share a post internally or bookmark it for later.", "Should", "P2"),
    ("Composer", "Text styling", "Instagram-style: pick font, text colour and background colour.", "Should", "P2"),
    ("Composer", "GIF picker", "Search and drop a GIF into a post.", "Should", "P2"),
    ("Composer", "Music / audio", "Attach a music clip or voice note to a post.", "Could", "P2"),
    ("Composer", "Location tag", "Tag a location (e.g. Zarthi HQ, event).", "Could", "P2"),
    ("Composer", "Mentions & #tags", "@mention colleagues and add #hashtags for discovery.", "Should", "P1"),
    ("R&R", "Top-5 Leaderboard", "Live card in the feed showing the top 5 employees this month.", "Must", "P1"),
    ("R&R", "Kudos & points", "Give kudos to teammates; points accumulate towards rank.", "Must", "P2"),
    ("R&R", "Badges", "Earn badges (Top Contributor, Rising Star, etc.).", "Should", "P2"),
    ("R&R", "Monthly winners", "Auto-crown winners and celebrate in the feed.", "Should", "P2"),
    ("Newsletter", "Subscribe card", "In-feed card to subscribe / unsubscribe to the newsletter.", "Should", "P2"),
    ("Newsletter", "Delivery", "Monthly newsletter sent to subscribed inboxes.", "Should", "P2"),
    ("Relevance", "Profile tagging", "Each user tagged by department + specialization (e.g. Marketing > SEO).", "Must", "P1"),
    ("Relevance", "Post auto-tagging", "Posts auto-tagged by author dept, keywords and #tags.", "Must", "P1"),
    ("Relevance", "Feed ranking", "Show each user their team/skill-relevant posts first.", "Must", "P1/P3"),
    ("Relevance", "Smart ranking (ML-lite)", "Learn from engagement to improve relevance over time.", "Could", "P3"),
    ("Admin", "User search", "Find any employee by name or email.", "Must", "P1"),
    ("Admin", "Activity profile", "See a user's posts, comments, kudos and recent activity.", "Must", "P1"),
    ("Admin", "Keyword map", "View the department keywords/tags driving a user's feed.", "Should", "P1"),
    ("Admin", "Analytics dashboard", "Engagement, active users, top posts, adoption by team.", "Should", "P3"),
    ("Platform", "Notifications", "Push + in-app notifications for kudos, mentions, comments.", "Should", "P2"),
    ("Platform", "SSO / login", "Single sign-on (Google / company login).", "Must", "P1"),
    ("Platform", "Moderation", "Report/flag a post; admin review queue.", "Should", "P2"),
]
data_rows(ws, 5, feat, [13, 18, 52, 9, 7], wrap_cols=[3])

# ============================================================
# 3) POST FORMATS & COMPOSER
# ============================================================
ws = wb.create_sheet("3. Formats & Composer")
ws.sheet_view.showGridLines = False
title_block(ws, "Post Formats & the Creator Tools", "Everything a user can make and how they can style it", 3, color=BRAND_D)
header_row(ws, 4, ["Post format", "Best used for", "Example"])
fmts = [
    ("Text post", "Quick updates, shout-outs, questions", "“Just closed a deal 🎉” on a coloured background"),
    ("Photo post", "Screenshots, team photos, charts", "Ranking screenshot, event photo"),
    ("Carousel", "Step-by-step, multiple creatives", "3 ad creatives swiped left-right"),
    ("Video", "Demos, walkthroughs, reels", "1-min product demo video"),
    ("GIF post", "Reactions, fun, celebrations", "Celebration GIF on an announcement"),
]
data_rows(ws, 5, fmts, [16, 34, 40], wrap_cols=[2, 3])

trow = 5 + len(fmts) + 1
ws.cell(trow, 1, "Creator tools (Instagram-style styling)").font = font(13, True, BRAND_D)
ws.row_dimensions[trow].height = 24
header_row(ws, trow + 1, ["Tool", "What the user can do", "Phase"])
tools = [
    ("Font style", "Choose from font presets (Classic, Elegant, Mono, Fun, Bold)", "P2"),
    ("Text colour", "Change the colour of the text", "P2"),
    ("Background", "Pick a solid colour or gradient background (text posts)", "P2"),
    ("GIF", "Search & insert a GIF", "P2"),
    ("Music / audio", "Attach a music clip or voice note", "P2"),
    ("Location", "Tag a place", "P2"),
    ("Mentions / #tags", "Tag people and topics for discovery + relevance", "P1"),
]
data_rows(ws, trow + 2, tools, [16, 50, 8], wrap_cols=[2])

# ============================================================
# 4) RELEVANCE ENGINE (the backend filter the CEO asked about)
# ============================================================
ws = wb.create_sheet("4. Relevance Engine")
ws.sheet_view.showGridLines = False
title_block(ws, "Smart Relevance — the backend filter",
            "How the app shows the RIGHT posts to the RIGHT employee, automatically", 4, color=ACCENT)

# how it works
ws.cell(4, 1, "How it works in 3 steps").font = font(13, True, BRAND_D)
steps = [
    ("1. Tag the user", "Every employee has a profile: Department + Specialization + Role (e.g. Marketing > SEO)."),
    ("2. Tag the post", "Every post is auto-tagged from the author's department, keywords and #hashtags."),
    ("3. Match & rank", "Feed shows posts whose tags match the viewer's tags first. Company-wide posts (R&R, newsletter, announcements) are shown to everyone."),
]
for i, (k, v) in enumerate(steps):
    rr = 5 + i; bg = LIGHT2 if i % 2 == 0 else WHITE
    a = ws.cell(rr, 1, k); a.font = font(11, True, BRAND_D); a.fill = fill(bg); a.alignment = left(); a.border = border
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
    b = ws.cell(rr, 2, v); b.font = font(10.5); b.fill = fill(bg); b.alignment = left(); b.border = border
    for cc in range(3, 5): ws.cell(rr, cc).fill = fill(bg); ws.cell(rr, cc).border = border
    ws.row_dimensions[rr].height = 36

mrow = 5 + len(steps) + 1
ws.cell(mrow, 1, "Department → Specialization → Keywords → What they see").font = font(13, True, BRAND_D)
ws.row_dimensions[mrow].height = 24
header_row(ws, mrow + 1, ["Department", "Specialization", "Auto-tags / keywords", "Feed shows them"])
mp = [
    ("Marketing", "SEO", "#seo #serp #backlinks #content #aeo", "Ranking wins, algorithm updates, content tips"),
    ("Marketing", "Paid Media", "#ppc #roas #meta-ads #google-ads #cpl", "Campaign results, creative tests, budgets"),
    ("Marketing", "SMM", "#social #reels #engagement #calendar", "Trends, post performance, content calendar"),
    ("Sales", "Inside / Field", "#pipeline #demo #closed-won #quota #crm", "Deal wins, pitch decks, objection handling"),
    ("Technology", "Eng / DevOps", "#release #deploy #api #bug #uptime", "Ships, incident notes, tech demos"),
    ("HR / People", "People Ops", "#hiring #onboarding #policy #event", "Hiring updates, policies, culture posts"),
    ("Everyone", "Company-wide", "#announcement #rnr #event #newsletter", "R&R, newsletter, leadership notes"),
]
data_rows(ws, mrow + 2, mp, [14, 16, 36, 36], wrap_cols=[3, 4])
ws.column_dimensions["A"].width = 14

# ============================================================
# 5) ADMIN / PEOPLE INSIGHTS
# ============================================================
ws = wb.create_sheet("5. Admin Insights")
ws.sheet_view.showGridLines = False
title_block(ws, "Admin — People Insights", "Search any employee and see their full activity picture", 3, color=INK)
header_row(ws, 4, ["Admin can…", "Detail", "Why it matters"])
adm = [
    ("Search a user", "By name or email id", "Find anyone in seconds"),
    ("See posts count", "How many posts they've made on the feed", "Spot active contributors"),
    ("See comments", "How many comments they've written", "Measure participation"),
    ("See kudos", "Kudos received + given", "Recognise culture champions"),
    ("View activity log", "Recent posts, comments, likes, subscriptions", "Understand behaviour over time"),
    ("View keyword map", "The department tags driving their feed", "Verify relevance is correct"),
    ("Filter by department", "Marketing / Sales / Tech / HR", "Team-level views"),
    ("Engagement analytics", "Active users, top posts, adoption by team (P3)", "Report adoption to leadership"),
]
data_rows(ws, 5, adm, [20, 40, 34], wrap_cols=[2, 3])

# ============================================================
# 6) ROLES & PERMISSIONS
# ============================================================
ws = wb.create_sheet("6. Roles")
ws.sheet_view.showGridLines = False
title_block(ws, "Roles & Permissions", "Who can do what", 5, color=BRAND)
header_row(ws, 4, ["Role", "Post", "React/Comment", "Give Kudos", "Admin / Analytics"])
roles = [
    ("Employee", "Yes", "Yes", "Yes", "No"),
    ("Team Lead", "Yes", "Yes", "Yes", "Team view only"),
    ("HR / People", "Yes", "Yes", "Yes", "User search + R&R"),
    ("Admin", "Yes", "Yes", "Yes", "Full (search, moderate, analytics)"),
    ("Super Admin", "Yes", "Yes", "Yes", "Full + settings & roles"),
]
data_rows(ws, 5, roles, [16, 10, 16, 12, 30], wrap_cols=[5])

# ============================================================
# 7) ROADMAP
# ============================================================
ws = wb.create_sheet("7. Roadmap")
ws.sheet_view.showGridLines = False
title_block(ws, "Roadmap — 3 phases", "A realistic path from MVP to full product", 4, color=BRAND_D)
header_row(ws, 4, ["Phase", "Timeline", "What ships", "Goal"])
road = [
    ("Phase 1 — MVP", "Weeks 1-8",
     "Feed (text/photo/carousel/video), reactions, comments, profiles + department tagging, basic Top-5 leaderboard, admin user search, SSO login.",
     "Get the feed live and people posting."),
    ("Phase 2 — Engage", "Weeks 9-14",
     "Rich composer (font/colour/background/GIF/music/location), kudos & points engine, badges, newsletter subscribe + delivery, notifications, moderation.",
     "Make it fun and rewarding to use daily."),
    ("Phase 3 — Smart & Scale", "Weeks 15-20",
     "Smart relevance ranking (learns from engagement), analytics dashboard, mobile app polish, integrations (HRMS, Slack/Email).",
     "Personalise the feed and prove ROI."),
]
data_rows(ws, 5, road, [18, 13, 55, 24], wrap_cols=[3, 4])
for r in range(5, 8):
    ws.row_dimensions[r].height = 80

# ============================================================
# 8) TECH STACK
# ============================================================
ws = wb.create_sheet("8. Tech & Cost")
ws.sheet_view.showGridLines = False
title_block(ws, "Suggested Tech Stack & Cost Drivers", "What it's built on (engineering can swap equivalents)", 3, color=INK)
header_row(ws, 4, ["Layer", "Suggested tech", "Note"])
tech = [
    ("Mobile app", "React Native (iOS + Android)", "One codebase, both stores"),
    ("Web / Admin", "React + Next.js", "Admin console & web feed"),
    ("Backend / API", "Node.js (NestJS) or Python (Django)", "Use existing team skills"),
    ("Database", "PostgreSQL + Redis (cache)", "Reliable + fast feeds"),
    ("Media storage", "AWS S3 + CloudFront CDN", "Photos / video / GIF"),
    ("Search", "Elasticsearch / OpenSearch", "Fast user & post search"),
    ("Notifications", "Firebase Cloud Messaging (push) + email", "Kudos, mentions, comments"),
    ("Auth", "SSO (Google / company)", "One-click secure login"),
    ("Hosting", "AWS / managed cloud", "Scales with users"),
]
data_rows(ws, 5, tech, [16, 36, 36], wrap_cols=[3])

crow = 5 + len(tech) + 1
ws.cell(crow, 1, "Main cost drivers").font = font(13, True, BRAND_D)
ws.row_dimensions[crow].height = 24
header_row(ws, crow + 1, ["Cost item", "Driver", "Note"])
cost = [
    ("Development", "Phase 1-3 build effort", "Largest one-time cost"),
    ("Cloud hosting", "Servers + database", "Scales with active users"),
    ("Media / CDN", "Video & image storage + bandwidth", "Grows with video usage"),
    ("Push / email", "Notification volume", "Low at start"),
    ("Maintenance", "Ongoing fixes + features", "~15-20% of build/yr"),
]
data_rows(ws, crow + 2, cost, [16, 36, 36], wrap_cols=[3])

# ============================================================
# 9) SUCCESS METRICS
# ============================================================
ws = wb.create_sheet("9. Success Metrics")
ws.sheet_view.showGridLines = False
title_block(ws, "Success Metrics — how we'll measure it", "Numbers to report back to leadership", 3, color=ACCENT)
header_row(ws, 4, ["Metric", "What it tells us", "Target (first 90 days)"])
kpi = [
    ("Active users (DAU/MAU)", "How many log in daily / monthly", "60%+ of staff monthly"),
    ("% contributors", "Share of staff who post (not just read)", "30%+ posting monthly"),
    ("Posts per week", "Content volume / liveliness", "Growing week on week"),
    ("Kudos given", "Recognition activity", "Steady weekly growth"),
    ("Engagement rate", "Likes + comments per post", "Healthy & rising"),
    ("Newsletter opt-in", "Newsletter subscribers", "50%+ subscribed"),
    ("Retention", "% still active after 60 days", "70%+"),
    ("eNPS lift", "Employee satisfaction change", "Measurable uplift"),
]
data_rows(ws, 5, kpi, [24, 40, 26], wrap_cols=[2, 3])

# Save
wb.save(OUT)
print("Saved:", OUT)
