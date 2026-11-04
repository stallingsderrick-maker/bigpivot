"""
Build: BigPivot_StrategicAnalysis(2026-11-04).docx
Full strategic analysis -- gaps, partnership structure, grants, tech requirements
"""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

OUT = r"C:\Users\Sterling\.openclaw\workspace\BigPivot\BigPivot_StrategicAnalysis(2026-11-04).docx"

NAVY  = RGBColor(0x1A,0x35,0x5E)
GOLD  = RGBColor(0xB8,0x86,0x20)
GREEN = RGBColor(0x10,0x60,0x28)
RED   = RGBColor(0xA0,0x00,0x00)
ORANGE= RGBColor(0xBF,0x50,0x00)
BLACK = RGBColor(0x00,0x00,0x00)
WHITE = RGBColor(0xFF,0xFF,0xFF)
PURP  = RGBColor(0x4B,0x0E,0x82)

BODY  = 12
TSIZE = 11

def shade(cell, hex_fill):
    tc=cell._tc; tcPr=tc.get_or_add_tcPr()
    shd=OxmlElement('w:shd')
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto')
    shd.set(qn('w:fill'),hex_fill); tcPr.append(shd)

def hr(doc, color='1A355E', thick=6):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(2); p.paragraph_format.space_after=Pt(2)
    pPr=p._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr')
    bot=OxmlElement('w:bottom')
    bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),str(thick))
    bot.set(qn('w:space'),'1'); bot.set(qn('w:color'),color)
    pBdr.append(bot); pPr.append(pBdr)

def pb(doc): doc.add_page_break()

def h1(doc, t, color=NAVY):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(14); p.paragraph_format.space_after=Pt(3)
    r=p.add_run(t); r.bold=True; r.font.size=Pt(14); r.font.color.rgb=color

def h2(doc, t, color=NAVY):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(9); p.paragraph_format.space_after=Pt(3)
    r=p.add_run(t); r.bold=True; r.font.size=Pt(13); r.font.color.rgb=color

def h3(doc, t, color=GOLD):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(5); p.paragraph_format.space_after=Pt(2)
    r=p.add_run(t); r.bold=True; r.font.size=Pt(12); r.font.color.rgb=color

def body(doc, t, color=BLACK, bold=False, italic=False):
    p=doc.add_paragraph()
    p.paragraph_format.space_after=Pt(4)
    r=p.add_run(t); r.font.size=Pt(BODY); r.font.color.rgb=color
    r.bold=bold; r.font.italic=italic

def bullet(doc, t, color=BLACK, bold=False):
    p=doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after=Pt(3)
    r=p.add_run(t); r.font.size=Pt(BODY); r.font.color.rgb=color; r.bold=bold

def mtable(doc, data, hdr='1A355E', alt='EEF2F8'):
    tbl=doc.add_table(rows=len(data), cols=len(data[0]))
    tbl.style='Table Grid'
    for i,row in enumerate(data):
        for j,val in enumerate(row):
            c=tbl.rows[i].cells[j]; c.text=str(val)
            for para in c.paragraphs:
                para.paragraph_format.space_before=Pt(3)
                para.paragraph_format.space_after=Pt(3)
                for run in para.runs:
                    run.font.size=Pt(TSIZE)
                    if i==0: run.bold=True; run.font.color.rgb=WHITE
                    else: run.font.color.rgb=BLACK
            if i==0: shade(c,hdr)
            elif i%2==0: shade(c,alt)
    doc.add_paragraph()

def callout(doc, label, text, bg='EEF2F8', lc=NAVY, tc=BLACK):
    tbl=doc.add_table(rows=1,cols=1); tbl.style='Table Grid'
    c=tbl.rows[0].cells[0]; shade(c,bg)
    p=c.paragraphs[0]; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6)
    if label:
        r1=p.add_run(label+"  "); r1.bold=True; r1.font.size=Pt(BODY); r1.font.color.rgb=lc
    r2=p.add_run(text); r2.font.size=Pt(BODY); r2.font.color.rgb=tc
    doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════
doc = Document()
for s in doc.sections:
    s.top_margin=Inches(0.9); s.bottom_margin=Inches(0.9)
    s.left_margin=Inches(1.1); s.right_margin=Inches(1.1)

# COVER
tbl=doc.add_table(rows=1,cols=1); tbl.style='Table Grid'
c=tbl.rows[0].cells[0]; shade(c,'1A355E')
p1=c.paragraphs[0]; p1.alignment=WD_ALIGN_PARAGRAPH.CENTER; p1.paragraph_format.space_before=Pt(16)
r1=p1.add_run("THE BIG PIVOT -- STRATEGIC ANALYSIS")
r1.bold=True; r1.font.size=Pt(20); r1.font.color.rgb=WHITE
p2=c.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
r2=p2.add_run("Contractor Marketplace App -- Full Build Strategy")
r2.bold=True; r2.font.size=Pt(13); r2.font.color.rgb=GOLD
p3=c.add_paragraph(); p3.alignment=WD_ALIGN_PARAGRAPH.CENTER
r3=p3.add_run("Gap Analysis  |  Partnership Structure  |  Grants  |  Tech Requirements  |  Go-to-Market")
r3.font.size=Pt(11); r3.font.color.rgb=WHITE
p4=c.add_paragraph(); p4.alignment=WD_ALIGN_PARAGRAPH.CENTER; p4.paragraph_format.space_after=Pt(16)
r4=p4.add_run("November 4, 2026  |  Confidential")
r4.font.size=Pt(10); r4.font.color.rgb=GOLD
doc.add_paragraph()

callout(doc,"THE CONCEPT IN ONE SENTENCE:",
    "A quality-gated, government-contract-ready contractor marketplace that connects pre-vetted tradespeople with "
    "local government, commercial, and residential clients -- with automated payments, subscription revenue, and "
    "a workshop pipeline feeding the platform.",
    'E8F4E8', GREEN, BLACK)

# ══════════════════════════════════════════════════════════════════════
# SECTION 1 -- WHAT'S STRONG
# ══════════════════════════════════════════════════════════════════════
h1(doc,"SECTION 1 -- WHAT IS ALREADY STRONG",GREEN); hr(doc,'106028')
body(doc,"This concept has real competitive advantages built in from day one. These should be protected and amplified:")
mtable(doc,[
    ("Strength","Why It Matters"),
    ("Quality gate -- No Jack Legs policy",
     "Thumbtack and Angi let virtually anyone list. This platform's hard filter (licensed, LLC, business bank account, financial capacity, capability statement) makes it the only marketplace where clients can trust the results without vetting vendors themselves. That is the entire value proposition."),
    ("Government contract focus",
     "Nobody else is building a marketplace specifically designed around Net-30/45 government payment timelines and capability statement requirements. That is an underserved gap in Detroit and Michigan."),
    ("Founder expertise",
     "Regina managed $20,000/week payrolls at 40% markup in staffing. She knows vendor qualification, contract compliance, and cash flow management from lived experience -- not theory."),
    ("Three revenue streams already identified",
     "Transactional cut (10% per job), recurring subscription ($25/month per contractor), and service revenue (bid-ready consulting). Three streams is the right structure for a marketplace."),
    ("Seasonal pipeline awareness",
     "Knowing when each trade type peaks allows proactive marketing, targeted contractor recruitment, and smarter bid pricing by season."),
    ("Workshop as pipeline",
     "The Big Pivot workshop is not just a revenue source -- it is a contractor recruitment funnel. Every workshop attendee is a potential platform subscriber."),
],'106028','EEF2F8')

# ══════════════════════════════════════════════════════════════════════
# SECTION 2 -- WHAT IS MISSING
# ══════════════════════════════════════════════════════════════════════
pb(doc)
h1(doc,"SECTION 2 -- WHAT IS MISSING (GAP ANALYSIS)",RED); hr(doc,'A00000')
body(doc,"These gaps need to be resolved before the platform launches. Some are business-critical, some are legal requirements.")

h2(doc,"Business and Legal Gaps",RED)
mtable(doc,[
    ("Gap","Why It Matters","Action Required"),
    ("No equity split defined between Regina and partner",
     "Without a formal partnership agreement, any dispute about ownership, profits, or decision-making has no resolution mechanism. This is the most dangerous gap.",
     "Execute an operating agreement before any development begins. See Section 3 for recommended structure."),
    ("Sedart Trades LLC status unclear",
     "The document references Sedart Trades as the brand but does not confirm it is a registered LLC with a business bank account.",
     "Confirm Michigan LLC registration at michigan.gov/lara. If not filed, file now ($50). Open a dedicated business bank account immediately."),
    ("No dispute resolution process",
     "What happens when a contractor does bad work? Who refunds the client? How is the platform protected from chargebacks?",
     "Define a formal dispute resolution policy before launch. Consider a 7-day inspection window before releasing final payment to contractor."),
    ("Liability exposure not addressed",
     "If a licensed contractor on the platform causes property damage or injury, the platform could be named in a lawsuit.",
     "Require all contractors to carry general liability insurance ($1M minimum) as a hard platform requirement. Consult an attorney about platform liability shield language in the terms of service."),
    ("Bonding requirements not mentioned",
     "Government contracts often require contractors to carry performance bonds. The platform's vendor requirements do not include bonding.",
     "Add performance bond requirement to vendor qualification for contractors pursuing government work. Make it a filter in the matching engine."),
    ("No background check requirement",
     "Licensing verifies skill -- it does not screen for criminal history. Clients letting contractors into their homes or government buildings need more assurance.",
     "Partner with a background check service (Checkr, Sterling Background) and add criminal background check as a required qualification step."),
    ("Payment processing partner not specified",
     "The document describes a middleman payment gateway but does not name the technology. Building a custom payment system from scratch is high-risk and expensive.",
     "Use Stripe Connect -- the industry standard for marketplace payments with built-in escrow, automatic splits, and 1099 generation for contractors."),
    ("Geographic scope undefined",
     "Is this Detroit only? Michigan? National from day one? The answer determines everything about tech build, marketing spend, and legal requirements.",
     "Start Detroit-focused. Prove the model in one city. Build the expansion playbook for Year 2."),
],'A00000','FAE8E8')

h2(doc,"Revenue and Pricing Gaps",ORANGE)
mtable(doc,[
    ("Gap","Issue","Recommendation"),
    ("VIP workshop tier undervalued",
     "A 30-minute 1-on-1 at $100 prices Regina's time at $200/hr -- below market for her expertise.",
     "Raise VIP to $175 to $200, or eliminate 1-on-1 from group workshop entirely and sell private consulting separately at $250 to $350/hr."),
    ("$25/month subscription may be too low",
     "If the platform is delivering government bids worth $5,000 to $500,000, a $25/month subscription is essentially free. Contractors will not question it. That means the price is probably too low.",
     "Test $49 to $99/month for active bid access. Offer a 60-day free trial to get the first 50 contractors on the platform."),
    ("10% cut vs 70/30 split -- pick one",
     "The document lists both a 10% finder's fee AND a 70/30 split as options. These are very different structures. The revenue model needs to be one consistent system.",
     "For small jobs (under $5,000): flat 10% platform fee. For large government contracts: negotiate a 5% fee (larger contract, same margin). For bid-ready services: flat fee only, no cut."),
    ("Bid-ready service pricing unclear",
     "$1,000 organized vs $1,500 disorganized is a good structure. But what does delivery look like? How long does it take? Who does the work?",
     "Define the deliverable: a completed capability statement, SAM.gov registration, MBE/WBE certification packet, and document folder. Standardize the service so it scales without Regina doing every one manually."),
    ("Net-30/45 cash flow gap not solved",
     "The platform gates entry on financial capacity -- but does not solve the problem of contractors waiting 30 to 45 days for government payment while the platform has already taken its cut.",
     "Consider an invoice factoring partnership (e.g., BlueVine, Fundbox) where contractors can get 80 to 90% of their invoice value immediately and the factoring company collects from the government. The platform earns a referral fee."),
],'BF5000','FFF3CD')

h2(doc,"Technology Gaps",PURP)
mtable(doc,[
    ("Gap","Issue","Action"),
    ("No MVP scope defined","The document describes the full vision -- not what gets built first. A full platform built all at once will take 12 to 18 months and $150,000+. That is too long and too expensive to start.",
     "Define a 90-day MVP: vendor profiles, job posting, basic matching by trade type, Stripe payment integration, and admin dashboard. Launch with that. Add rating system and subscription billing in Month 4."),
    ("Mobile vs web not decided","A contractor in the field needs a mobile app. A client posting a government contract needs a web interface.",
     "Build web-first (React or similar). Use responsive design so it works on mobile browser. Build native mobile app (iOS/Android) in Year 2 once the model is proven."),
    ("No mention of automated verification","How does the platform verify that a contractor's license is real? Manual review by Regina/Derrick does not scale past 50 vendors.",
     "Integrate with state licensing databases via API (many states offer this). For Michigan, use the LARA license lookup API. Manual review for initial launch is acceptable -- automate it in Phase 2."),
    ("Capability statement generator not spec'd","The document mentions helping contractors generate capability statements for a fee -- but does not describe what tool or process does this.",
     "Build a guided form that generates a formatted capability statement as a PDF. Charge $150 to $300 per generation. This is a sellable product on its own."),
],'4B0E82','F0E8FA')

# ══════════════════════════════════════════════════════════════════════
# SECTION 3 -- PARTNERSHIP STRUCTURE
# ══════════════════════════════════════════════════════════════════════
pb(doc)
h1(doc,"SECTION 3 -- PARTNERSHIP STRUCTURE",NAVY); hr(doc)
body(doc,"Derrick brings three things this business needs and cannot easily replace: technical development capability, "
     "government contracting relationships, and the ability to operate in professional environments where Regina's expertise "
     "alone may not always be received as intended. That combination is worth formalizing.")

h2(doc,"Recommended Structure")
mtable(doc,[
    ("Element","Recommendation","Rationale"),
    ("Entity","Sedart Trades LLC (or new co-owned LLC)",
     "If Sedart Trades is already registered to Regina, add Derrick as a member via operating agreement amendment. If starting fresh, form a new LLC together."),
    ("Ownership Split","51% Regina / 49% Derrick",
     "Regina owns the concept, the industry expertise, and the existing brand. 51/49 keeps her as majority owner while making Derrick a true partner. This also preserves WBE certification eligibility (women-owned, majority female)."),
    ("Regina's Contributions","Business concept, industry relationships, vendor qualification expertise, workshop delivery, client outreach",
     "Regina brings the product knowledge and the domain credibility. She is the face of the brand with contractors."),
    ("Derrick's Contributions","Platform development, government/school district relationships, business structure, technical operations, co-representation for select client meetings",
     "Derrick brings the build capability and the institutional credibility in environments that require it. His DoIT and school district connections are direct leads for the platform."),
    ("Compensation -- Pre-Revenue","Derrick invoices for development work at an agreed rate (e.g., $75 to $125/hr) until platform launches. This creates a documented investment that converts to equity credit.",
     "Prevents Derrick from working for free while also preventing Regina from carrying all the financial risk."),
    ("Compensation -- Post-Revenue","Profit distributions proportional to ownership (51/49) after all operating expenses",
     "Clean, simple, and legally standard for LLCs."),
    ("Decision Authority","Regina has final say on vendor qualifications and workshop content. Derrick has final say on technology decisions. Joint decision required for pricing changes and new markets.",
     "Each partner's domain expertise governs their area. Prevents deadlock on day-to-day decisions."),
    ("Exit Clause","Either partner can exit with 90-day notice. Departing partner's interest goes to the other at fair market value (assessed by agreed formula).",
     "Protects both parties if the relationship does not work out. Without this clause, a partner who wants to leave has no clean path."),
],'1A355E')

callout(doc,"CRITICAL STEP:",
    "Before any development begins, before any grant is applied for, and before any client conversation happens -- "
    "both parties need a signed operating agreement drafted by a Michigan business attorney. "
    "The cost is $500 to $1,500 and it protects both of them. Everything else depends on this.",
    'FFE8E8', RED, BLACK)

h2(doc,"Certifications to Pursue Immediately")
mtable(doc,[
    ("Certification","Issuing Body","Value","Eligibility"),
    ("WBE -- Women's Business Enterprise","WBEC Great Lakes",
     "Unlocks corporate supplier diversity contracts (DTE, Ford, GM, Stellantis). Required for many government RFPs.",
     "Regina is eligible as majority female owner. Apply at wbecgreatlakes.org"),
    ("DBE -- Disadvantaged Business Enterprise","Michigan DOT (MDOT)",
     "Required for federally funded transportation and infrastructure contracts. Opens a large government contract pipeline.",
     "Regina qualifies as woman + potential socially disadvantaged status. Apply at michigan.gov/mdot"),
    ("Detroit Business Opportunity Program (MBE/WBE)","City of Detroit -- Civil Rights",
     "Required to bid on City of Detroit contracts as a minority or women-owned vendor. Direct path to City procurement.",
     "Apply at detroitmi.gov/departments/civil-rights-inclusion-opportunity-department/detroit-business-opportunity-program"),
    ("MMSDC MBE","Michigan Minority Supplier Development Council",
     "Opens corporate supplier diversity across Michigan's largest employers. Pairs well with WBE.",
     "If Regina qualifies as minority-owned (or with a future Black co-owner structure), apply at mmsdc.org"),
],'106028')

# ══════════════════════════════════════════════════════════════════════
# SECTION 4 -- GRANT FUNDING
# ══════════════════════════════════════════════════════════════════════
pb(doc)
h1(doc,"SECTION 4 -- GRANT FUNDING OPPORTUNITIES",NAVY); hr(doc)
body(doc,"Multiple funding sources exist for a Detroit-based, women-owned tech platform serving contractors and small businesses. "
     "These should be pursued in parallel with platform development -- not after.")

mtable(doc,[
    ("Grant / Fund","Funder","Amount","Fit","First Step"),
    ("New Economy Initiative (NEI)","Community Foundation for SE Michigan",
     "$50,000 to $200,000+",
     "STRONG -- NEI specifically funds SE Michigan small business support organizations and tech platforms that serve entrepreneurs. This platform serves contractors directly.",
     "Apply at neweconomyinitiative.org. The platform itself, plus the workshop, positions as an entrepreneur support organization."),
    ("Detroit Small Business Technology Fund","City of Detroit / Rocket Community Fund",
     "Up to $1,000 (microgrant)",
     "MODERATE -- Small amount but accessible. Eligibility: Detroit-based, 10 or fewer employees, under $500K revenue.",
     "Apply through a partner organization. Check detroitmi.gov for current partner list."),
    ("MEDC Entrepreneur Support Grant","Michigan Economic Development Corporation",
     "Varies -- program awarded $8.8M total in 2025",
     "STRONG -- MEDC specifically funds organizations supporting minority and women entrepreneurs. This platform IS that.",
     "Apply at michiganbusiness.org. Watch for next grant cycle announcement."),
    ("Motor City Match","City of Detroit / DEGC",
     "Up to $100,000 (cash + in-kind)",
     "MODERATE -- Primarily for businesses with physical space, but tech companies have been funded. Round 33 closes October 1, 2026.",
     "Check if Round 34 opens. Apply at motorcitymatch.com"),
    ("WIOA Workforce Innovation -- Adult",
     "U.S. Dept. of Labor via Michigan Works",
     "$50,000 to $150,000",
     "STRONG -- The platform trains and qualifies contractors (workforce development). The bid-ready services and capability statement generator are workforce development activities.",
     "Contact Southeast Michigan Works! (semworks.net) to discuss grant eligibility."),
    ("SBA Community Advantage / Microloan",
     "U.S. Small Business Administration",
     "Up to $50,000",
     "MODERATE -- This is a loan, not a grant, but at low interest rates. Good for initial platform build costs.",
     "Contact Detroit SCORE or DEGC for SBA microloan access."),
    ("Ralph C. Wilson Jr. Foundation",
     "Ralph C. Wilson Jr. Foundation",
     "Varies",
     "MODERATE -- Focuses on SE Michigan entrepreneurship and community development. A Detroit contractor platform fits.",
     "Apply at rcwjrf.org -- watch for open grant cycles."),
    ("JPMorgan Chase Small Business Grant",
     "JPMorgan Chase / Advancing Black Pathways",
     "$30,000 to $100,000",
     "STRONG -- JPMorgan has a Detroit-specific small business grant program. Regina as a Black woman entrepreneur is a strong fit.",
     "Watch chase.com/business/small-business-grants for open cycles."),
],'1A355E','EEF2F8')

# ══════════════════════════════════════════════════════════════════════
# SECTION 5 -- TECH BUILD PLAN
# ══════════════════════════════════════════════════════════════════════
pb(doc)
h1(doc,"SECTION 5 -- WHAT NEEDS TO BE BUILT",NAVY); hr(doc)
body(doc,"Derrick needs a clear technical scope before any development begins. This is a phased build -- not everything at once.")

h2(doc,"Phase 1 -- MVP (90 Days) -- What You MUST Build to Launch")
mtable(doc,[
    ("Feature","Description","Tech Approach"),
    ("Vendor application and profile","Contractors fill out an application: license info, LLC docs, bank account verification, trade type, service area. Admin reviews and approves.",
     "Web form with document upload. Admin dashboard for review. No automation yet -- manual approval is fine for the first 100 vendors."),
    ("Client job posting","Clients (residential, commercial, or government) post a job: trade type, location, timeline, budget range.",
     "Simple web form. Email notification to matched vendors."),
    ("Basic matching engine","Filter contractors by trade type and service area. Send job description to all matches via email.",
     "No AI needed at launch. Simple database query: trade type = client request AND service area includes client location."),
    ("Payment gateway with escrow","Client pays deposit. Platform holds it. Contractor gets paid on completion.",
     "Stripe Connect -- handles escrow, automatic splits, 1099 generation, and dispute management. This is non-negotiable -- do not build a custom payment system."),
    ("Admin dashboard","Regina and Derrick can see all vendors, all jobs, all payments, approve/reject vendor applications.",
     "Simple internal admin panel. Can be built with a tool like Retool or custom-coded."),
    ("Basic vendor profiles","Public-facing profile: photo, bio, trade types, service area, star rating, Google review import.",
     "Static profile pages. Google review import via Google Places API (free tier is sufficient at launch)."),
],'1A355E')

h2(doc,"Phase 2 -- Full Platform (Months 4 to 9)")
mtable(doc,[
    ("Feature","Description"),
    ("Subscription billing","$25 to $99/month contractor subscription via Stripe recurring billing. Access to bid notifications."),
    ("Rating and review system","Clients rate contractors after job completion. 1 to 5 stars + written review. Ratings visible on profile."),
    ("Bidding system","Vendors submit bids on posted jobs. Client selects best bid. Platform facilitates the selection."),
    ("Capability statement generator","Guided form that produces a formatted, government-ready capability statement as a PDF download."),
    ("License verification integration","Michigan LARA license lookup API to auto-verify contractor license status during application."),
    ("Mobile-responsive UI","Ensure all web interfaces work cleanly on a smartphone browser. No native mobile app yet."),
    ("Automated seasonal marketing","Email campaigns to contractors and clients based on seasonal trade demand calendar."),
],'106028')

h2(doc,"Phase 3 -- Scale (Year 2)")
for item in [
    "Native iOS and Android mobile apps for contractors (job alerts, bid submission, payment tracking)",
    "Expansion to other Michigan cities (Flint, Lansing, Grand Rapids)",
    "Invoice factoring partnership integration (BlueVine or Fundbox) for Net-30/45 cash flow",
    "AI-powered matching (move beyond simple trade-type filtering to rating, past performance, and availability)",
    "Government procurement portal integration (City of Detroit eProcurement, DWSD vendor portal)",
]:
    bullet(doc, item)
doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════
# SECTION 6 -- GO TO MARKET
# ══════════════════════════════════════════════════════════════════════
h1(doc,"SECTION 6 -- HOW TO GET THE FIRST CUSTOMERS",NAVY); hr(doc)

h2(doc,"First 90 Days -- Launch Strategy")
mtable(doc,[
    ("Goal","Target","How"),
    ("50 qualified vendors on the platform","Licensed Detroit-area contractors in HVAC, plumbing, electrical, landscaping, roofing",
     "Regina runs a Big Pivot workshop -- every attendee is offered free 60-day platform access. Direct outreach to contractors she already knows from her staffing days."),
    ("10 paying clients posting jobs","Small Detroit businesses, property managers, school districts",
     "Derrick leverages his DoIT and school district relationships. UPrep alone has 10 campuses that need facility vendors."),
    ("First government contract placed through the platform","One City of Detroit or school district job",
     "Use Derrick's existing City relationships. Even a $5,000 landscaping or cleaning contract proves the model."),
    ("Thompson Foundation connection activated","Introductory meeting with new Thompson Foundation leadership",
     "Regina re-engages her contact. Frame it as a contract to supply vetted vendors to school system operations -- not a job offer."),
],'1A355E')

h2(doc,"Revenue Projection -- Year 1")
mtable(doc,[
    ("Stream","Conservative","Moderate","Strong"),
    ("Contractor subscriptions (50 to 200 contractors x $49/mo x 12 mo)","$29,400","$73,500","$117,600"),
    ("Platform cut on jobs (10% of job value -- $500K to $2M in jobs placed)","$50,000","$100,000","$200,000"),
    ("Bid-ready consulting ($1,000 to $1,500 x 20 to 60 contractors)","$20,000","$50,000","$90,000"),
    ("Big Pivot workshops ($50 to $100 x 10 to 30 workshops x 20 attendees)","$10,000","$30,000","$60,000"),
    ("Capability statement generation ($200 x 20 to 100)","$4,000","$12,000","$20,000"),
    ("TOTAL YEAR 1","$113,400","$265,500","$487,600"),
],'106028')

callout(doc,"BOTTOM LINE:",
    "This is a real business with multiple proven revenue streams. The concept is solid. "
    "The gaps are fixable -- most of them before a single line of code is written. "
    "Get the operating agreement signed, file for WBE certification, pick Stripe Connect as the payment layer, "
    "and start the NEI and JPMorgan grant applications this month. "
    "Derrick builds the MVP in 90 days. Regina fills it with contractors from the workshop. "
    "The Thompson Foundation is the first big client conversation.",
    'E8F4E8', GREEN, BLACK)

doc.save(OUT)
print(f"DONE: {OUT}")
