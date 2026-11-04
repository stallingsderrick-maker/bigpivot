"""
Build: BigPivot_TechnicalRequirements(2026-11-04).docx
Full spec document for Claude Code / developer to build the app
"""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = r"C:\Users\Sterling\.openclaw\workspace\BigPivot\BigPivot_TechnicalRequirements(2026-11-04).docx"

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

def h1(doc, t, color=NAVY):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(16); p.paragraph_format.space_after=Pt(3)
    r=p.add_run(t); r.bold=True; r.font.size=Pt(15); r.font.color.rgb=color

def h2(doc, t, color=NAVY):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(3)
    r=p.add_run(t); r.bold=True; r.font.size=Pt(13); r.font.color.rgb=color

def h3(doc, t, color=GOLD):
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(2)
    r=p.add_run(t); r.bold=True; r.font.size=Pt(12); r.font.color.rgb=color

def body(doc, t, bold=False, color=BLACK):
    p=doc.add_paragraph()
    p.paragraph_format.space_after=Pt(4)
    r=p.add_run(t); r.font.size=Pt(BODY); r.font.color.rgb=color; r.bold=bold

def bullet(doc, t, color=BLACK, bold=False):
    p=doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after=Pt(2)
    r=p.add_run(t); r.font.size=Pt(BODY); r.font.color.rgb=color; r.bold=bold

def code(doc, t):
    p=doc.add_paragraph()
    p.paragraph_format.space_after=Pt(2)
    r=p.add_run(t); r.font.size=Pt(10); r.font.name='Courier New'; r.font.color.rgb=PURP

def mtable(doc, data, hdr='1A355E', alt='EEF2F8', col_widths=None):
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

def pb(doc): doc.add_page_break()

# ═══════════════════════════════════════════════════════════════
doc = Document()
for s in doc.sections:
    s.top_margin=Inches(1.0); s.bottom_margin=Inches(1.0)
    s.left_margin=Inches(1.1); s.right_margin=Inches(1.1)

# COVER BANNER
tbl=doc.add_table(rows=1,cols=1); tbl.style='Table Grid'
c=tbl.rows[0].cells[0]; shade(c,'1A355E')
for txt, sz, col in [
    ("THE BIG PIVOT -- PLATFORM", 20, WHITE),
    ("Full Technical Requirements & Developer Specification", 13, GOLD),
    ("Version 1.0  |  November 4, 2026  |  Confidential", 11, WHITE),
]:
    p=c.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before=Pt(8); p.paragraph_format.space_after=Pt(4)
    r=p.add_run(txt); r.bold=(sz>11); r.font.size=Pt(sz); r.font.color.rgb=col
doc.add_paragraph()

callout(doc,"PURPOSE:",
    "This document is the complete technical specification for The Big Pivot -- a contractor marketplace platform. "
    "It is written so that an AI coding agent (such as Claude Code) or any developer can build the full application "
    "from scratch using only this document as the source of truth. "
    "Every section is complete, specific, and actionable.",
    'E8F4E8', GREEN, BLACK)

# ═══ SECTION 1: PRODUCT OVERVIEW ═══════════════════════════════
h1(doc,"SECTION 1 -- PRODUCT OVERVIEW")
body(doc,"The Big Pivot is a web and mobile application that connects pre-vetted, government-contract-ready skilled tradespeople "
     "with clients that have real work -- local government agencies, commercial property managers, school districts, and homeowners. "
     "It is Detroit-first and quality-obsessed: only fully licensed, incorporated, financially stable contractors get in.")

h2(doc,"1.1 Core Value Proposition")
mtable(doc,[
    ("For Clients","For Contractors"),
    ("Find qualified contractors in minutes -- not days","Access government and commercial bids you can't find on your own"),
    ("Every contractor is pre-vetted -- licensed, LLC, bonded","Get paid through a secure escrow system -- no chasing checks"),
    ("Automated payments with escrow protection","Build your government contract portfolio with bid-ready tools"),
    ("Rate contractors and see real reviews","$49/month gives you unlimited access to curated bid opportunities"),
],'106028','EEF2F8')

h2(doc,"1.2 Target Users")
mtable(doc,[
    ("User Type","Description","Primary Actions"),
    ("Client","Residential homeowner, property manager, commercial business, or government agency needing skilled trade work",
     "Post jobs, review bids, select contractor, make payments, leave reviews"),
    ("Contractor","Licensed, incorporated skilled tradesperson -- HVAC, electrical, plumbing, landscaping, roofing, concrete, cleaning, general contracting",
     "Complete profile, receive bid opportunities, submit bids, complete jobs, receive payments"),
    ("Platform Admin","Regina and Derrick (platform operators)",
     "Approve/reject contractor applications, manage disputes, view analytics, manage subscriptions"),
],'1A355E','EEF2F8')

h2(doc,"1.3 Geographic Scope -- Phase 1")
body(doc,"Detroit, Michigan and surrounding metro area (Wayne, Oakland, Macomb counties). "
     "Phase 2 expands to Flint, Lansing, Grand Rapids. Phase 3 goes statewide Michigan.")

# ═══ SECTION 2: FEATURES ════════════════════════════════════════
pb(doc)
h1(doc,"SECTION 2 -- COMPLETE FEATURE LIST")

h2(doc,"2.1 MVP Features (Build These First -- 90 Days)")
mtable(doc,[
    ("Feature","Description","Priority"),
    ("Contractor Application","Multi-step application form: license info, LLC documentation, trade type, service area, business bank account verification, photo upload","P1 -- Must Have"),
    ("Admin Review Dashboard","Regina/Derrick approve or reject contractor applications. View all submitted documents. Send approval/rejection email automatically.","P1 -- Must Have"),
    ("Contractor Profile Page","Public-facing page: photo, bio, trade types, service area, star rating, review count, verification badges (Licensed, LLC, Bonded)","P1 -- Must Have"),
    ("Client Job Posting","Form: job title, trade type, description, location, timeline, budget range. Submits to matching engine.","P1 -- Must Have"),
    ("Basic Matching Engine","When a job is posted: query database for contractors matching trade type + service area. Send email notification to all matches with job details.","P1 -- Must Have"),
    ("Payment Gateway (Stripe Connect)","Client pays deposit. Platform holds in escrow. Platform takes 10% cut automatically. Contractor receives remaining balance upon job completion.","P1 -- Must Have"),
    ("Contractor Subscription Billing","$49/month recurring subscription via Stripe. Contractors must have active subscription to receive bid notifications.","P1 -- Must Have"),
    ("Review and Rating System","After job completion, client rates contractor 1-5 stars with written review. Rating appears on public profile immediately.","P1 -- Must Have"),
    ("Email Notification System","Automated emails: job matches to contractors, application status, payment confirmations, review received","P1 -- Must Have"),
    ("Authentication","Email/password login. Role-based access: contractor, client, admin. JWT tokens. Email verification on signup.","P1 -- Must Have"),
],'106028')

h2(doc,"2.2 Phase 2 Features (Months 4-9)")
mtable(doc,[
    ("Feature","Description"),
    ("Bidding System","Contractor receives job posting, submits a bid (price + timeline + message). Client reviews all bids and selects one. Platform notifies the winner."),
    ("Capability Statement Generator","Guided form generates a formatted government-ready capability statement PDF. Sold as a service: $200 per statement."),
    ("Michigan LARA License Verification","API call to Michigan LARA license lookup to auto-verify contractor license during application. Flag expired or inactive licenses."),
    ("Google Reviews Import","Pull contractor's Google Business reviews via Google Places API and display on profile."),
    ("Seasonal Email Campaigns","Automated marketing emails to contractors and clients based on seasonal trade demand calendar (HVAC in winter, landscaping in spring, etc.)"),
    ("Invoice Factoring Integration","Partner with BlueVine or Fundbox. Contractors can click 'Get Paid Now' on a completed job and receive 85% immediately vs waiting for client net-30/45."),
    ("Contractor Document Vault","Secure storage for contractor documents: license, LLC cert, insurance certificate, W-9. Shared with clients on request."),
],'1A355E')

h2(doc,"2.3 Phase 3 Features (Year 2)")
for item in [
    "Native iOS and Android mobile apps (React Native)",
    "Advanced matching algorithm (AI-based: rating, past performance, availability, win rate)",
    "Government procurement portal integration (City of Detroit eProcurement, DWSD vendor portal)",
    "Statewide expansion (Flint, Lansing, Grand Rapids, Kalamazoo)",
    "Contractor analytics dashboard (earnings trends, win rate, response time metrics)",
    "Client portal with multi-job management and vendor favorites list",
]:
    bullet(doc, item)
doc.add_paragraph()

# ═══ SECTION 3: DATA MODELS ══════════════════════════════════════
pb(doc)
h1(doc,"SECTION 3 -- DATA MODELS (DATABASE SCHEMA)")

h2(doc,"3.1 Users Table")
mtable(doc,[
    ("Field","Type","Description"),
    ("id","UUID (PK)","Unique user identifier"),
    ("email","VARCHAR(255) UNIQUE","Login email"),
    ("password_hash","VARCHAR(255)","Bcrypt hashed password"),
    ("role","ENUM('client','contractor','admin')","User role"),
    ("first_name","VARCHAR(100)","First name"),
    ("last_name","VARCHAR(100)","Last name"),
    ("phone","VARCHAR(20)","Phone number"),
    ("email_verified","BOOLEAN","Email verification status"),
    ("created_at","TIMESTAMP","Account creation date"),
    ("updated_at","TIMESTAMP","Last update"),
],'1A355E')

h2(doc,"3.2 Contractor Profiles Table")
mtable(doc,[
    ("Field","Type","Description"),
    ("id","UUID (PK)","Unique profile ID"),
    ("user_id","UUID (FK -> users)","Linked user account"),
    ("business_name","VARCHAR(255)","LLC or business name"),
    ("bio","TEXT","Short professional bio (1-2 paragraphs)"),
    ("trade_types","VARCHAR[] (array)","List: ['hvac','electrical','plumbing','landscaping','roofing','concrete','cleaning','general']"),
    ("service_area_zips","VARCHAR[]","Array of ZIP codes served"),
    ("hourly_rate_min","INTEGER","Minimum hourly rate (dollars)"),
    ("hourly_rate_max","INTEGER","Maximum hourly rate (dollars)"),
    ("license_number","VARCHAR(100)","State license number"),
    ("license_state","VARCHAR(2)","State abbreviation"),
    ("license_verified","BOOLEAN","Admin-verified license"),
    ("llc_verified","BOOLEAN","Admin-verified LLC registration"),
    ("insurance_verified","BOOLEAN","Admin-verified insurance certificate"),
    ("bond_verified","BOOLEAN","Admin-verified bonding"),
    ("capability_statement_url","VARCHAR(500)","Link to capability statement PDF"),
    ("profile_photo_url","VARCHAR(500)","Profile photo URL"),
    ("application_status","ENUM('pending','approved','rejected','suspended')","Current status"),
    ("stripe_account_id","VARCHAR(255)","Stripe Connect account ID for payments"),
    ("subscription_status","ENUM('active','past_due','canceled','none')","Subscription status"),
    ("avg_rating","DECIMAL(3,2)","Computed average rating 0.00-5.00"),
    ("review_count","INTEGER","Total reviews received"),
    ("created_at","TIMESTAMP","Profile creation date"),
],'1A355E')

h2(doc,"3.3 Jobs Table")
mtable(doc,[
    ("Field","Type","Description"),
    ("id","UUID (PK)","Unique job ID"),
    ("client_id","UUID (FK -> users)","Client who posted the job"),
    ("title","VARCHAR(255)","Job title"),
    ("description","TEXT","Full job description"),
    ("trade_type","VARCHAR(50)","Required trade type"),
    ("location_address","VARCHAR(500)","Job site address"),
    ("location_zip","VARCHAR(10)","ZIP code"),
    ("timeline_start","DATE","Requested start date"),
    ("timeline_end","DATE","Requested end date"),
    ("budget_min","INTEGER","Minimum budget (dollars)"),
    ("budget_max","INTEGER","Maximum budget (dollars)"),
    ("payment_type","ENUM('fixed','hourly')","Payment structure"),
    ("status","ENUM('open','matched','in_progress','completed','disputed','canceled')","Job status"),
    ("selected_contractor_id","UUID (FK -> contractor_profiles)","Chosen contractor (null until selected)"),
    ("created_at","TIMESTAMP","Job posting date"),
],'1A355E')

h2(doc,"3.4 Bids Table")
mtable(doc,[
    ("Field","Type","Description"),
    ("id","UUID (PK)","Unique bid ID"),
    ("job_id","UUID (FK -> jobs)","Job being bid on"),
    ("contractor_id","UUID (FK -> contractor_profiles)","Bidding contractor"),
    ("amount","INTEGER","Bid amount in dollars"),
    ("timeline_days","INTEGER","Estimated days to complete"),
    ("cover_letter","TEXT","Contractor's message to client"),
    ("status","ENUM('pending','accepted','rejected','withdrawn')","Bid status"),
    ("created_at","TIMESTAMP","Bid submission date"),
],'1A355E')

h2(doc,"3.5 Payments Table")
mtable(doc,[
    ("Field","Type","Description"),
    ("id","UUID (PK)","Unique payment ID"),
    ("job_id","UUID (FK -> jobs)","Related job"),
    ("client_id","UUID (FK -> users)","Paying client"),
    ("contractor_id","UUID (FK -> contractor_profiles)","Receiving contractor"),
    ("amount_total","INTEGER","Total paid by client (cents)"),
    ("platform_fee","INTEGER","10% platform cut (cents)"),
    ("contractor_payout","INTEGER","90% to contractor (cents)"),
    ("stripe_payment_intent_id","VARCHAR(255)","Stripe payment intent ID"),
    ("stripe_transfer_id","VARCHAR(255)","Stripe transfer ID to contractor"),
    ("status","ENUM('pending','held','released','refunded','disputed')","Payment status"),
    ("created_at","TIMESTAMP","Payment date"),
],'1A355E')

h2(doc,"3.6 Reviews Table")
mtable(doc,[
    ("Field","Type","Description"),
    ("id","UUID (PK)","Unique review ID"),
    ("job_id","UUID (FK -> jobs)","Reviewed job"),
    ("reviewer_id","UUID (FK -> users)","Client who left review"),
    ("contractor_id","UUID (FK -> contractor_profiles)","Reviewed contractor"),
    ("rating","INTEGER (1-5)","Star rating"),
    ("review_text","TEXT","Written review"),
    ("created_at","TIMESTAMP","Review date"),
],'1A355E')

# ═══ SECTION 4: TECH STACK ══════════════════════════════════════
pb(doc)
h1(doc,"SECTION 4 -- RECOMMENDED TECH STACK")

h2(doc,"4.1 Stack Overview")
mtable(doc,[
    ("Layer","Technology","Why"),
    ("Frontend","Next.js 14 (App Router) + TypeScript","Industry standard React framework. Server-side rendering for SEO. File-based routing. Easy deployment."),
    ("Styling","Tailwind CSS + shadcn/ui","Fast professional styling. Consistent design system. Saves weeks of UI work."),
    ("Backend / API","Next.js API Routes (same project)","No separate backend server needed for MVP. All in one project. Easy to scale later."),
    ("Database","PostgreSQL (via Supabase)","Reliable SQL database. Supabase adds auth, real-time, and file storage. Free tier is generous."),
    ("ORM","Prisma","Type-safe database queries. Easy schema migrations. Works perfectly with Next.js."),
    ("Authentication","NextAuth.js","Handles sessions, JWT, email/password, and social login. Integrates with Prisma."),
    ("Payments","Stripe Connect","Industry standard for marketplace payments. Handles escrow, automatic splits, subscriptions, and 1099 generation."),
    ("Email","Resend + React Email","Simple transactional email API. React Email lets you build beautiful email templates in JSX."),
    ("File Storage","Supabase Storage","For profile photos, license documents, capability statements. Free tier: 1GB."),
    ("Deployment","Vercel","One-click deployment from GitHub. Auto-deploys on every push. Free tier handles MVP traffic easily."),
    ("Version Control","GitHub","Source code hosting. Enables collaboration and rollback."),
],'1A355E')

h2(doc,"4.2 Project Folder Structure")
for line in [
    "bigpivot/",
    "  app/                    -- Next.js App Router pages",
    "    (auth)/               -- Login, signup, verify-email",
    "    (dashboard)/          -- Contractor dashboard (protected)",
    "    (admin)/              -- Admin panel (protected, admin only)",
    "    contractors/          -- Public contractor listing + profiles",
    "    jobs/                 -- Post a job, view jobs",
    "    api/                  -- API routes (backend logic)",
    "      auth/               -- NextAuth endpoints",
    "      contractors/        -- Contractor CRUD",
    "      jobs/               -- Job posting and matching",
    "      payments/           -- Stripe webhooks and payment logic",
    "      subscriptions/      -- Stripe subscription management",
    "  components/             -- Reusable React components",
    "    ui/                   -- shadcn/ui base components",
    "    contractors/          -- ContractorCard, ContractorProfile",
    "    jobs/                 -- JobCard, JobPostForm",
    "    payments/             -- PaymentForm, EscrowStatus",
    "  lib/                    -- Utility functions",
    "    db.ts                 -- Prisma client singleton",
    "    stripe.ts             -- Stripe client setup",
    "    email.ts              -- Resend email client",
    "    matching.ts           -- Job-contractor matching logic",
    "  prisma/",
    "    schema.prisma         -- Database schema (all models above)",
    "    migrations/           -- Auto-generated by Prisma migrate",
    "  public/                 -- Static assets (logo, icons)",
    "  .env.local              -- Environment variables (never commit!)",
]:
    code(doc, line)
doc.add_paragraph()

# ═══ SECTION 5: ENVIRONMENT VARIABLES ═══════════════════════════
h1(doc,"SECTION 5 -- ENVIRONMENT VARIABLES")
body(doc,"Create a .env.local file in the project root with these variables. NEVER commit this file to GitHub. Add .env.local to .gitignore immediately.")
mtable(doc,[
    ("Variable","Where to Get It","Used For"),
    ("DATABASE_URL","Supabase project settings","PostgreSQL connection string"),
    ("NEXTAUTH_SECRET","Run: openssl rand -base64 32","Session encryption"),
    ("NEXTAUTH_URL","http://localhost:3000 (dev) / your domain (prod)","Auth redirect URL"),
    ("STRIPE_SECRET_KEY","Stripe Dashboard -- API Keys","Server-side Stripe calls"),
    ("STRIPE_PUBLISHABLE_KEY","Stripe Dashboard -- API Keys","Client-side Stripe elements"),
    ("STRIPE_WEBHOOK_SECRET","Stripe Dashboard -- Webhooks","Verify webhook signatures"),
    ("NEXT_PUBLIC_STRIPE_PRICE_ID","Stripe Dashboard -- Products","Contractor subscription price ID"),
    ("RESEND_API_KEY","resend.com -- API Keys","Sending transactional emails"),
    ("NEXT_PUBLIC_SUPABASE_URL","Supabase project settings","File storage URL"),
    ("NEXT_PUBLIC_SUPABASE_ANON_KEY","Supabase project settings","File storage public key"),
    ("GOOGLE_PLACES_API_KEY","Google Cloud Console","Google review import (Phase 2)"),
],'106028')

# ═══ SECTION 6: KEY BUSINESS LOGIC ══════════════════════════════
pb(doc)
h1(doc,"SECTION 6 -- KEY BUSINESS LOGIC")

h2(doc,"6.1 The Matching Engine")
body(doc,"When a client posts a job, the platform immediately finds and notifies matching contractors. Here is exactly how it works:")
for step in [
    "1. Client submits job posting form with trade_type and location_zip",
    "2. Query contractor_profiles WHERE application_status = 'approved' AND subscription_status = 'active' AND trade_types contains job.trade_type AND service_area_zips contains job.location_zip",
    "3. Sort results by avg_rating DESC (highest rated first)",
    "4. Send email to each matching contractor with: job title, description, budget range, location, link to submit bid",
    "5. Update job status to 'matched'",
    "6. If zero contractors match: send admin notification to review and manually outreach",
]:
    bullet(doc, step)
doc.add_paragraph()

h2(doc,"6.2 Payment Flow (Stripe Connect)")
mtable(doc,[
    ("Step","Who","What Happens"),
    ("1. Contractor Onboards","Contractor","Contractor connects their bank account via Stripe Connect onboarding flow during profile setup"),
    ("2. Client Pays Deposit","Client","Client pays 50% of agreed job amount through Stripe. Funds held in platform escrow."),
    ("3. Job Starts","System","Job status updated to 'in_progress'. Both parties notified."),
    ("4. Job Completed","Contractor","Contractor marks job complete. Client receives notification to review and confirm."),
    ("5. Client Confirms","Client","Client confirms completion within 72 hours (or auto-confirms after 72 hours)."),
    ("6. Final Payment","Client","Client pays remaining 50% balance."),
    ("7. Platform Takes Cut","System","Platform automatically takes 10% of total job amount as platform fee via Stripe"),
    ("8. Contractor Paid","System","Remaining 90% transferred to contractor Stripe account. Available per their bank transfer schedule."),
    ("9. Review Prompted","System","Email sent to client prompting review. Email sent to contractor to review client."),
],'1A355E')

h2(doc,"6.3 Contractor Subscription Logic")
for item in [
    "New contractor applications do NOT require active subscription to apply",
    "Once approved by admin, contractor must start $49/month subscription to receive bid notifications",
    "Subscription managed via Stripe recurring billing",
    "If subscription lapses (payment fails), contractor is immediately removed from matching engine and notified",
    "Contractor can pause subscription (stops receiving bids but profile stays active) -- max 90 days",
    "Platform admin can grant free trial period (up to 60 days) for early adopters",
]:
    bullet(doc, item)
doc.add_paragraph()

h2(doc,"6.4 Dispute Resolution Process")
for item in [
    "Client opens dispute within 7 days of job completion confirmation",
    "Admin receives notification and reviews dispute details from both parties",
    "Platform admin has 5 business days to make a ruling",
    "Ruling options: Full refund to client, Partial refund, Full payout to contractor",
    "Contractor with 2+ disputes in 6 months is flagged for review and may be suspended",
    "All dispute communications logged in admin dashboard",
]:
    bullet(doc, item)
doc.add_paragraph()

# ═══ SECTION 7: API ENDPOINTS ════════════════════════════════════
pb(doc)
h1(doc,"SECTION 7 -- API ENDPOINTS")
body(doc,"These are the Next.js API routes to build. All protected routes require a valid session token.")

h2(doc,"7.1 Authentication")
mtable(doc,[
    ("Method","Endpoint","Description","Auth"),
    ("POST","api/auth/signup","Create new user account","Public"),
    ("POST","api/auth/signin","Login with email/password (via NextAuth)","Public"),
    ("POST","api/auth/signout","Logout and clear session","Authenticated"),
    ("GET","api/auth/verify-email","Verify email from link","Public"),
    ("POST","api/auth/forgot-password","Send password reset email","Public"),
    ("POST","api/auth/reset-password","Reset password with token","Public"),
],'106028')

h2(doc,"7.2 Contractors")
mtable(doc,[
    ("Method","Endpoint","Description","Auth"),
    ("GET","api/contractors","List approved contractors (with filter params: trade, zip, min_rating)","Public"),
    ("GET","api/contractors/[id]","Get single contractor profile","Public"),
    ("POST","api/contractors/apply","Submit contractor application","Authenticated (new contractor)"),
    ("PUT","api/contractors/[id]","Update own profile","Contractor only"),
    ("GET","api/contractors/[id]/reviews","Get contractor reviews","Public"),
    ("POST","api/contractors/stripe-connect","Start Stripe Connect onboarding","Contractor only"),
],'1A355E')

h2(doc,"7.3 Jobs")
mtable(doc,[
    ("Method","Endpoint","Description","Auth"),
    ("POST","api/jobs","Post a new job (triggers matching engine)","Authenticated client"),
    ("GET","api/jobs","List client's own jobs","Authenticated client"),
    ("GET","api/jobs/[id]","Get job details","Authenticated"),
    ("PUT","api/jobs/[id]/complete","Mark job as complete","Contractor on job"),
    ("PUT","api/jobs/[id]/confirm","Client confirms completion","Client on job"),
    ("POST","api/jobs/[id]/dispute","Open a dispute","Authenticated party on job"),
],'106028')

h2(doc,"7.4 Bids")
mtable(doc,[
    ("Method","Endpoint","Description","Auth"),
    ("POST","api/bids","Submit bid on a job","Active contractor"),
    ("GET","api/bids/job/[jobId]","Get all bids for a job","Client who owns job"),
    ("GET","api/bids/contractor","Get own submitted bids","Contractor only"),
    ("PUT","api/bids/[id]/accept","Accept a bid (selects contractor)","Client who owns job"),
    ("PUT","api/bids/[id]/withdraw","Withdraw own bid","Contractor who submitted"),
],'1A355E')

h2(doc,"7.5 Payments (Stripe Webhooks)")
mtable(doc,[
    ("Method","Endpoint","Description","Auth"),
    ("POST","api/payments/create-intent","Create Stripe payment intent for deposit","Authenticated client"),
    ("POST","api/payments/release","Release escrow to contractor","Authenticated client or auto"),
    ("POST","api/stripe/webhook","Stripe webhook handler (subscription events, payment events)","Stripe signature verified"),
    ("GET","api/payments/history","Client payment history","Authenticated client"),
    ("GET","api/payments/earnings","Contractor earnings history","Contractor only"),
],'106028')

h2(doc,"7.6 Admin")
mtable(doc,[
    ("Method","Endpoint","Description","Auth"),
    ("GET","api/admin/applications","List pending contractor applications","Admin only"),
    ("PUT","api/admin/applications/[id]/approve","Approve contractor application","Admin only"),
    ("PUT","api/admin/applications/[id]/reject","Reject contractor application","Admin only"),
    ("GET","api/admin/disputes","List open disputes","Admin only"),
    ("PUT","api/admin/disputes/[id]/resolve","Resolve dispute with ruling","Admin only"),
    ("GET","api/admin/analytics","Platform stats (revenue, users, jobs)","Admin only"),
],'1A355E')

# ═══ SECTION 8: HOW STERLING MAKES CHANGES ═══════════════════════
pb(doc)
h1(doc,"SECTION 8 -- HOW TO MAKE CHANGES (NON-TECHNICAL GUIDE)",GOLD)
body(doc,"This section explains how Sterling can make changes to the platform without being a full-time developer. "
     "The goal is to minimize the time Sterling needs to touch code.")

h2(doc,"8.1 Two Types of Changes")
mtable(doc,[
    ("Type","Examples","How Hard","How to Do It"),
    ("Content changes","Change pricing, update descriptions, add trade types, change email wording",
     "Easy -- no code","Edit the config files described below. No deployment needed for most."),
    ("Feature changes","Add a new page, change the matching logic, add a new payment flow",
     "Medium -- some code","Give the updated requirements to Claude Code. It will make the changes and show you a diff."),
    ("Major changes","New database tables, Stripe integration changes, authentication changes",
     "Complex -- full dev work","Schedule a focused Claude Code session with a detailed spec. Test on staging first."),
],'1A355E')

h2(doc,"8.2 Easy Change: Content and Config")
body(doc,"These items are stored in a central config file (lib/config.ts) so Sterling can change them without touching any other code:")
mtable(doc,[
    ("Item","How to Change"),
    ("Subscription price ($49/month)","Update STRIPE_PRICE_ID in Stripe dashboard, update display in config.ts"),
    ("Platform fee percentage (10%)","Change PLATFORM_FEE_PERCENT in lib/config.ts"),
    ("Trade types list","Add/remove items from TRADE_TYPES array in lib/config.ts"),
    ("Service area ZIP codes","Update DEFAULT_SERVICE_AREA in lib/config.ts"),
    ("Email templates","Edit files in emails/ folder -- they are written in plain React JSX"),
    ("Page text and headlines","Edit the text directly in the page component files in app/ folder"),
    ("Colors and branding","Edit tailwind.config.ts -- colors are defined in one place"),
],'106028')

h2(doc,"8.3 Making Changes with Claude Code")
body(doc,"This is Sterling's primary workflow for feature changes. Claude Code can make surgical changes to the codebase without breaking anything else.")
for step in [
    "1. DESCRIBE what you want changed in plain English: 'I want to add a background check requirement to the contractor application'",
    "2. OPEN the project folder in your terminal: cd C:\\Users\\Sterling\\.openclaw\\workspace\\BigPivot\\prototype",
    "3. RUN Claude Code: claude --permission-mode bypassPermissions --print 'Your description here'",
    "4. REVIEW the changes Claude Code made (it will show you a diff)",
    "5. TEST locally: npm run dev -- open localhost:3000 in your browser",
    "6. DEPLOY: git add . && git commit -m 'Add background check requirement' && git push",
    "7. Vercel automatically deploys the new version within 2 minutes",
]:
    bullet(doc, step)
doc.add_paragraph()

h2(doc,"8.4 Rollback -- If Something Breaks")
body(doc,"Because everything is in GitHub with version control, rolling back is always possible:")
for step in [
    "git log -- shows all recent changes with dates and descriptions",
    "git revert HEAD -- undoes the last change and creates a new commit",
    "git checkout [commit-hash] [filename] -- restore a specific file to a previous version",
    "In Vercel dashboard: click 'Deployments' and 'Redeploy' any previous version with one click",
]:
    bullet(doc, step)
doc.add_paragraph()

callout(doc,"STERLING'S RULE:",
    "Never edit files directly on the server or in Vercel. "
    "Always make changes locally, test with 'npm run dev', then push to GitHub. "
    "GitHub is the single source of truth. Vercel reads from GitHub automatically.",
    'FFF3CD', ORANGE, BLACK)

# ═══ SECTION 9: LAUNCH CHECKLIST ════════════════════════════════
pb(doc)
h1(doc,"SECTION 9 -- LAUNCH CHECKLIST")

h2(doc,"9.1 Pre-Launch -- Business")
mtable(doc,[
    ("Item","Owner","Done?"),
    ("Operating agreement signed (Regina and partner)","Both parties","[ ]"),
    ("Sedart Trades LLC verified in Michigan LARA","Regina","[ ]"),
    ("Business bank account opened for LLC","Regina","[ ]"),
    ("WBE application submitted (WBEC Great Lakes)","Regina","[ ]"),
    ("Detroit Business Opportunity Program (DBOP) application submitted","Regina","[ ]"),
    ("NEI grant application submitted","Solomon (Grant Agent)","[ ]"),
    ("JPMorgan Chase Detroit grant application submitted","Solomon (Grant Agent)","[ ]"),
    ("Contractor insurance requirement policy drafted (by attorney)","Attorney","[ ]"),
    ("Terms of Service and Privacy Policy drafted (by attorney)","Attorney","[ ]"),
],'106028')

h2(doc,"9.2 Pre-Launch -- Technical")
mtable(doc,[
    ("Item","Owner","Done?"),
    ("GitHub repository created and code pushed","Sterling/Developer","[ ]"),
    ("Supabase project created and database schema applied","Sterling/Developer","[ ]"),
    ("Stripe account created and Connect enabled","Sterling","[ ]"),
    ("Resend account created and domain verified","Sterling","[ ]"),
    ("All .env.local variables configured","Sterling/Developer","[ ]"),
    ("Vercel project connected to GitHub repo","Sterling","[ ]"),
    ("Custom domain purchased and configured","Sterling","[ ]"),
    ("Email domain verified (for professional outgoing email)","Sterling","[ ]"),
    ("All 6 pages tested on desktop and mobile","Sterling","[ ]"),
    ("Payment flow tested end-to-end in Stripe test mode","Sterling/Developer","[ ]"),
    ("Admin dashboard tested with sample contractor application","Regina/Sterling","[ ]"),
],'1A355E')

h2(doc,"9.3 First 30 Days After Launch")
mtable(doc,[
    ("Goal","Target","How"),
    ("Contractor applications","25 approved contractors","Run Big Pivot workshops. Direct outreach to known contractors."),
    ("Client job postings","10 jobs posted","UPrep campuses (Sterling's connection). Derrick's school district contacts."),
    ("First payment processed","1 job completed through platform","Manage closely. This proves the payment system works end-to-end."),
    ("Thompson Foundation meeting","1 intro meeting","Regina re-engages existing contact. Bring capability statement + pitch deck."),
    ("First government job posted","1 city or school district job","Derrick leverages DoIT/school district relationships."),
],'106028')

doc.save(OUT)
print(f"DONE: {OUT}")
