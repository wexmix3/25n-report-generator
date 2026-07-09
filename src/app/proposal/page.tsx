"use client";

export default function ProposalPage() {
  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

        * {
          box-sizing: border-box;
          margin: 0;
          padding: 0;
          -webkit-print-color-adjust: exact;
          print-color-adjust: exact;
        }

        body {
          font-family: 'Inter', -apple-system, sans-serif;
          background: #f1f5f9;
          color: #1e293b;
        }

        .print-btn {
          position: fixed;
          top: 24px;
          right: 24px;
          background: #0f2044;
          color: white;
          border: none;
          padding: 10px 20px;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          z-index: 100;
          font-family: inherit;
        }
        .print-btn:hover { background: #1a3461; }

        .page {
          width: 816px;
          min-height: 1056px;
          background: white;
          margin: 40px auto;
          padding: 0;
          box-shadow: 0 4px 24px rgba(0,0,0,0.12);
        }

        /* ── COVER HEADER ── */
        .cover {
          background: #0f2044;
          padding: 48px 56px 40px;
          color: white;
        }
        .cover-eyebrow {
          font-size: 11px;
          font-weight: 600;
          letter-spacing: 0.12em;
          text-transform: uppercase;
          color: #c9a84c;
          margin-bottom: 12px;
        }
        .cover-title {
          font-size: 28px;
          font-weight: 700;
          line-height: 1.2;
          margin-bottom: 8px;
        }
        .cover-subtitle {
          font-size: 15px;
          color: #94a3b8;
          font-weight: 400;
        }
        .cover-meta {
          display: flex;
          gap: 32px;
          margin-top: 28px;
          padding-top: 24px;
          border-top: 1px solid rgba(255,255,255,0.12);
          font-size: 12.5px;
          color: #94a3b8;
        }
        .cover-meta-item strong {
          display: block;
          color: white;
          font-weight: 500;
          margin-bottom: 2px;
        }

        /* ── BODY ── */
        .body {
          padding: 44px 56px 52px;
        }

        /* ── SECTIONS ── */
        .section {
          margin-bottom: 36px;
        }
        .section-label {
          font-size: 10px;
          font-weight: 700;
          letter-spacing: 0.1em;
          text-transform: uppercase;
          color: #c9a84c;
          margin-bottom: 6px;
        }
        .section-title {
          font-size: 16px;
          font-weight: 700;
          color: #0f2044;
          margin-bottom: 12px;
          padding-bottom: 10px;
          border-bottom: 2px solid #e2e8f0;
        }

        /* ── EXEC SUMMARY ── */
        .exec-box {
          background: #f8fafc;
          border-left: 4px solid #0f2044;
          padding: 18px 20px;
          border-radius: 0 6px 6px 0;
        }
        .exec-box p {
          font-size: 13.5px;
          line-height: 1.7;
          color: #334155;
        }

        /* ── DELIVERABLES ── */
        .deliverable-item {
          display: flex;
          gap: 14px;
          padding: 14px 0;
          border-bottom: 1px solid #f1f5f9;
        }
        .deliverable-item:last-child { border-bottom: none; }
        .deliverable-num {
          width: 28px;
          height: 28px;
          border-radius: 50%;
          background: #0f2044;
          color: white;
          font-size: 12px;
          font-weight: 700;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
          margin-top: 1px;
        }
        .deliverable-content strong {
          font-size: 13.5px;
          font-weight: 600;
          color: #0f2044;
          display: block;
          margin-bottom: 3px;
        }
        .deliverable-content p {
          font-size: 12.5px;
          line-height: 1.6;
          color: #64748b;
        }
        .deliverable-badge {
          display: inline-block;
          font-size: 10px;
          font-weight: 600;
          padding: 2px 8px;
          border-radius: 12px;
          margin-left: 8px;
          vertical-align: middle;
        }
        .badge-done { background: #dcfce7; color: #16a34a; }
        .badge-wip  { background: #fef9c3; color: #a16207; }

        /* ── SERVICES GRID ── */
        .services-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 16px;
          margin-top: 4px;
        }
        .service-card {
          border: 1.5px solid #e2e8f0;
          border-radius: 8px;
          padding: 18px 20px;
        }
        .service-card-header {
          display: flex;
          align-items: center;
          justify-content: space-between;
          margin-bottom: 12px;
        }
        .service-card-name {
          font-size: 13px;
          font-weight: 700;
          color: #0f2044;
        }
        .service-card-price {
          font-size: 16px;
          font-weight: 700;
          color: #c9a84c;
        }
        .service-card-price span {
          font-size: 11px;
          font-weight: 400;
          color: #94a3b8;
        }
        .service-list {
          list-style: none;
        }
        .service-list li {
          font-size: 12px;
          color: #475569;
          padding: 4px 0;
          padding-left: 16px;
          position: relative;
          line-height: 1.4;
        }
        .service-list li::before {
          content: '✓';
          position: absolute;
          left: 0;
          color: #16a34a;
          font-size: 11px;
          font-weight: 700;
        }

        /* ── ROI ── */
        .roi-grid {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 12px;
          margin-bottom: 16px;
        }
        .roi-card {
          background: #f8fafc;
          border: 1px solid #e2e8f0;
          border-radius: 8px;
          padding: 14px 16px;
          text-align: center;
        }
        .roi-card-number {
          font-size: 22px;
          font-weight: 700;
          color: #0f2044;
          margin-bottom: 4px;
        }
        .roi-card-label {
          font-size: 11px;
          color: #64748b;
          line-height: 1.3;
        }
        .roi-note {
          font-size: 12.5px;
          color: #475569;
          line-height: 1.65;
        }
        .roi-note strong {
          color: #0f2044;
        }

        /* ── HEAD OF AI BOX ── */
        .hoa-box {
          background: #fffbeb;
          border: 1.5px solid #fde68a;
          border-radius: 8px;
          padding: 18px 20px;
        }
        .hoa-box p {
          font-size: 13px;
          color: #78350f;
          line-height: 1.65;
        }
        .hoa-analogy {
          margin-top: 10px;
          padding-top: 10px;
          border-top: 1px solid #fde68a;
          font-size: 12px;
          color: #92400e;
          font-style: italic;
        }

        /* ── WHY ME ── */
        .why-me-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 10px;
        }
        .why-me-item {
          display: flex;
          gap: 10px;
          align-items: flex-start;
          padding: 12px 14px;
          background: #f8fafc;
          border-radius: 6px;
        }
        .why-me-icon {
          font-size: 16px;
          flex-shrink: 0;
          margin-top: 1px;
        }
        .why-me-item strong {
          font-size: 12.5px;
          font-weight: 600;
          color: #0f2044;
          display: block;
          margin-bottom: 2px;
        }
        .why-me-item p {
          font-size: 11.5px;
          color: #64748b;
          line-height: 1.5;
        }

        /* ── INVESTMENT TABLE ── */
        .investment-table {
          width: 100%;
          border-collapse: collapse;
          margin-bottom: 14px;
        }
        .investment-table th {
          text-align: left;
          font-size: 10.5px;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.06em;
          color: #64748b;
          padding: 8px 12px;
          background: #f8fafc;
          border-bottom: 1px solid #e2e8f0;
        }
        .investment-table td {
          padding: 12px 12px;
          font-size: 13px;
          color: #334155;
          border-bottom: 1px solid #f1f5f9;
        }
        .investment-table td:last-child {
          font-weight: 600;
          color: #0f2044;
          text-align: right;
        }
        .investment-table th:last-child {
          text-align: right;
        }
        .investment-total td {
          font-weight: 700;
          color: #0f2044;
          font-size: 14px;
          padding-top: 14px;
          border-bottom: none;
        }
        .investment-total td:last-child {
          color: #c9a84c;
          font-size: 16px;
        }
        .pilot-box {
          background: #0f2044;
          color: white;
          padding: 16px 20px;
          border-radius: 8px;
          font-size: 13px;
          line-height: 1.6;
        }
        .pilot-box strong {
          color: #c9a84c;
        }

        /* ── FOOTER ── */
        .footer {
          padding: 20px 56px;
          border-top: 1px solid #e2e8f0;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .footer-name {
          font-size: 12px;
          font-weight: 600;
          color: #0f2044;
        }
        .footer-contact {
          font-size: 11.5px;
          color: #64748b;
        }

        /* ── PRINT ── */
        @media print {
          body { background: white; }
          .print-btn { display: none !important; }
          .page {
            margin: 0;
            box-shadow: none;
            width: 100%;
            min-height: unset;
          }
          .cover {
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
          }
          .section { page-break-inside: avoid; }
          .deliverable-item { page-break-inside: avoid; }
          .services-grid { page-break-inside: avoid; }
          .roi-grid { page-break-inside: avoid; }
          .hoa-box { page-break-inside: avoid; }
          .why-me-grid { page-break-inside: avoid; }
          .investment-table { page-break-inside: avoid; }
          .pilot-box { page-break-inside: avoid; }
          @page {
            size: letter;
            margin: 0.5in;
          }
        }
      `}</style>

      <button className="print-btn" onClick={() => window.print()}>
        Download PDF
      </button>

      <div className="page">

        {/* ── HEADER ── */}
        <div className="cover">
          <div className="cover-eyebrow">Consulting Proposal</div>
          <div className="cover-title">AI Tools &amp; Automation Services</div>
          <div className="cover-subtitle">Presented to 25N Coworking Executive Board</div>
          <div className="cover-meta">
            <div className="cover-meta-item">
              <strong>Prepared by</strong>
              Max Wexley
            </div>
            <div className="cover-meta-item">
              <strong>Date</strong>
              May 2026
            </div>
            <div className="cover-meta-item">
              <strong>Contact</strong>
              maxwexley@wexadvisory.com
            </div>
            <div className="cover-meta-item">
              <strong>Engagement Type</strong>
              Subscription + Hourly
            </div>
          </div>
        </div>

        <div className="body">

          {/* ── EXEC SUMMARY ── */}
          <div className="section">
            <div className="section-label">Overview</div>
            <div className="section-title">Executive Summary</div>
            <div className="exec-box">
              <p>
                25N already has evidence this works. Over the past several months, I have delivered a PowerPoint revamp,
                a fully functional Competitive Analysis tool, and an ongoing Finance Dashboard modernization — all at no cost.
                This proposal formalizes that engagement. For <strong>$300/month</strong> your team gets unlimited access to all
                deployed AI tools and automations. For <strong>$100/hour</strong>, you get direct, on-demand development of new
                automations and integrations. No agency overhead. No onboarding delay. A builder who already knows your business.
              </p>
            </div>
          </div>

          {/* ── WHAT'S ALREADY WORKING ── */}
          <div className="section">
            <div className="section-label">Track Record</div>
            <div className="section-title">What's Already Working</div>

            <div className="deliverable-item">
              <div className="deliverable-num">1</div>
              <div className="deliverable-content">
                <strong>
                  PowerPoint Deck Revamp
                  <span className="deliverable-badge badge-done">Delivered</span>
                </strong>
                <p>
                  Applied financial presentation standards from daily work at BBVA to sharpen a 25N board-ready deck.
                  Fast turnaround, finance-fluent output — the first demonstration of judgment, not just technical skill.
                </p>
              </div>
            </div>

            <div className="deliverable-item">
              <div className="deliverable-num">2</div>
              <div className="deliverable-content">
                <strong>
                  Competitive Analysis Tool
                  <span className="deliverable-badge badge-done">Delivered — Free</span>
                </strong>
                <p>
                  Built and shared a custom competitive intelligence tool at no cost. Research that previously took hours
                  of manual work now runs on demand. This was a proof of capability, not a sales pitch.
                </p>
              </div>
            </div>

            <div className="deliverable-item">
              <div className="deliverable-num">3</div>
              <div className="deliverable-content">
                <strong>
                  Finance Dashboard Modernization
                  <span className="deliverable-badge badge-wip">In Progress</span>
                </strong>
                <p>
                  Currently architecting three options for eliminating manual month-end reporting: VBA/Macros, Python automation,
                  and a PDF report generation pipeline. The CFO chooses the approach — I build it.
                  This proposal is asking the board to formalize what is already underway.
                </p>
              </div>
            </div>
          </div>

          {/* ── SERVICES ── */}
          <div className="section">
            <div className="section-label">Services</div>
            <div className="section-title">What You Get Going Forward</div>
            <div className="services-grid">
              <div className="service-card">
                <div className="service-card-header">
                  <div className="service-card-name">Monthly Tool Subscription</div>
                  <div className="service-card-price">$300<span>/mo</span></div>
                </div>
                <ul className="service-list">
                  <li>Access to all deployed AI &amp; automation tools</li>
                  <li>Hosting and uptime maintenance</li>
                  <li>Bug fixes and minor updates</li>
                  <li>Monthly strategy check-in call</li>
                  <li>Priority response on critical issues</li>
                </ul>
              </div>
              <div className="service-card">
                <div className="service-card-header">
                  <div className="service-card-name">Project &amp; Advisory Work</div>
                  <div className="service-card-price">$100<span>/hr</span></div>
                </div>
                <ul className="service-list">
                  <li>New automation builds on request</li>
                  <li>System integrations and data pipelines</li>
                  <li>CFO and board reporting tools</li>
                  <li>Operational workflow redesign</li>
                  <li>Scoped flat-fee projects available</li>
                </ul>
              </div>
            </div>
          </div>

          {/* ── ROI ── */}
          <div className="section">
            <div className="section-label">Business Case</div>
            <div className="section-title">Return on Investment</div>
            <div className="roi-grid">
              <div className="roi-card">
                <div className="roi-card-number">15–20 hrs</div>
                <div className="roi-card-label">staff time reclaimed monthly across finance &amp; ops</div>
              </div>
              <div className="roi-card">
                <div className="roi-card-number">$525–900</div>
                <div className="roi-card-label">monthly labor savings at blended staff rate</div>
              </div>
              <div className="roi-card">
                <div className="roi-card-number">$300/mo</div>
                <div className="roi-card-label">subscription cost — less than a single SaaS seat</div>
              </div>
            </div>
            <p className="roi-note">
              The subscription pays for itself in the first week of the month. Manual reporting errors
              carry their own cost — one restatement or a bad number in a board deck is worth more than a year
              of this engagement. <strong>Each new automation compounds the return:</strong> Month 6 ROI is
              materially better than Month 1 as the tool suite expands. A custom software agency would charge
              $15,000–$50,000 to build what already exists here, plus $2,000–$5,000/month to maintain it.
            </p>
          </div>

          {/* ── HEAD OF AI ── */}
          <div className="section">
            <div className="section-label">Addressing a Common Question</div>
            <div className="section-title">What About the Head of AI?</div>
            <div className="hoa-box">
              <p>
                The Head of AI owns the <strong>strategy</strong>: vendor evaluation, data governance, and aligning AI
                with company direction. That is a full-time organizational role.
                I own the <strong>execution</strong>: custom tools built for 25N's exact workflows — not off-the-shelf
                software adapted to fit. These are additive, not competing. The tools I build give your Head of AI a
                working foundation on day one instead of a greenfield. They set the direction; I build the infrastructure
                that makes that direction possible.
              </p>
              <div className="hoa-analogy">
                A CFO doesn't eliminate the need for an analyst. Strategy and execution live at different layers.
              </div>
            </div>
          </div>

          {/* ── WHY ME ── */}
          <div className="section">
            <div className="section-label">Qualifications</div>
            <div className="section-title">Why Me</div>
            <div className="why-me-grid">
              <div className="why-me-item">
                <div className="why-me-icon">📊</div>
                <div>
                  <strong>Finance Background, Not Just Code</strong>
                  <p>Analyst at BBVA. The dashboard wasn't built by a developer guessing what a CFO needs — it was built by someone who produces this work every day.</p>
                </div>
              </div>
              <div className="why-me-item">
                <div className="why-me-icon">✅</div>
                <div>
                  <strong>Three Deliverables Already in Hand</strong>
                  <p>PPT revamp, competitive analysis tool, active dashboard project. Age didn't produce those — judgment and execution did.</p>
                </div>
              </div>
              <div className="why-me-item">
                <div className="why-me-icon">🚀</div>
                <div>
                  <strong>Shipped Products, Real Users</strong>
                  <p>Chapterly (live iOS/web app), Personal Finance Dashboard, CourtBook NYC. These aren't experiments — they are production applications.</p>
                </div>
              </div>
              <div className="why-me-item">
                <div className="why-me-icon">🤝</div>
                <div>
                  <strong>Skin in the Game</strong>
                  <p>Monthly subscription model: if I stop delivering value, you cancel. The incentive structure is perfectly aligned with your interests.</p>
                </div>
              </div>
              <div className="why-me-item">
                <div className="why-me-icon">⚡</div>
                <div>
                  <strong>Zero Overhead</strong>
                  <p>No benefits, no desk, no onboarding, no account manager in the middle. Direct access to the builder who already knows your business.</p>
                </div>
              </div>
              <div className="why-me-item">
                <div className="why-me-icon">🎯</div>
                <div>
                  <strong>Already Invested in 25N</strong>
                  <p>The competitive analysis and PPT work were delivered for free. That's not a sales tactic — it's a demonstration that I take this seriously.</p>
                </div>
              </div>
            </div>
          </div>

          {/* ── INVESTMENT ── */}
          <div className="section">
            <div className="section-label">Pricing</div>
            <div className="section-title">Investment</div>
            <table className="investment-table">
              <thead>
                <tr>
                  <th>Service</th>
                  <th>Description</th>
                  <th>Amount</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>Monthly Tool Subscription</td>
                  <td>Full access to AI tools, automation suite, hosting &amp; maintenance</td>
                  <td>$300/mo</td>
                </tr>
                <tr>
                  <td>Project &amp; Advisory Hours</td>
                  <td>New builds, integrations, and strategic projects (est. 5–10 hrs/mo)</td>
                  <td>$100/hr</td>
                </tr>
                <tr className="investment-total">
                  <td>Estimated Monthly Total</td>
                  <td></td>
                  <td>$800–$1,300</td>
                </tr>
              </tbody>
            </table>
            <div className="pilot-box">
              <strong>Recommended: 3-Month Pilot</strong>&nbsp; Approve a 90-day pilot engagement.
              At the end of Month 3, we review ROI together — time saved, tools deployed, value delivered —
              and decide what to expand. No long-term commitment required to get started.
            </div>
          </div>

        </div>

        {/* ── FOOTER ── */}
        <div className="footer">
          <div>
            <div className="footer-name">Max Wexley — Consulting Proposal</div>
            <div className="footer-contact">maxwexley@wexadvisory.com · wexadvisory.com</div>
          </div>
          <div className="footer-contact">Confidential · May 2026</div>
        </div>
      </div>
    </>
  );
}
