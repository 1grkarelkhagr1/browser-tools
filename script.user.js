// ==UserScript==
// @name         Weekly Staff
// @namespace    browser-tools
// @version      1.2
// @description  Reports
// @match        https://www.scheduleit.com/cloud/std2/*
// @match        https://www.scheduleit.com/cloud/std2/reports_create.php*
// @updateURL    https://raw.githubusercontent.com/1grkarelkhagr1/browser-tools/main/script.user.js
// @downloadURL  https://raw.githubusercontent.com/1grkarelkhagr1/browser-tools/main/script.user.js
// @run-at       document-end
// @grant        GM_registerMenuCommand
// @grant        GM_openInTab
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addValueChangeListener
// @grant        GM_removeValueChangeListener
// ==/UserScript==

(() => {
  /* ================= CONFIG ================= =*/
  const MOBILE_SCALE = 0.6;
  const FILE_LABEL = "Weekly_Staff";
  const EMAIL_FILE_LABEL = "Weekly_Staff_Email_List";

  const THIS_WEEK_MENU_ITEM_NAME = "This Week (Fri)";
  const NEXT_WEEK_MENU_ITEM_NAME = "Next Week (Sat, Sun)";
  const FIXED_DATE_RANGE_MENU_ITEM_NAME = "Fixed Date Range";

  /*
    Saturday-based working week day numbers:

    1 = Saturday
    2 = Sunday
    3 = Monday
    4 = Tuesday
    5 = Wednesday
    6 = Thursday
    7 = Friday

    All changeable filter_from_id values are entered here.

    ## Email List automatically uses the same filter_from_id as Person. ##
  */
  const FILTER_IDS = {
    fixedDateRange: {
      job: "265",
      person: "266",
      accommodation: "267"
    },

    thisWeek: {
      1: { // Saturday
        job: "273",
        person: "274",
        accommodation: "272"
      },
      2: { // Sunday
        job: "276",
        person: "277",
        accommodation: "275"
      },
      /*3: { // Monday
        job: "ENTER_THIS_WEEK_MONDAY_JOB_ID",
        person: "ENTER_THIS_WEEK_MONDAY_PERSON_ID",
        accommodation: "ENTER_THIS_WEEK_MONDAY_ACCOMMODATION_ID"
      },
      4: { // Tuesday
        job: "ENTER_THIS_WEEK_TUESDAY_JOB_ID",
        person: "ENTER_THIS_WEEK_TUESDAY_PERSON_ID",
        accommodation: "ENTER_THIS_WEEK_TUESDAY_ACCOMMODATION_ID"
      },
      5: { // Wednesday
        job: "ENTER_THIS_WEEK_WEDNESDAY_JOB_ID",
        person: "ENTER_THIS_WEEK_WEDNESDAY_PERSON_ID",
        accommodation: "ENTER_THIS_WEEK_WEDNESDAY_ACCOMMODATION_ID"
      },
      6: { // Thursday
        job: "ENTER_THIS_WEEK_THURSDAY_JOB_ID",
        person: "ENTER_THIS_WEEK_THURSDAY_PERSON_ID",
        accommodation: "ENTER_THIS_WEEK_THURSDAY_ACCOMMODATION_ID"
      },
      7: { // Friday
        job: "ENTER_THIS_WEEK_FRIDAY_JOB_ID",
        person: "ENTER_THIS_WEEK_FRIDAY_PERSON_ID",
        accommodation: "ENTER_THIS_WEEK_FRIDAY_ACCOMMODATION_ID"
      }*/
    },

    nextWeek: {
      /*1: { // Saturday
        job: "ENTER_NEXT_WEEK_SATURDAY_JOB_ID",
        person: "ENTER_NEXT_WEEK_SATURDAY_PERSON_ID",
        accommodation: "ENTER_NEXT_WEEK_SATURDAY_ACCOMMODATION_ID"
      },
      2: { // Sunday
        job: "ENTER_NEXT_WEEK_SUNDAY_JOB_ID",
        person: "ENTER_NEXT_WEEK_SUNDAY_PERSON_ID",
        accommodation: "ENTER_NEXT_WEEK_SUNDAY_ACCOMMODATION_ID"
      },
      3: { // Monday
        job: "ENTER_NEXT_WEEK_MONDAY_JOB_ID",
        person: "ENTER_NEXT_WEEK_MONDAY_PERSON_ID",
        accommodation: "ENTER_NEXT_WEEK_MONDAY_ACCOMMODATION_ID"
      },
      4: { // Tuesday
        job: "ENTER_NEXT_WEEK_TUESDAY_JOB_ID",
        person: "ENTER_NEXT_WEEK_TUESDAY_PERSON_ID",
        accommodation: "ENTER_NEXT_WEEK_TUESDAY_ACCOMMODATION_ID"
      },
      5: { // Wednesday
        job: "ENTER_NEXT_WEEK_WEDNESDAY_JOB_ID",
        person: "ENTER_NEXT_WEEK_WEDNESDAY_PERSON_ID",
        accommodation: "ENTER_NEXT_WEEK_WEDNESDAY_ACCOMMODATION_ID"
      },
      6: { // Thursday
        job: "ENTER_NEXT_WEEK_THURSDAY_JOB_ID",
        person: "ENTER_NEXT_WEEK_THURSDAY_PERSON_ID",
        accommodation: "ENTER_NEXT_WEEK_THURSDAY_ACCOMMODATION_ID"
      },*/
      7: { // Friday
        job: "270",
        person: "271",
        accommodation: "269"
      }
    }
  };

  const REPORTS = [
    {
      key: "job",
      label: "Job",
      report_template: "2131",
      report_splitby: "1",
      r_sortby: "1",
      e_sortby: "1",
      mode: "tab"
    },
    {
      key: "person",
      label: "Person",
      report_template: "2130",
      report_splitby: "1",
      r_sortby: "1",
      e_sortby: "1",
      mode: "tab"
    },
    {
      key: "accommodation",
      label: "Accommodation",
      report_template: "2132",
      report_splitby: "1",
      r_sortby: "1",
      e_sortby: "1",
      mode: "tab"
    },
    {
      key: "email",
      label: "Email List",
      report_template: "1133",
      report_splitby: "1",
      r_sortby: "1",
      e_sortby: "1",
      mode: "email"
    }
  ];

  function getSaturdayBasedDay() {
    return ((new Date().getDay() + 1) % 7) + 1;
  }

  function getReportsWithFilterIds(filterIds) {
    return REPORTS.map(report => {
      const filterFromId =
        report.key === "email"
          ? filterIds.person
          : filterIds[report.key];

      if (!filterFromId || filterFromId.startsWith("ENTER_")) {
        throw new Error(`Missing filter_from_id for ${report.label}`);
      }

      return {
        ...report,
        filter_from_id: filterFromId
      };
    });
  }

  function getThisWeekReports() {
    const day = getSaturdayBasedDay();
    const filterIds = FILTER_IDS.thisWeek[day];

    if (!filterIds) {
      throw new Error(
        "This Week is not configured for today's day."
      );
    }

    return getReportsWithFilterIds(filterIds);
  }

  function getNextWeekReports() {
    const day = getSaturdayBasedDay();
    const filterIds = FILTER_IDS.nextWeek[day];

    if (!filterIds) {
      throw new Error(
        "Next Week is not configured for today's day."
      );
    }

    return getReportsWithFilterIds(filterIds);
  }

  function getFixedDateRangeReports() {
    return getReportsWithFilterIds(FILTER_IDS.fixedDateRange);
  }

  /* ================= CSS ================= */
  const CAPTURE_CSS = `
html, body {
  margin: 0 !important;
  padding: 0 !important;
  background: #fff !important;
  overflow-x: hidden !important;
}

body {
  font-family: "Roboto", Arial, Helvetica, sans-serif !important;
}

html, body, table, th, td, p, span, strong, label, a, button {
  -webkit-text-size-adjust: 100% !important;
  text-size-adjust: 100% !important;
}

* { box-sizing: border-box !important; }

:root {
  --tabs-height: 46px;
  --jumpbar-height: 50px;
  --jump-offset: calc(var(--tabs-height) + var(--jumpbar-height) + 4px);
}

#top-tabs {
  position: fixed;
  top: 0; left: 0; right: 0;
  z-index: 100000;
  background: #000;
  height: var(--tabs-height);
  display: flex;
  align-items: stretch;
  gap: 0;
  padding: 0;
  overflow-x: auto;
  overflow-y: hidden;
  white-space: nowrap;
  -webkit-overflow-scrolling: touch;
}

#top-tabs .tm-tab-link {
  display: flex;
  flex: 1;
  min-width: 100px;
  align-items: center;
  justify-content: center;
  height: 100%;
  margin: 0;
  padding: 0 12px;
  text-align: center;
  border-radius: 0;
  background: #000;
  color: #fff;
  text-decoration: none !important;
  border-top: 2px solid #ccc;
  border-bottom: 2px solid #ccc;
  border-left: 2px solid #ccc;
  border-right: 0;
  font-weight: 700;
  cursor: pointer;
}

#top-tabs .tm-tab-link:last-child {
  border-right: 2px solid #ccc;
}

body {
  padding-top: calc(var(--tabs-height) + var(--jumpbar-height)) !important;
}

.tm-panel {
  display: none;
}

.tm-panel-default {
  display: block;
}

body:has(.tm-panel:target) .tm-panel,
body:has(.tm-anchor:target) .tm-panel {
  display: none;
}

body:has(.tm-panel:target) .tm-panel-default,
body:has(.tm-anchor:target) .tm-panel-default {
  display: none;
}

#tm-panel-0:target,
body:has(#tm-panel-0 .tm-anchor:target) #tm-panel-0,
#tm-panel-1:target,
body:has(#tm-panel-1 .tm-anchor:target) #tm-panel-1,
#tm-panel-2:target,
body:has(#tm-panel-2 .tm-anchor:target) #tm-panel-2,
#tm-panel-3:target,
body:has(#tm-panel-3 .tm-anchor:target) #tm-panel-3,
#tm-panel-4:target,
body:has(#tm-panel-4 .tm-anchor:target) #tm-panel-4,
#tm-panel-5:target,
body:has(#tm-panel-5 .tm-anchor:target) #tm-panel-5,
#tm-panel-6:target,
body:has(#tm-panel-6 .tm-anchor:target) #tm-panel-6,
#tm-panel-7:target,
body:has(#tm-panel-7 .tm-anchor:target) #tm-panel-7 {
  display: block;
}

.tm-panel .tm-jumpbar {
  position: fixed;
  top: var(--tabs-height);
  left: 0;
  right: 0;
  z-index: 99999;
  background: #FFA500 !important;
  border-bottom: 1px solid #ccc;
  height: var(--jumpbar-height);
  display: flex;
  align-items: center;
  padding: 0 8px;
}

.tm-panel .jb-inner {
  display: flex;
  gap: 10px;
  align-items: center;
  width: 100%;
}

.tm-panel .jb-label {
  font-weight: 700;
  white-space: nowrap;
}

.tm-panel .jb-scroll {
  flex: 1;
  overflow-x: auto;
  white-space: nowrap;
  -webkit-overflow-scrolling: touch;
}

.tm-panel .jb-scroll a {
  display: inline-block;
  margin-right: 8px;
  padding: 5px 9px;
  border-radius: 999px;
  background: rgba(255,255,255,0.92);
  border: 1px solid rgba(0,0,0,0.25);
  color: #000;
  text-decoration: none;
  font-weight: 600;
}

body:not(:has(.tm-panel:target)):not(:has(.tm-anchor:target)) #top-tabs a[href="#tm-panel-0"],
body:has(#tm-panel-0:target) #top-tabs a[href="#tm-panel-0"],
body:has(#tm-panel-0 .tm-anchor:target) #top-tabs a[href="#tm-panel-0"],
body:has(#tm-panel-1:target) #top-tabs a[href="#tm-panel-1"],
body:has(#tm-panel-1 .tm-anchor:target) #top-tabs a[href="#tm-panel-1"],
body:has(#tm-panel-2:target) #top-tabs a[href="#tm-panel-2"],
body:has(#tm-panel-2 .tm-anchor:target) #top-tabs a[href="#tm-panel-2"],
body:has(#tm-panel-3:target) #top-tabs a[href="#tm-panel-3"],
body:has(#tm-panel-3 .tm-anchor:target) #top-tabs a[href="#tm-panel-3"],
body:has(#tm-panel-4:target) #top-tabs a[href="#tm-panel-4"],
body:has(#tm-panel-4 .tm-anchor:target) #top-tabs a[href="#tm-panel-4"],
body:has(#tm-panel-5:target) #top-tabs a[href="#tm-panel-5"],
body:has(#tm-panel-5 .tm-anchor:target) #top-tabs a[href="#tm-panel-5"],
body:has(#tm-panel-6:target) #top-tabs a[href="#tm-panel-6"],
body:has(#tm-panel-6 .tm-anchor:target) #top-tabs a[href="#tm-panel-6"],
body:has(#tm-panel-7:target) #top-tabs a[href="#tm-panel-7"],
body:has(#tm-panel-7 .tm-anchor:target) #top-tabs a[href="#tm-panel-7"] {
  color: #FFA500;
}

#top-tabs,
#top-tabs *,
.tm-jumpbar,
.tm-jumpbar * {
  font-size: 1em;
}

@media (max-width: 768px) {
  #top-tabs,
  #top-tabs *,
  .tm-jumpbar,
  .tm-jumpbar * {
    font-size: 0.92em !important;
    line-height: 1.05 !important;
  }

  #top-tabs {
    padding: 4px 6px;
    height: 40px;
    gap: 6px;
  }

  .tm-panel .tm-jumpbar {
    padding: 0 6px;
  }

  :root {
    --tabs-height: 40px;
    --jumpbar-height: 40px;
  }

  #top-tabs .tm-tab-link {
    padding: 8px 10px !important;
    border-radius: 999px !important;
    border: 2px solid #ccc !important;
    margin-right: 0 !important;
  }

  .tm-panel .jb-scroll a {
    padding: 8px 10px !important;
    border-radius: 999px !important;
  }
}

.tm-report-wrap {
  margin: 0 !important;
  padding: 0 !important;
}

.tm-report-wrap #reportbody {
  margin: 0 !important;
  padding: 6px !important;
  background: #fff !important;
  border-radius: 0 !important;
  overflow-x: auto !important;
}

.tm-report-wrap #globalMultiSelected,
.tm-report-wrap #date_filter_string,
.tm-report-wrap .hidden-print {
  display: none !important;
}

.tm-report-wrap p {
  margin: 0 !important;
  padding: 0 !important;
}

.tm-report-wrap h1,
.tm-report-wrap h2,
.tm-report-wrap h3 {
  margin-top: 0 !important;
  margin-bottom: 0 !important;
  padding: 0 !important;
}

.tm-report-wrap p:empty {
  display: none !important;
  margin: 0 !important;
  padding: 0 !important;
}

.tm-report-wrap p:has(> br:only-child) {
  display: none !important;
}

.tm-report-wrap p > span > strong {
  font-size: 20px !important;
  font-weight: 700 !important;
}

.tm-report-wrap p > span {
  font-size: 20px !important;
}

@media (hover: none) and (pointer: coarse) {
  .tm-report-wrap p > span > strong,
  .tm-report-wrap p > span {
    font-size: 22px !important;
  }
}

table { max-width: 100% !important; }

.tm-anchor {
  display: block;
  height: 0;
  position: relative;
  top: calc(-1 * var(--jump-offset));
}

@media (hover: none) and (pointer: coarse) {
  .tm-anchor {
    top: calc(-1 * var(--jump-offset) / var(--mobile-scale));
  }
}

.tm-scale-wrap {
  zoom: 1;
  transform: none;
  transform-origin: top left;
}

@media (hover: none) and (pointer: coarse) {
  .tm-scale-wrap {
    zoom: var(--mobile-scale);
  }
}

@supports (-webkit-touch-callout: none) {
  @media (hover: none) and (pointer: coarse) {
    .tm-scale-wrap {
      zoom: 1 !important;
      transform: scale(var(--mobile-scale)) !important;
      width: calc(100% / var(--mobile-scale)) !important;
    }
  }
}

html { scroll-behavior: smooth; }

.tm-report-wrap .res_color_inner {
  padding-top: 2px !important;
  padding-bottom: 2px !important;
}

.tm-report-wrap a.tm-row-link,
.tm-report-wrap a.tm-row-link:visited {
  display: block !important;
  color: inherit !important;
  text-decoration: none !important;
  cursor: pointer !important;
}

`;

  /* ================= HELPERS ================= */
  const RUN = "tm_run";
  const IDX = "tm_idx";
  const KEY = id => `scheduleit:${id}`;

  const ymd = d => {
    const dt = d || new Date();
    return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
  };

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function cleanFilename(s) {
    return String(s)
      .replace(/[\\/:*?"<>|]/g, "")
      .replace(/\s+/g, "_")
      .trim();
  }

  function normalizeName(s) {
    return String(s || "")
      .replace(/^>\s*/, "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  function getReportMeta(doc) {
    const dateEl = doc.querySelector("#reportbody h1 + p span");
    const rawDate = dateEl ? dateEl.textContent.trim() : "";

    const match = rawDate.match(/(\d{2})\/(\d{2})\/(\d{4})/);

    let isoDate = "";
    if (match) {
      const [, d, m, y] = match;
      isoDate = `${y}-${m}-${d}`;
    }

    return { date: isoDate };
  }

  function buildUrl(date, cfg) {
    const p = new URLSearchParams({
      v_date: date,
      v_days: "14",
      report_template: cfg.report_template,
      report_splitby: cfg.report_splitby || "1",
      r_sortby: cfg.r_sortby || "1",
      e_sortby: cfg.e_sortby || "1",
      datetime_format: "dd/mm24",
      filter_from_id: cfg.filter_from_id
    });
    return `https://www.scheduleit.com/cloud/std2/reports_create.php?${p}`;
  }

  function waitForReportBody(timeout = 15000) {
    return new Promise((resolve, reject) => {
      const start = Date.now();

      function check() {
        const rb = document.querySelector("#reportbody");
        if (rb && rb.textContent.trim()) {
          resolve(rb);
          return;
        }

        if (Date.now() - start > timeout) {
          reject(new Error("Timed out waiting for #reportbody"));
          return;
        }

        setTimeout(check, 250);
      }

      check();
    });
  }

  function downloadFile(filename, content, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";

    document.body.appendChild(a);
    a.click();

    setTimeout(() => {
      a.remove();
      URL.revokeObjectURL(url);
    }, 2000);
  }

  function downloadHtml(filename, html) {
    downloadFile(filename, html, "text/html");
  }

  function downloadTxt(filename, text) {
    downloadFile(filename, text, "text/plain");
  }

  function extract(doc, panelIndex) {
    const rb = doc.querySelector("#reportbody");
    if (!rb) throw new Error("No report body");

    const clone = rb.cloneNode(true);
    clone.querySelectorAll("script, iframe, noscript").forEach(n => n.remove());

    // Detect "(No Resource Found)" and remove everything from that point onwards
    const allStrongs = Array.from(clone.querySelectorAll("p > span > strong"));

    const noResStrong = allStrongs.find(s =>
      s.textContent.includes("(No Resource Found)")
    );

    if (noResStrong) {
      const p = noResStrong.closest("p");
      if (p && p.parentNode) {
        let node = p;
        while (node) {
          const next = node.nextSibling;
          node.remove();
          node = next;
        }
      }
    }

    const items = [];
    const seen = new Set();
    let i = 0;

    clone.querySelectorAll("p > span > strong").forEach(strong => {
      const raw = strong.textContent.trim();
      const name = raw.startsWith("> ") ? raw.slice(2) : raw;
      if (!name || name.includes("(No Resource Found)")) return;

      const k = normalizeName(name);
      if (seen.has(k)) return;
      seen.add(k);

      const p = strong.closest("p");
      if (!p || !p.parentNode) return;

      const id = `tm-panel-${panelIndex}__name-${String(i++).padStart(4, "0")}`;

      const a = clone.ownerDocument.createElement("span");
      a.className = "tm-anchor";
      a.id = id;
      p.parentNode.insertBefore(a, p);

      items.push({ name, key: k, id, panelIndex });
    });

    const links = items.map(it =>
      `<a href="#${it.id}">${escapeHtml(it.name)}</a>`
    ).join("");

    return {
      items,
      jumpbar: `
<div class="tm-jumpbar">
  <div class="jb-inner">
    <div class="jb-label">Jump:</div>
    <div class="jb-scroll">${links}</div>
  </div>
</div>`,
      report: clone.outerHTML
    };
  }

  function extractEmails(doc) {
    const rb = doc.querySelector("#reportbody");
    if (!rb) throw new Error("No report body");

    const toEmails = [];
    const ccEmails = [];

    const toSeen = new Set();
    const ccSeen = new Set();

    let section = "";

    const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig;

    const lines = Array.from(rb.querySelectorAll("p"))
      .map(p => (p.textContent || "").trim())
      .filter(Boolean);

    for (const line of lines) {
      if (/^to\s*:?$/i.test(line)) {
        section = "to";
        continue;
      }

      if (/^cc\s*:?$/i.test(line)) {
        section = "cc";
        continue;
      }

      const matches = line.match(emailRegex);
      if (!matches) continue;

      for (const email of matches) {
        const cleanEmail = email.toLowerCase();

        if (section === "to") {
          if (!toSeen.has(cleanEmail)) {
            toSeen.add(cleanEmail);
            toEmails.push(cleanEmail);
          }
        } else if (section === "cc") {
          if (!ccSeen.has(cleanEmail)) {
            ccSeen.add(cleanEmail);
            ccEmails.push(cleanEmail);
          }
        }
      }
    }

    const txtParts = [];

    if (toEmails.length) {
      txtParts.push(`To:\n${toEmails.join(", ")}`);
    }

    if (ccEmails.length) {
      txtParts.push(`Cc:\n${ccEmails.join(", ")}`);
    }

    return {
      to: toEmails,
      cc: ccEmails,
      emails: [...toEmails, ...ccEmails],
      txt: txtParts.join("\n\n")
    };
  }

  function findAnchorForRowText(rowText, currentPanelIndex, anchorGroups) {
    const text = normalizeName(rowText);
    if (!text) return null;

    // exact match first
    if (anchorGroups.has(text)) {
      const candidates = anchorGroups.get(text);
      return candidates.find(c => c.panelIndex !== currentPanelIndex) || candidates[0];
    }

    // then "contains" match either way
    for (const [key, candidates] of anchorGroups.entries()) {
      if (text.includes(key) || key.includes(text)) {
        return candidates.find(c => c.panelIndex !== currentPanelIndex) || candidates[0];
      }
    }

    return null;
  }

  function linkRowsInReport(reportHtml, currentPanelIndex, anchorGroups) {
    const temp = document.createElement("div");
    temp.innerHTML = reportHtml;

    temp.querySelectorAll(".res_color_inner").forEach(row => {
      const target = findAnchorForRowText(row.textContent || "", currentPanelIndex, anchorGroups);
      if (!target) return;

      if (row.closest("a")) return;

      const link = temp.ownerDocument.createElement("a");
      link.className = "tm-row-link";
      link.href = `#${target.id}`;

      row.parentNode.insertBefore(link, row);
      link.appendChild(row);
    });

    return temp.innerHTML;
  }

  function buildTabbedHtml({ title, panels }) {
    const tabLinks = panels.map((p, i) =>
      `<a class="tm-tab-link" href="#tm-panel-${i}">${escapeHtml(p.label)}</a>`
    ).join("\n");

    const panelHtml = panels.map((p, i) => `
<div class="tm-panel ${i === 0 ? "tm-panel-default" : ""}" id="tm-panel-${i}">
  ${p.jumpbar}
  <div class="tm-scale-wrap">
    <div class="tm-report-wrap">
      ${p.report}
    </div>
  </div>
</div>`).join("\n");

    return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=5, user-scalable=yes">
<title>${escapeHtml(title)}</title>
<style>
:root { --mobile-scale: ${MOBILE_SCALE}; }
${CAPTURE_CSS}
</style>
</head>
<body>
<div id="top-tabs">
  ${tabLinks}
</div>
${panelHtml}
</body>
</html>`;
  }

  /* ================= WORKER ================= */
  async function worker() {
    const url = new URL(location.href);
    const runId = url.searchParams.get(RUN);
    const idx = Number(url.searchParams.get(IDX));

    if (!runId || Number.isNaN(idx)) return;

    try {
      await waitForReportBody(15000);

      const cfg = REPORTS[idx];
      const { date } = getReportMeta(document);

      if (cfg.mode === "email") {
        const { emails, txt } = extractEmails(document);

        await GM_setValue(KEY(runId), {
          ok: true,
          idx,
          label: cfg.label,
          mode: cfg.mode,
          date,
          emails,
          txt
        });
      } else {
        const { jumpbar, report, items } = extract(document, idx);

        await GM_setValue(KEY(runId), {
          ok: true,
          idx,
          label: cfg.label,
          mode: cfg.mode,
          date,
          jumpbar,
          report,
          items
        });
      }
    } catch (err) {
      await GM_setValue(KEY(runId), {
        ok: false,
        idx,
        label: REPORTS[idx]?.label || `Report ${idx + 1}`,
        error: String(err)
      });
    }

    setTimeout(() => window.close(), 300);
  }

  /* ================= CONTROLLER ================= */
  function waitForResult(runId) {
    return new Promise((resolve, reject) => {
      let listenerId;

      listenerId = GM_addValueChangeListener(KEY(runId), (_, __, value) => {
        if (!value) return;
        GM_removeValueChangeListener(listenerId);

        if (value.ok) {
          resolve(value);
        } else {
          reject(new Error(value.error || "Unknown capture error"));
        }
      });
    });
  }

  async function captureOne(runId, cfg, idx) {
    await GM_setValue(KEY(runId), null);

    const resultPromise = waitForResult(runId);

    const u = new URL(buildUrl(ymd(), cfg));
    u.searchParams.set(RUN, runId);
    u.searchParams.set(IDX, String(idx));

    GM_openInTab(u.toString(), { active: true, insert: true });

    return await resultPromise;
  }

  async function runCapture(selectedReports) {
    const panels = [];
    let emailExport = null;

    for (let idx = 0; idx < selectedReports.length; idx++) {
      const runId = `${Date.now().toString(36)}-${idx}`;
      const result = await captureOne(runId, selectedReports[idx], idx);

      if (result.mode === "email") {
        emailExport = result;
      } else {
        panels.push({
          label: result.label,
          jumpbar: result.jumpbar,
          report: result.report,
          date: result.date,
          items: Array.isArray(result.items) ? result.items : []
        });
      }
    }

    // Build name -> anchors map across all tabs
    const anchorGroups = new Map();
    for (const panel of panels) {
      for (const item of panel.items) {
        if (!anchorGroups.has(item.key)) {
          anchorGroups.set(item.key, []);
        }
        anchorGroups.get(item.key).push({
          id: item.id,
          panelIndex: item.panelIndex,
          name: item.name
        });
      }
    }

    // Add cross-tab row links into each report
    panels.forEach((panel, panelIndex) => {
      panel.report = linkRowsInReport(panel.report, panelIndex, anchorGroups);
    });

    const html = buildTabbedHtml({
      title: `Schedule It ${ymd()} - tabbed`,
      panels
    });

    const date = panels[0]?.date || emailExport?.date || ymd();
    const title = cleanFilename(FILE_LABEL);

    downloadHtml(`${date}_${title}.html`, html);

    if (emailExport && emailExport.txt) {
      const emailTitle = cleanFilename(EMAIL_FILE_LABEL);
      downloadTxt(`${date}_${emailTitle}.txt`, emailExport.txt);
    }
  }

  function controller() {
    GM_registerMenuCommand(THIS_WEEK_MENU_ITEM_NAME, async () => {
      try {
        await runCapture(getThisWeekReports());
      } catch (err) {
        alert(`Capture failed: ${err.message}`);
      }
    });

    GM_registerMenuCommand(NEXT_WEEK_MENU_ITEM_NAME, async () => {
      try {
        await runCapture(getNextWeekReports());
      } catch (err) {
        alert(`Capture failed: ${err.message}`);
      }
    });

    GM_registerMenuCommand(FIXED_DATE_RANGE_MENU_ITEM_NAME, async () => {
      try {
        await runCapture(getFixedDateRangeReports());
      } catch (err) {
        alert(`Capture failed: ${err.message}`);
      }
    });
  }

  location.pathname.includes("reports_create.php") ? worker() : controller();
})();
