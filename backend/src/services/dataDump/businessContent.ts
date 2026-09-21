import { pick, randomInt, type SeededSource } from "./seededRandom.js";
import type { NamingStyle } from "../../types/dataDump.js";

/**
 * Realistic professional content pools (spec §8/§11/§14/§15/§16: "avoid meaningless default names
 * such as Folder1/File1... unless the user explicitly selects synthetic/simple naming"). Pure data
 * + pure selection functions — no Graph, no I/O, fully unit-testable.
 */

export const DEPARTMENTS = [
  "Finance",
  "HR",
  "Sales",
  "Marketing",
  "Engineering",
  "Product",
  "Operations",
  "Legal",
  "Customer Success",
  "IT",
] as const;

export const PROJECT_NAMES = [
  "Project Alpha",
  "Project Beta",
  "Project Gamma",
  "Project Horizon",
  "Project Meridian",
  "Project Nova",
  "Project Orion",
  "Project Summit",
] as const;

const FIRST_NAMES = [
  "Alex", "Jordan", "Taylor", "Morgan", "Casey", "Riley", "Jamie", "Avery", "Cameron", "Drew",
  "Priya", "Wei", "Fatima", "Hiro", "Elena", "Marcus", "Nadia", "Oscar", "Sana", "Liam",
];
const LAST_NAMES = [
  "Carter", "Nguyen", "Patel", "Johansson", "Rossi", "Kimura", "Okafor", "Silva", "Novak", "Reyes",
  "Fischer", "Dubois", "Kowalski", "Haddad", "Larsen", "Petrov", "Cohen", "Diaz", "Moreau", "Andersson",
];

export function randomPersonName(source: SeededSource): { first: string; last: string; display: string } {
  const first = pick(source.rng, FIRST_NAMES);
  const last = pick(source.rng, LAST_NAMES);
  return { first, last, display: `${first} ${last}` };
}

export function randomDepartment(source: SeededSource): string {
  return pick(source.rng, DEPARTMENTS);
}

export function randomProjectName(source: SeededSource): string {
  return pick(source.rng, PROJECT_NAMES);
}

export interface FolderNode {
  name: string;
  children: FolderNode[];
}

/** Professional, folder-appropriate name pool — deliberately one flat pool reused at every depth (root, sub-folder, sub-sub-folder, ...) rather than a fixed per-category taxonomy, since the folder tree's SHAPE is now caller-controlled (total root folders × sub-folders-per-folder × nested levels) and any node at any depth needs an equally plausible name. */
const FOLDER_NAME_POOL = [
  "Finance", "HR", "Sales", "Marketing", "Engineering", "Product", "Operations", "Legal",
  "Customer Success", "IT", "Documents", "Reports", "Presentations", "Archive", "Projects",
  "Budgets", "Forecasts", "Invoices", "Audit", "Policies", "Recruitment", "Employees",
  "Onboarding", "Benefits", "Training", "Pipeline", "Contracts", "Proposals", "Renewals",
  "Territory Plans", "Campaigns", "Events", "Brand Assets", "Content Calendar", "Analytics",
  "Architecture", "Sprint Notes", "Design Docs", "Release Notes", "Postmortems", "Templates",
  "Resources", "Planning", "Reviews", "Notes", "Drafts", "Backups", "Working Files",
  "Deliverables", "Correspondence", "Reference", "Records", "Summaries", "Data", "Assets",
  "Vendors", "Clients", "Compliance", "Risk Management", "Strategy", "Research", "Onboarding Kits",
  "Playbooks", "Runbooks", "Case Studies", "Roadmaps", "Metrics", "Dashboards", "Governance",
] as const;

/** Deterministically picks a name never used before in this tree — scans the pool starting from a seeded offset (so runs still vary), and once the whole pool is exhausted falls back to numbered "Folder Set N" names that are still guaranteed unique. */
function nextUniqueFolderName(source: SeededSource, used: Set<string>): string {
  const startIndex = Math.floor(source.rng() * FOLDER_NAME_POOL.length);
  for (let i = 0; i < FOLDER_NAME_POOL.length; i++) {
    const candidate = FOLDER_NAME_POOL[(startIndex + i) % FOLDER_NAME_POOL.length]!;
    if (!used.has(candidate)) {
      used.add(candidate);
      return candidate;
    }
  }
  let n = 2;
  while (used.has(`Folder Set ${n}`)) n++;
  const name = `Folder Set ${n}`;
  used.add(name);
  return name;
}

/**
 * Builds a UNIFORM folder tree: `totalFolders` root folders, each recursively given
 * `subFoldersPerFolder` children down to `nestedLevels` levels deep (nestedLevels === 1 means root
 * folders only, no children at all). Every node gets a real, non-generic, GUARANTEED-unique name
 * across the whole tree (professional mode), or a flat, positionally-unique Folder1/Folder1-1/...
 * name (synthetic mode) — never two nodes sharing a name, at any depth.
 */
export function buildFolderTree(source: SeededSource, totalFolders: number, subFoldersPerFolder: number, nestedLevels: number, namingStyle: NamingStyle): FolderNode[] {
  const depth = Math.max(1, nestedLevels);
  const used = new Set<string>();

  function buildChildren(parentPath: string, level: number): FolderNode[] {
    if (level >= depth || subFoldersPerFolder <= 0) return [];
    const children: FolderNode[] = [];
    for (let i = 0; i < subFoldersPerFolder; i++) {
      const name = namingStyle === "synthetic" ? `${parentPath}-${i + 1}` : nextUniqueFolderName(source, used);
      children.push({ name, children: buildChildren(name, level + 1) });
    }
    return children;
  }

  const roots: FolderNode[] = [];
  for (let i = 0; i < totalFolders; i++) {
    const name = namingStyle === "synthetic" ? `Folder${i + 1}` : nextUniqueFolderName(source, used);
    roots.push({ name, children: buildChildren(name, 1) });
  }
  return roots;
}

/** Total node count across an entire folder tree (every root, sub-folder, and nested folder — the tree already contains exactly this many folders once created). */
export function countFolderNodes(tree: FolderNode[]): number {
  let count = 0;
  for (const node of tree) count += 1 + countFolderNodes(node.children);
  return count;
}

// ---------------------------------------------------------------------------
// Professional document content — full-fidelity generators for docx/xlsx/pdf/
// pptx/txt/csv. Each produces a complete, internally-consistent document
// (named stakeholders, real dates/dollar figures, spreadsheet formulas,
// numbered contract clauses with a signature block) rather than one
// placeholder paragraph repeated per section, and each also decides the
// file's own professional name (e.g. "Budget_Summary_2026", "Project_Charter")
// so a file's name and its content always describe the same thing.
// ---------------------------------------------------------------------------

export function formatLongDate(ms: number): string {
  return new Intl.DateTimeFormat("en-US", { month: "long", day: "numeric", year: "numeric" }).format(new Date(ms));
}

function formatCompactDate(ms: number): string {
  return new Date(ms).toISOString().slice(0, 10);
}

function addDays(ms: number, days: number): number {
  return ms + days * 24 * 60 * 60 * 1000;
}

/** Picks up to `count` distinct items from `pool` without replacement (Fisher-Yates partial shuffle). */
function sampleUnique<T>(source: SeededSource, pool: readonly T[], count: number): T[] {
  const arr = [...pool];
  const n = Math.min(count, arr.length);
  for (let i = 0; i < n; i++) {
    const j = i + Math.floor(source.rng() * (arr.length - i));
    [arr[i], arr[j]] = [arr[j]!, arr[i]!];
  }
  return arr.slice(0, n);
}

/** "Project Charter" -> "Project_Charter" — filesystem-safe, matches the underscore convention of real corporate file shares. */
function slugifyTitle(title: string): string {
  return title
    .replace(/[^\w\s-]/g, "")
    .trim()
    .replace(/\s+/g, "_");
}

function companyAbbrev(name: string): string {
  const words = name.split(" ").filter(Boolean);
  if (words.length === 1) return words[0]!.slice(0, 3).toUpperCase();
  return words.map((w) => w[0]).join("").toUpperCase();
}

const STREET_NAMES = [
  "Riverside Parkway", "Market Street", "Commerce Drive", "Innovation Way", "Corporate Boulevard",
  "Industrial Way", "Lakeview Avenue", "Sunset Boulevard", "Harbor Street", "Congress Avenue",
  "Enterprise Drive", "Meridian Street",
];
const CITIES = [
  { city: "Denver", state: "CO", zip: "80202" },
  { city: "Austin", state: "TX", zip: "78701" },
  { city: "Seattle", state: "WA", zip: "98101" },
  { city: "Chicago", state: "IL", zip: "60601" },
  { city: "Atlanta", state: "GA", zip: "30301" },
  { city: "Boston", state: "MA", zip: "02110" },
  { city: "San Diego", state: "CA", zip: "92101" },
  { city: "Charlotte", state: "NC", zip: "28202" },
  { city: "Minneapolis", state: "MN", zip: "55401" },
  { city: "Phoenix", state: "AZ", zip: "85001" },
];

function randomAddress(source: SeededSource): string {
  const num = randomInt(source.rng, 100, 9800);
  const street = pick(source.rng, STREET_NAMES);
  const suite = randomInt(source.rng, 100, 480);
  const place = pick(source.rng, CITIES);
  return `${num} ${street}, Suite ${suite}, ${place.city}, ${place.state} ${place.zip}`;
}

const EXEC_TITLES = ["Chief Executive Officer", "Chief Operating Officer", "Chief Financial Officer", "Managing Director", "General Counsel", "VP of Operations"];
function randomExecTitle(source: SeededSource): string {
  return pick(source.rng, EXEC_TITLES);
}

// ---- Project Charter (docx) ------------------------------------------------

export interface CharterContent {
  projectName: string;
  sponsor: { name: string; title: string };
  manager: { name: string; title: string };
  department: string;
  startDateMs: number;
  targetCompletionMs: number;
  purpose: string;
  objectives: string[];
  scopeIn: string;
  scopeOut: string;
  stakeholders: { name: string; role: string; responsibility: string }[];
  phases: { name: string; startMs: number; endMs: number }[];
  budget: number;
  risks: { risk: string; mitigation: string }[];
  approvalDateMs: number;
}

const LEGACY_PROBLEM_POOL = [
  "built on unsupported infrastructure, resulting in recurring performance issues, limited mobile compatibility, and rising maintenance costs",
  "reliant on manual, spreadsheet-based processes that are error-prone and difficult to audit",
  "running on a vendor platform reaching end-of-life support within the next twelve months",
  "unable to scale with the organization's current headcount and transaction volume",
  "fragmented across multiple disconnected tools, creating duplicate data entry and reporting gaps",
];
const SYSTEM_AREA_POOL = ["client-facing portal", "internal reporting system", "case management platform", "procurement workflow", "employee self-service system", "billing platform"];
const RISK_POOL = [
  { risk: "Data migration errors", mitigation: "mitigated through staged migration and validation checkpoints." },
  { risk: "Resource availability during peak periods", mitigation: "mitigated by securing contractor backup resources in advance." },
  { risk: "Scope creep from stakeholder requests", mitigation: "mitigated through a formal change-control process." },
  { risk: "Vendor delivery delays", mitigation: "mitigated through weekly vendor checkpoint calls and contractual SLAs." },
  { risk: "Budget overrun", mitigation: "mitigated through monthly budget-to-actual reviews with Finance." },
  { risk: "User adoption resistance", mitigation: "mitigated through early stakeholder training and a phased rollout." },
  { risk: "Security or compliance gaps", mitigation: "mitigated through a pre-launch review with the Information Security team." },
];
const STAKEHOLDER_ROLE_POOL = [
  { role: "Lead Engineer", responsibility: "Technical architecture and delivery" },
  { role: "QA Lead", responsibility: "Test planning and quality sign-off" },
  { role: "Customer Success Lead", responsibility: "Client communication and training" },
  { role: "Business Analyst", responsibility: "Requirements gathering and documentation" },
  { role: "Change Management Lead", responsibility: "Stakeholder communication and adoption" },
];
const PHASE_NAME_POOL = ["Discovery & Requirements", "Design & Architecture", "Development", "Testing & UAT", "Launch & Decommission"];

function objectivePool(source: SeededSource): string[] {
  const pct = () => randomInt(source.rng, 15, 60);
  return [
    `Reduce average processing time by ${pct()}% compared to the current baseline.`,
    "Achieve full responsive support across desktop, tablet, and mobile devices.",
    "Migrate all existing records and historical data without loss.",
    `Decommission the legacy platform by the end of Q${randomInt(source.rng, 1, 4)} ${randomInt(source.rng, 2026, 2027)}.`,
    `Maintain a customer satisfaction score of 4.${randomInt(source.rng, 3, 8)}/5 or higher post-launch.`,
    `Achieve ${pct()}% automated test coverage prior to launch.`,
    `Reduce support ticket volume by ${pct()}%.`,
    `Complete staff onboarding within ${randomInt(source.rng, 2, 6)} weeks of go-live.`,
    `Improve system uptime to 99.${randomInt(source.rng, 5, 9)}%.`,
    `Cut operating costs by $${randomInt(source.rng, 20, 150)},000 annually.`,
  ];
}

export function buildCharterContent(source: SeededSource): CharterContent {
  const department = randomDepartment(source);
  const projectName = randomProjectName(source);
  const sponsor = randomPersonName(source);
  const manager = randomPersonName(source);
  const startDateMs = Date.now() - randomInt(source.rng, 0, 60) * 24 * 60 * 60 * 1000;
  const targetCompletionMs = addDays(startDateMs, randomInt(source.rng, 120, 240));

  const otherStakeholders = sampleUnique(source, STAKEHOLDER_ROLE_POOL, randomInt(source.rng, 1, STAKEHOLDER_ROLE_POOL.length));
  const stakeholders = [
    { name: sponsor.display, role: "Project Sponsor", responsibility: "Funding approval and strategic direction" },
    { name: manager.display, role: "Project Manager", responsibility: "Day-to-day execution and reporting" },
    ...otherStakeholders.map((r) => ({ name: randomPersonName(source).display, role: r.role, responsibility: r.responsibility })),
  ];

  const phaseNames = PHASE_NAME_POOL.slice(0, randomInt(source.rng, 3, PHASE_NAME_POOL.length));
  const totalSpan = targetCompletionMs - startDateMs;
  const phases: CharterContent["phases"] = [];
  let cursor = startDateMs;
  for (let i = 0; i < phaseNames.length; i++) {
    const share = Math.round(totalSpan / phaseNames.length);
    const endMs = i === phaseNames.length - 1 ? targetCompletionMs : cursor + share;
    phases.push({ name: phaseNames[i]!, startMs: cursor, endMs });
    cursor = endMs;
  }

  return {
    projectName,
    sponsor: { name: sponsor.display, title: `VP of ${pick(source.rng, ["Customer Experience", "Operations", "Engineering", "Finance", "Product", "Information Technology"])}` },
    manager: { name: manager.display, title: "Project Manager" },
    department,
    startDateMs,
    targetCompletionMs,
    purpose: `The ${projectName} will replace the organization's legacy ${pick(source.rng, SYSTEM_AREA_POOL)} with a modern, secure, and scalable platform. The current system is ${pick(source.rng, LEGACY_PROBLEM_POOL)}. This project addresses these gaps while improving the overall ${department.toLowerCase()} experience.`,
    objectives: sampleUnique(source, objectivePool(source), 4),
    scopeIn: `redesign and development of the ${pick(source.rng, SYSTEM_AREA_POOL)}, data migration from the legacy system, integration with existing systems, and staff training`,
    scopeOut: "changes to unrelated backend systems and any initiative not explicitly named above, which will be addressed separately",
    stakeholders,
    phases,
    budget: randomInt(source.rng, 80, 650) * 1000,
    risks: sampleUnique(source, RISK_POOL, randomInt(source.rng, 3, 5)),
    approvalDateMs: startDateMs - randomInt(source.rng, 1, 5) * 24 * 60 * 60 * 1000,
  };
}

// ---- Business Requirements Document (docx) --------------------------------

export interface RequirementsContent {
  projectName: string;
  department: string;
  author: { name: string; title: string };
  overview: string;
  functionalRequirements: string[];
  nonFunctionalRequirements: string[];
  assumptions: string[];
  approver: { name: string; title: string };
  approvalDateMs: number;
}

const FUNCTIONAL_REQ_POOL = [
  "Users must be able to search and filter records by date, owner, and status.",
  "The system must support role-based access control with at least three permission tiers.",
  "Users must be able to export any report to PDF or CSV.",
  "The system must send an email notification when a record changes status.",
  "Administrators must be able to bulk-import records from a CSV template.",
  "The system must retain a full audit trail of every create, update, and delete action.",
  "Users must be able to save and re-run their own custom filters.",
  "The system must support single sign-on via the organization's identity provider.",
  "Reviewers must be able to approve or reject a submission with a required comment.",
  "The system must display real-time status updates without requiring a page refresh.",
];
const NON_FUNCTIONAL_REQ_POOL = [
  "95% of page loads must complete in under 2 seconds under normal load.",
  "The system must maintain 99.9% uptime, measured monthly.",
  "All data in transit and at rest must be encrypted.",
  "The system must support at least 500 concurrent users without degradation.",
  "The system must be fully usable on the latest two major versions of Chrome, Edge, and Safari.",
  "Backups must run nightly with a 30-day retention window.",
];
const ASSUMPTION_POOL = [
  "Existing user accounts and permissions will be provided by IT prior to migration.",
  "Third-party API access will be available in the staging environment before UAT begins.",
  "Business stakeholders will be available for weekly review sessions throughout development.",
  "No changes to the underlying data model are required from other in-flight projects.",
  "Training materials will be reviewed by the Customer Success team before go-live.",
];

export function buildRequirementsContent(source: SeededSource): RequirementsContent {
  const department = randomDepartment(source);
  const projectName = randomProjectName(source);
  const author = randomPersonName(source);
  return {
    projectName,
    department,
    author: { name: author.display, title: randomJobTitle(source) },
    overview: `This document defines the functional and non-functional requirements for ${projectName}, sponsored by the ${department} team. It is intended for the project team, QA, and any stakeholder validating scope prior to development.`,
    functionalRequirements: sampleUnique(source, FUNCTIONAL_REQ_POOL, randomInt(source.rng, 5, 7)),
    nonFunctionalRequirements: sampleUnique(source, NON_FUNCTIONAL_REQ_POOL, randomInt(source.rng, 3, 5)),
    assumptions: sampleUnique(source, ASSUMPTION_POOL, randomInt(source.rng, 2, 4)),
    approver: { name: randomPersonName(source).display, title: randomExecTitle(source) },
    approvalDateMs: Date.now() - randomInt(source.rng, 0, 30) * 24 * 60 * 60 * 1000,
  };
}

// ---- Standard Operating Procedure (docx) -----------------------------------

export interface PolicyContent {
  title: string;
  department: string;
  owner: { name: string; title: string };
  purpose: string;
  scope: string;
  steps: string[];
  responsibilities: { role: string; duty: string }[];
  effectiveDateMs: number;
  reviewDateMs: number;
  policyRef: string;
}

const POLICY_TOPIC_POOL = [
  {
    title: "Expense Reimbursement",
    steps: [
      "Employee submits an itemized expense report with receipts within 30 days of purchase.",
      "Manager reviews the report for policy compliance and approves or returns it within 3 business days.",
      "Finance validates the approved report against the current spending policy.",
      "Finance issues reimbursement in the next scheduled payroll cycle.",
    ],
  },
  {
    title: "New Hire Onboarding",
    steps: [
      "HR sends the offer letter and new-hire paperwork upon acceptance.",
      "IT provisions accounts, hardware, and system access before the start date.",
      "Manager assigns an onboarding buddy and a 30/60/90-day plan.",
      "HR schedules orientation and benefits enrollment within the first week.",
    ],
  },
  {
    title: "Change Management",
    steps: [
      "Requestor submits a change request describing scope, risk, and rollback plan.",
      "Change Advisory Board reviews and approves or rejects the request.",
      "Approved changes are scheduled during the next maintenance window.",
      "Change owner confirms success and closes the request with a post-change summary.",
    ],
  },
  {
    title: "Data Retention",
    steps: [
      "Data owner classifies each dataset by sensitivity and required retention period.",
      "IT configures automated archival at the end of the active retention window.",
      "Legal reviews any request to extend retention beyond the standard policy.",
      "Data marked for deletion is purged on the next scheduled cycle and logged for audit.",
    ],
  },
  {
    title: "Incident Response",
    steps: [
      "On-call engineer triages the incident and assigns a severity level within 15 minutes.",
      "Incident commander coordinates response and stakeholder communication.",
      "Team resolves the incident and confirms service is restored.",
      "Team publishes a postmortem with root cause and follow-up actions within 5 business days.",
    ],
  },
];

export function buildPolicyContent(source: SeededSource): PolicyContent {
  const topic = pick(source.rng, POLICY_TOPIC_POOL);
  const department = randomDepartment(source);
  const owner = randomPersonName(source);
  const effectiveDateMs = Date.now() - randomInt(source.rng, 0, 180) * 24 * 60 * 60 * 1000;
  return {
    title: topic.title,
    department,
    owner: { name: owner.display, title: randomJobTitle(source) },
    purpose: `This procedure defines the standard process for ${topic.title.toLowerCase()} within the ${department} organization, ensuring consistency, accountability, and auditability.`,
    scope: `This procedure applies to all ${department} staff and any employee submitting a request that falls under ${topic.title.toLowerCase()}.`,
    steps: topic.steps,
    responsibilities: [
      { role: "Process Owner", duty: "Maintains this procedure and reviews it annually." },
      { role: "Manager", duty: "Ensures direct reports follow this procedure." },
      { role: "Employee", duty: "Follows the steps above and escalates exceptions promptly." },
    ],
    effectiveDateMs,
    reviewDateMs: addDays(effectiveDateMs, 365),
    policyRef: `${department.slice(0, 3).toUpperCase()}-${randomInt(source.rng, 100, 999)}`,
  };
}

export type DocxPlan =
  | { kind: "charter"; fileNameBase: string; charter: CharterContent }
  | { kind: "requirements"; fileNameBase: string; requirements: RequirementsContent }
  | { kind: "policy"; fileNameBase: string; policy: PolicyContent };

export function buildDocxPlan(source: SeededSource): DocxPlan {
  const kind = pick(source.rng, ["charter", "requirements", "policy"] as const);
  if (kind === "charter") {
    const charter = buildCharterContent(source);
    const suffix = source.rng() < 0.3 ? `_${new Date(charter.startDateMs).getFullYear()}` : "";
    return { kind, fileNameBase: `Project_Charter${suffix}`, charter };
  }
  if (kind === "requirements") {
    const requirements = buildRequirementsContent(source);
    return { kind, fileNameBase: pick(source.rng, ["Business_Requirements_Document", "Requirements_Specification"]), requirements };
  }
  const policy = buildPolicyContent(source);
  return { kind, fileNameBase: `${slugifyTitle(policy.title)}_SOP`, policy };
}

// ---- Budget / headcount workbook (xlsx) ------------------------------------

export interface QuarterlyWorkbookContent {
  title: string;
  subtitle: string;
  categoryHeader: string;
  categories: { name: string; quarters: [number, number, number, number] }[];
  unit: "currency" | "count";
  notes: string[];
}

const BUDGET_CATEGORY_POOL = ["Software Licensing", "Contractor Resources", "Cloud Infrastructure", "Training & Enablement", "Quality Assurance Tools", "Travel & Entertainment", "Marketing Spend", "Facilities"];
const HEADCOUNT_CATEGORY_POOL = ["Engineering", "Product", "Sales", "Customer Success", "Marketing", "Operations", "Finance", "HR"];

export function buildXlsxPlan(source: SeededSource): { kind: "budget" | "headcount"; fileNameBase: string; content: QuarterlyWorkbookContent } {
  const year = randomInt(source.rng, 2026, 2027);
  const projectName = randomProjectName(source);
  if (source.rng() < 0.6) {
    const categories = sampleUnique(source, BUDGET_CATEGORY_POOL, randomInt(source.rng, 5, 7)).map((name) => ({
      name,
      quarters: [randomInt(source.rng, 3, 60) * 1000, randomInt(source.rng, 3, 60) * 1000, randomInt(source.rng, 3, 60) * 1000, randomInt(source.rng, 3, 60) * 1000] as [number, number, number, number],
    }));
    return {
      kind: "budget",
      fileNameBase: `Budget_Summary_${year}`,
      content: {
        title: `FY${year} Budget Summary — ${projectName}`,
        subtitle: "Prepared by: Finance & Program Management Office | Currency: USD",
        categoryHeader: "Category",
        categories,
        unit: "currency",
        notes: [
          `Blue figures are budgeted inputs approved by Finance on ${formatLongDate(Date.now() - randomInt(source.rng, 30, 120) * 24 * 60 * 60 * 1000)}.`,
          `Contingency Reserve is held at 5% of total quarterly spend per Finance policy FIN-${randomInt(source.rng, 10, 99)}.`,
        ],
      },
    };
  }
  const categories = sampleUnique(source, HEADCOUNT_CATEGORY_POOL, randomInt(source.rng, 4, 6)).map((name) => ({
    name,
    quarters: [randomInt(source.rng, 1, 8), randomInt(source.rng, 1, 8), randomInt(source.rng, 1, 8), randomInt(source.rng, 1, 8)] as [number, number, number, number],
  }));
  return {
    kind: "headcount",
    fileNameBase: `Headcount_Plan_${year}`,
    content: {
      title: `FY${year} Headcount Plan`,
      subtitle: "Prepared by: People Operations | Planned net-new hires by quarter",
      categoryHeader: "Department",
      categories,
      unit: "count",
      notes: [`Figures reflect approved requisitions as of ${formatLongDate(Date.now())} and are subject to quarterly budget review.`],
    },
  };
}

// ---- Contracts: Master Services Agreement / Mutual NDA (pdf) --------------

export interface ContractContent {
  agreementTitle: string;
  vendorCompany: string;
  vendorAddress: string;
  clientCompany: string;
  clientAddress: string;
  effectiveDateMs: number;
  termMonths: number;
  projectRef: string;
  clauses: { title: string; body: string }[];
  vendorSigner: { name: string; title: string };
  clientSigner: { name: string; title: string };
  contractRef: string;
}

const CONSULTING_FIRM_POOL = ["Northfield Consulting Group", "Meridian Advisory Partners", "Beacon Hill Consulting", "Summit Bridge Advisors", "Gray Harbor Consulting"];

function msaClauses(projectName: string, projectRef: string): { title: string; body: string }[] {
  return [
    { title: "1. Scope of Services", body: `Consultant agrees to provide professional advisory and implementation services related to Client's ${projectName}, as further described in Statement of Work No. ${projectRef}, attached and incorporated by reference.` },
    { title: "2. Term", body: "This Agreement shall commence on the Effective Date and continue for the term stated above, unless terminated earlier in accordance with the Termination clause. The parties may renew this Agreement for additional terms by mutual written consent." },
    { title: "3. Fees and Payment", body: "Client agrees to pay Consultant in accordance with the fee schedule set forth in the applicable Statement of Work. Invoices shall be issued monthly and are due within thirty (30) days of receipt. Late payments accrue interest at 1.5% per month." },
    { title: "4. Confidentiality", body: "Each party agrees to protect the other's confidential information using the same degree of care it uses for its own confidential information, and not to disclose such information to third parties without prior written consent, except as required by law." },
    { title: "5. Intellectual Property", body: "All work product created specifically for Client under an approved Statement of Work shall be owned by Client upon full payment. Consultant retains ownership of its pre-existing methodologies, tools, and templates used in delivering the Services." },
    { title: "6. Termination", body: "Either party may terminate this Agreement for convenience with sixty (60) days' written notice, or immediately for material breach that remains uncured after fifteen (15) days' written notice." },
    { title: "7. Governing Law", body: "This Agreement shall be governed by and construed in accordance with the laws of the state in which Consultant maintains its principal office, without regard to its conflict of laws principles." },
  ];
}

function ndaClauses(partyA: string, partyB: string): { title: string; body: string }[] {
  return [
    { title: "1. Purpose", body: `${partyA} and ${partyB} wish to explore a potential business relationship and, in connection with that opportunity, may disclose confidential information to one another.` },
    { title: "2. Definition of Confidential Information", body: "Confidential Information means any non-public technical, business, or financial information disclosed by either party, whether oral or written, that is designated as confidential or would reasonably be understood to be confidential." },
    { title: "3. Obligations", body: "Each party agrees to use the other's Confidential Information solely to evaluate the potential relationship, and to protect it using at least the same degree of care it uses for its own confidential information." },
    { title: "4. Exclusions", body: "Confidential Information does not include information that is or becomes publicly available through no fault of the receiving party, or that was already lawfully known prior to disclosure." },
    { title: "5. Term", body: "This Agreement remains in effect for the term stated above. Confidentiality obligations survive termination for a period of three (3) years." },
    { title: "6. No License", body: "Nothing in this Agreement grants either party any license or ownership right in the other's Confidential Information, patents, or trade secrets." },
    { title: "7. Governing Law", body: "This Agreement shall be governed by and construed in accordance with the laws of the state in which the disclosing party maintains its principal office." },
  ];
}

export function buildContractContent(source: SeededSource, kind: "msa" | "nda"): ContractContent {
  const vendorCompany = kind === "msa" ? pick(source.rng, CONSULTING_FIRM_POOL) : randomCompanyName(source);
  const clientPool = COMPANY_NAMES.filter((c) => c !== vendorCompany);
  const clientCompany = pick(source.rng, clientPool.length > 0 ? clientPool : COMPANY_NAMES);
  const projectName = randomProjectName(source);
  const projectRef = `SOW-${randomInt(source.rng, 2026, 2027)}-${String(randomInt(source.rng, 1, 999)).padStart(3, "0")}`;
  const contractRef = `${companyAbbrev(vendorCompany)}-${companyAbbrev(clientCompany)}-${randomInt(source.rng, 2026, 2027)}-${String(randomInt(source.rng, 1, 999)).padStart(3, "0")}`;
  return {
    agreementTitle: kind === "msa" ? "Master Services Agreement" : "Mutual Non-Disclosure Agreement",
    vendorCompany,
    vendorAddress: randomAddress(source),
    clientCompany,
    clientAddress: randomAddress(source),
    effectiveDateMs: Date.now() - randomInt(source.rng, 0, 90) * 24 * 60 * 60 * 1000,
    termMonths: pick(source.rng, [6, 12, 18, 24]),
    projectRef,
    clauses: kind === "msa" ? msaClauses(projectName, projectRef) : ndaClauses(vendorCompany, clientCompany),
    vendorSigner: { name: randomPersonName(source).display, title: randomExecTitle(source) },
    clientSigner: { name: randomPersonName(source).display, title: randomExecTitle(source) },
    contractRef,
  };
}

export function buildPdfPlan(source: SeededSource): { kind: "msa" | "nda"; fileNameBase: string; content: ContractContent } {
  const kind = pick(source.rng, ["msa", "nda"] as const);
  const content = buildContractContent(source, kind);
  const fileNameBase =
    kind === "msa"
      ? pick(source.rng, ["Client_Agreement_Signed", "Service_Agreement_Signed", "Master_Agreement_Final"])
      : pick(source.rng, ["NDA_Signed", "Confidentiality_Agreement_Signed", "NDA_Executed"]);
  return { kind, fileNameBase, content };
}

// ---- Presentation deck (pptx) ---------------------------------------------

export interface PresentationContent {
  title: string;
  subtitle: string;
  slides: { heading: string; bullets: string[] }[];
}

const PRESENTATION_OUTLINE = ["Company Overview", "Business Strategy", "Market Analysis", "Revenue Performance", "Customer Growth", "Product Roadmap", "Next Steps"];
const PRESENTATION_BULLET_POOL: Record<string, string[]> = {
  "Company Overview": ["Founded to modernize how the business serves its customers", "Operating across multiple regions with a growing team", "Focused on sustainable, profitable growth"],
  "Business Strategy": ["Prioritize retention over pure new-logo acquisition", "Invest in automation to protect margin", "Expand into two adjacent market segments"],
  "Market Analysis": ["Total addressable market continues to grow year over year", "Competitive pressure concentrated among a few key players", "Customer demand shifting toward self-service options"],
  "Revenue Performance": ["Revenue trending ahead of plan for the current fiscal year", "Net retention remains above the 12-month average", "New bookings concentrated in the enterprise segment"],
  "Customer Growth": ["Net new customer count up over the prior period", "Churn remains within the target range", "Expansion revenue outpacing new-logo revenue"],
  "Product Roadmap": ["Near-term focus on performance and reliability improvements", "Mid-term investment in self-service onboarding", "Long-term bet on platform extensibility"],
  "Next Steps": ["Finalize budget and staffing for the next quarter", "Confirm executive sponsorship for the initiative", "Schedule a follow-up review in 30 days"],
};

export function buildPresentationPlan(source: SeededSource): { fileNameBase: string; content: PresentationContent } {
  const department = randomDepartment(source);
  const quarter = `Q${randomInt(source.rng, 1, 4)}_${randomInt(source.rng, 2026, 2027)}`;
  const title = pick(source.rng, ["Quarterly Business Review", "Board Presentation", "Sales Kickoff Deck", "Product Roadmap Review", "Company All-Hands"]);
  return {
    fileNameBase: `${slugifyTitle(title)}_${quarter}`,
    content: {
      title,
      subtitle: `${department} · ${quarter.replace("_", " ")}`,
      slides: PRESENTATION_OUTLINE.map((heading) => ({ heading, bullets: sampleUnique(source, PRESENTATION_BULLET_POOL[heading] ?? [], 3) })),
    },
  };
}

// ---- Status memo (txt) -----------------------------------------------------

export interface MemoContent {
  title: string;
  dateMs: number;
  from: { name: string; title: string };
  to: string;
  subject: string;
  points: string[];
  actionItems: { owner: string; item: string; due: string }[];
}

const MEMO_TOPIC_POOL = ["Weekly Status Update", "Meeting Minutes", "Project Update Memo", "Budget Review Notes"];
const MEMO_POINT_POOL = [
  "Development remains on track against the current milestone plan.",
  "QA identified two medium-severity issues, both with fixes scheduled this week.",
  "Vendor delivery for the integration module slipped by three business days.",
  "Stakeholder feedback from last week's demo has been incorporated into the backlog.",
  "Budget spend to date remains within 5% of the approved forecast.",
  "Onboarding for the two newest team members is complete.",
  "The staging environment refresh is scheduled for this weekend.",
  "Customer pilot feedback has been largely positive, with one open concern on reporting.",
];
const MEMO_ACTION_POOL = ["Follow up with the vendor on timeline", "Update the project tracker", "Circulate the revised budget", "Confirm attendee list for next review", "Draft the customer-facing summary"];

export function buildMemoPlan(source: SeededSource): { fileNameBase: string; content: MemoContent } {
  const department = randomDepartment(source);
  const author = randomPersonName(source);
  const title = pick(source.rng, MEMO_TOPIC_POOL);
  const dateMs = Date.now() - randomInt(source.rng, 0, 60) * 24 * 60 * 60 * 1000;
  const owners = sampleUnique(source, [author.display, randomPersonName(source).display, randomPersonName(source).display], 2);
  return {
    fileNameBase: `${slugifyTitle(title)}_${formatCompactDate(dateMs)}`,
    content: {
      title,
      dateMs,
      from: { name: author.display, title: randomJobTitle(source) },
      to: `${department} Team`,
      subject: `${title} — ${department}`,
      points: sampleUnique(source, MEMO_POINT_POOL, randomInt(source.rng, 3, 5)),
      actionItems: owners.map((owner) => ({ owner, item: pick(source.rng, MEMO_ACTION_POOL), due: formatLongDate(addDays(dateMs, randomInt(source.rng, 3, 10))) })),
    },
  };
}

// ---- Data export (csv) -----------------------------------------------------

export interface DataExportRow {
  employeeId: string;
  department: string;
  region: string;
  quarter: string;
  revenue: number;
  target: number;
  actual: number;
}

const REGIONS = ["North America", "EMEA", "APAC", "LATAM"];

export function buildDataExportRows(source: SeededSource, rowCount: number): DataExportRow[] {
  const rows: DataExportRow[] = [];
  for (let i = 0; i < rowCount; i++) {
    const target = randomInt(source.rng, 50_000, 500_000);
    const variance = randomInt(source.rng, -20, 20) / 100;
    rows.push({
      employeeId: `EMP-${1000 + i}`,
      department: randomDepartment(source),
      region: pick(source.rng, REGIONS),
      quarter: `Q${randomInt(source.rng, 1, 4)} ${randomInt(source.rng, 2026, 2027)}`,
      revenue: Math.round(target * (1 + variance)),
      target,
      actual: Math.round(target * (1 + variance)),
    });
  }
  return rows;
}

export function buildDataExportPlan(source: SeededSource): { fileNameBase: string; rows: DataExportRow[] } {
  const year = randomInt(source.rng, 2026, 2027);
  return {
    fileNameBase: pick(source.rng, [`Regional_Revenue_Data_Export_${year}`, `Sales_Performance_Export_${year}`, `Quarterly_Metrics_Export_${year}`]),
    rows: buildDataExportRows(source, randomInt(source.rng, 20, 60)),
  };
}

// ---- Chart image (png/jpg) and archive (zip) — naming only ----------------

export function buildChartPlan(source: SeededSource): { fileNameBase: string } {
  return { fileNameBase: pick(source.rng, ["Quarterly_Revenue_Chart", "Regional_Sales_Breakdown", "Budget_Allocation_Chart", "Customer_Growth_Trend", "Pipeline_Forecast_Chart"]) };
}

export function buildArchivePlan(source: SeededSource): { fileNameBase: string; memo: MemoContent } {
  const { content: memo } = buildMemoPlan(source);
  const year = randomInt(source.rng, 2026, 2027);
  return {
    fileNameBase: pick(source.rng, [`Project_Archive_${year}`, `Backup_${randomDepartment(source).replace(/\s+/g, "_")}_${year}`, `Document_Bundle_${year}`]),
    memo,
  };
}

const EMAIL_SUBJECTS = [
  "Quarterly Business Review",
  "Project Status Update",
  "Customer Meeting Follow-up",
  "Invoice Approval Required",
  "HR Policy Update",
  "Release Planning",
  "Migration Validation Results",
  "Weekly Team Sync Notes",
  "Budget Approval Request",
  "Onboarding Checklist",
];

export function randomEmailSubject(source: SeededSource): string {
  return pick(source.rng, EMAIL_SUBJECTS);
}

export function buildEmailBody(source: SeededSource, subject: string): string {
  return [
    `Hi team,`,
    ``,
    `Sharing an update regarding "${subject}". Please review the attached details and let me know if you have questions.`,
    ``,
    `Thanks,`,
    `${randomPersonName(source).display}`,
  ].join("\n");
}

const CHANNEL_MESSAGE_TEMPLATES = [
  "Good morning team. Please review the Q3 forecast before Friday.",
  "The migration validation is complete.",
  "Can someone review the latest project requirements?",
  "The customer demo has been scheduled for tomorrow.",
  "Please review the latest release notes.",
  "Reminder: standup moved to 10am today.",
  "Great work on the launch everyone!",
  "Uploading the updated deck to the Files tab now.",
];

export function randomChannelMessage(source: SeededSource): string {
  return pick(source.rng, CHANNEL_MESSAGE_TEMPLATES);
}

const CHANNEL_NAME_POOLS: Record<string, string[]> = {
  Finance: ["General", "Budget Planning", "Forecasting", "Announcements"],
  Engineering: ["General", "Development", "QA", "Releases"],
  Marketing: ["General", "Campaigns", "Events", "Announcements"],
  Sales: ["General", "Pipeline Reviews", "Territory Planning", "Announcements"],
  HR: ["General", "Policies", "Onboarding", "Announcements"],
};

const EVENT_TITLES = [
  "Quarterly Business Review",
  "Sprint Planning",
  "Customer Kickoff Call",
  "Budget Review",
  "All-Hands Meeting",
  "1:1 Check-in",
  "Product Roadmap Review",
  "Vendor Onboarding Call",
  "Migration Cutover Window",
  "Training Session",
];

export function randomEventTitle(source: SeededSource): string {
  return pick(source.rng, EVENT_TITLES);
}

export function buildEventBody(source: SeededSource, title: string): string {
  return `${title} — synthetic event generated by CloudFuze Data Dump for ${randomDepartment(source)}.`;
}

const JOB_TITLES = ["Account Manager", "Solutions Engineer", "Program Manager", "Financial Analyst", "HR Business Partner", "Software Engineer", "Sales Director", "Marketing Specialist"];
const COMPANY_NAMES = ["Contoso Ltd", "Fabrikam Inc", "Northwind Traders", "Adventure Works", "Tailspin Toys", "Wingtip Toys", "Litware Inc"];

export function randomJobTitle(source: SeededSource): string {
  return pick(source.rng, JOB_TITLES);
}

export function randomCompanyName(source: SeededSource): string {
  return pick(source.rng, COMPANY_NAMES);
}

export function randomPhoneNumber(source: SeededSource): string {
  const n = () => randomInt(source.rng, 0, 9);
  return `+1-555-${n()}${n()}${n()}-${n()}${n()}${n()}${n()}`;
}

export function channelNamesForTeam(teamDepartment: string, count: number): string[] {
  const pool = CHANNEL_NAME_POOLS[teamDepartment] ?? ["General", "Discussion", "Planning", "Announcements"];
  const names = pool.slice(0, count);
  while (names.length < count) names.push(`Channel ${names.length + 1}`);
  return names;
}
