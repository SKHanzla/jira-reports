import { NextResponse } from "next/server";

const EMAIL = process.env.JIRA_EMAIL!;
const API_TOKEN = process.env.JIRA_API_TOKEN!;
const DOMAIN = process.env.JIRA_DOMAIN!;

const MAX_RESULTS_PER_PAGE = 100;
const ABSOLUTE_MAX_RESULTS = 5000;

const FIELDS = [
  "summary","project","issuetype","status","parent",
  "customfield_10014","customfield_10011","labels",
  "worklog","timespent","timeoriginalestimate",
];

function getAuth(): string {
  return Buffer.from(`${EMAIL}:${API_TOKEN}`).toString("base64");
}

async function fetchJiraIssues(auth: string, jql: string): Promise<Record<string, unknown>[]> {
  const url = `https://${DOMAIN}/rest/api/3/search/jql`;
  const headers = { Authorization: `Basic ${auth}`, Accept: "application/json" };
  const allIssues: Record<string, unknown>[] = [];
  let startAt = 0;

  while (true) {
    const params = new URLSearchParams({
      jql,
      fields: FIELDS.join(","),
      maxResults: String(MAX_RESULTS_PER_PAGE),
      startAt: String(startAt),
    });
    const res = await fetch(`${url}?${params}`, { headers });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(`Jira API ${res.status}: ${txt.slice(0, 300)}`);
    }
    const data = await res.json();
    const issues: Record<string, unknown>[] = data.issues || [];
    const total: number = data.total || 0;
    allIssues.push(...issues);
    if (allIssues.length >= total || allIssues.length >= ABSOLUTE_MAX_RESULTS || issues.length === 0) break;
    startAt = allIssues.length;
  }
  return allIssues;
}

async function getWorklogHours(auth: string, issueKey: string): Promise<number> {
  try {
    const res = await fetch(`https://${DOMAIN}/rest/api/3/issue/${issueKey}/worklog`, {
      headers: { Authorization: `Basic ${auth}`, Accept: "application/json" },
    });
    if (!res.ok) return 0;
    const data = await res.json();
    const worklogs: Record<string, unknown>[] = data.worklogs || [];
    const secs = worklogs.reduce((s, w) => s + ((w.timeSpentSeconds as number) || 0), 0);
    return Math.round((secs / 3600) * 100) / 100;
  } catch { return 0; }
}

export async function GET() {
  try {
    const auth = getAuth();

    // Step 1: Fetch CDEF Epics
    let epicIssues = await fetchJiraIssues(auth, 'labels = "CDEF" AND (type = Epic OR issuetype = Epic)');

    if (!epicIssues.length) {
      const all = await fetchJiraIssues(auth, 'labels = "CDEF"');
      epicIssues = all.filter(
        (i) => ((i.fields as Record<string,unknown>)?.issuetype as Record<string,unknown>)?.name?.toString().toLowerCase() === "epic"
      );
    }
    if (!epicIssues.length) {
      return NextResponse.json({ error: "No CDEF Epics found" }, { status: 404 });
    }

    // Build epic map
    const epicKeys: string[] = [];
    const epicInfoMap: Record<string, { epic_name: string; project_key: string; project_name: string }> = {};
    const projectsWithEpics = new Set<string>();

    for (const issue of epicIssues) {
      const key = issue.key as string;
      const f = issue.fields as Record<string, unknown>;
      const project = (f.project as Record<string, unknown>) || {};
      const projectKey = (project.key as string) || "UNKNOWN";
      epicKeys.push(key);
      epicInfoMap[key] = {
        epic_name: (f.summary as string) || key,
        project_key: projectKey,
        project_name: (project.name as string) || "",
      };
      projectsWithEpics.add(projectKey);
    }

    // AGP fallback
    if (!projectsWithEpics.has("AGP")) {
      const agpEpics = await fetchJiraIssues(auth, 'project = AGP AND labels = "CDEF" AND (type = Epic OR issuetype = Epic)');
      for (const e of agpEpics) {
        const key = e.key as string;
        if (!epicInfoMap[key]) {
          const f = e.fields as Record<string, unknown>;
          const project = (f.project as Record<string, unknown>) || {};
          epicKeys.push(key);
          epicInfoMap[key] = {
            epic_name: (f.summary as string) || key,
            project_key: (project.key as string) || "AGP",
            project_name: (project.name as string) || "",
          };
          projectsWithEpics.add((project.key as string) || "AGP");
        }
      }
    }

    // Step 2: Fetch child stories per project
    const epicsByProject: Record<string, string[]> = {};
    for (const ek of epicKeys) {
      const proj = ek.split("-")[0];
      if (!epicsByProject[proj]) epicsByProject[proj] = [];
      epicsByProject[proj].push(ek);
    }

    const storyIssues: Record<string, unknown>[] = [];
    for (const [proj, projEpicKeys] of Object.entries(epicsByProject)) {
      const parentFilter = projEpicKeys.map((ek) => `parent = ${ek}`).join(" OR ");
      const epicLinkFilter = projEpicKeys.map((ek) => `"Epic Link" = ${ek}`).join(" OR ");
      const jql = `project = ${proj} AND ((${parentFilter}) OR (${epicLinkFilter}))`;
      const issues = await fetchJiraIssues(auth, jql);
      storyIssues.push(...issues);
    }

    // Step 3: Process stories and fetch worklogs
    const rows: Record<string, unknown>[] = [];
    for (const issue of storyIssues) {
      const key = issue.key as string;
      const f = issue.fields as Record<string, unknown>;

      let epicKey: string | null = null;
      const parent = f.parent as Record<string, unknown> | null;
      if (parent) epicKey = parent.key as string;
      if (!epicKey) epicKey = f.customfield_10014 as string | null;
      if (!epicKey || !epicInfoMap[epicKey]) continue;

      const epicInfo = epicInfoMap[epicKey];
      const project = (f.project as Record<string, unknown>) || {};
      const estSecs = f.timeoriginalestimate as number | null;
      const estHours = estSecs ? Math.round((estSecs / 3600) * 100) / 100 : 0;
      const hoursLogged = await getWorklogHours(auth, key);
      const hrsLeft = Math.max(0, Math.round((estHours - hoursLogged) * 100) / 100);

      rows.push({
        epic_name: epicInfo.epic_name,
        story_name: `${epicInfo.epic_name}-${f.summary as string}`,
        project_key: (project.key as string) || epicInfo.project_key,
        labels: "CDEF",
        project_name: (project.name as string) || epicInfo.project_name,
        original_estimate: estHours,
        total_hours_logged: hoursLogged,
        hrs_left: hrsLeft,
      });
    }

    return NextResponse.json({ rows, epicCount: epicKeys.length, storyCount: rows.length });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
