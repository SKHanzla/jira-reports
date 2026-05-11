import { NextResponse } from "next/server";

const EMAIL = process.env.JIRA_EMAIL!;
const API_TOKEN = process.env.JIRA_USER_API_TOKEN!;
const DOMAIN = process.env.JIRA_DOMAIN!;

const JQL = "worklogDate >= -30d";
const MAX_RESULTS = 100;

const FIELDS = [
  "summary", "project", "issuetype", "status",
  "parent", "customfield_10014", "customfield_10011",
  "labels", "worklog", "timespent", "timeoriginalestimate",
];

export interface ProjectReportRow {
  project_name: string;
  project_key: string;
  epic_name: string;
  story_name: string;
  label: string;
  status: string;
  original_estimate: number;
  hours_logged: number;
  hours_left: number;
}

function getAuth(): string {
  return Buffer.from(`${EMAIL}:${API_TOKEN}`).toString("base64");
}

async function fetchAllIssues(auth: string): Promise<Record<string, unknown>[]> {
  const url = `https://${DOMAIN}/rest/api/3/search/jql`;
  const headers = { Authorization: `Basic ${auth}`, Accept: "application/json" };
  const all: Record<string, unknown>[] = [];
  let startAt = 0;

  while (true) {
    const params = new URLSearchParams({
      jql: JQL,
      fields: FIELDS.join(","),
      maxResults: String(MAX_RESULTS),
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
    all.push(...issues);
    startAt = all.length;
    if (startAt >= total || issues.length === 0) break;
  }
  return all;
}

async function fetchEpicDetails(
  auth: string,
  issueKey: string,
  cache: Record<string, { isEpic: boolean; name: string; labels: string[] }>
): Promise<{ epicName: string; labels: string[] }> {
  if (!cache[issueKey]) {
    const headers = { Authorization: `Basic ${auth}`, Accept: "application/json" };
    try {
      const res = await fetch(
        `https://${DOMAIN}/rest/api/3/issue/${issueKey}?fields=summary,customfield_10011,labels,issuetype`,
        { headers }
      );
      if (!res.ok) {
        cache[issueKey] = { isEpic: false, name: "", labels: [] };
      } else {
        const f = (await res.json()).fields || {};
        const itype = (f.issuetype?.name || "").toLowerCase();
        cache[issueKey] = itype === "epic"
          ? { isEpic: true, name: f.customfield_10011 || f.summary || issueKey, labels: f.labels || [] }
          : { isEpic: false, name: "", labels: [] };
      }
    } catch {
      cache[issueKey] = { isEpic: false, name: "", labels: [] };
    }
  }
  const info = cache[issueKey];
  return info.isEpic ? { epicName: info.name, labels: info.labels } : { epicName: "", labels: [] };
}

async function getTotalHoursLogged(auth: string, issueKey: string): Promise<number> {
  const headers = { Authorization: `Basic ${auth}`, Accept: "application/json" };
  try {
    const res = await fetch(`https://${DOMAIN}/rest/api/3/issue/${issueKey}/worklog`, { headers });
    if (!res.ok) return 0;
    const data = await res.json();
    const totalSecs = (data.worklogs || []).reduce(
      (s: number, wl: Record<string, unknown>) => s + ((wl.timeSpentSeconds as number) || 0),
      0
    );
    return Math.round((totalSecs / 3600) * 100) / 100;
  } catch {
    return 0;
  }
}

export async function GET() {
  try {
    const auth = getAuth();
    const issues = await fetchAllIssues(auth);
    const epicCache: Record<string, { isEpic: boolean; name: string; labels: string[] }> = {};
    const rows: ProjectReportRow[] = [];

    for (const issue of issues) {
      const key = issue.key as string;
      const f = issue.fields as Record<string, unknown>;

      const project = (f.project as Record<string, unknown>) || {};
      const projectKey = (project.key as string) || "UNKNOWN";
      const projectName = (project.name as string) || projectKey;
      const itype = ((f.issuetype as Record<string, unknown>)?.name as string || "").toLowerCase();
      const summary = (f.summary as string) || "No Summary";
      const statusName = ((f.status as Record<string, unknown>)?.name as string) || "Unknown";

      const origSecs = (f.timeoriginalestimate as number) || 0;
      const originalEstimate = origSecs ? Math.round((origSecs / 3600) * 100) / 100 : 0;

      let epicName = "";
      const storyName = summary;
      let labelsArr: string[] = [];

      if (itype === "epic") {
        epicName = (f.customfield_10011 as string) || summary;
        labelsArr = (f.labels as string[]) || [];
      } else {
        let epicKey = f.customfield_10014 as string | null;
        if (!epicKey) {
          const parent = f.parent as Record<string, unknown> | null;
          if (parent) epicKey = parent.key as string;
        }
        if (epicKey) {
          const { epicName: en, labels: ll } = await fetchEpicDetails(auth, epicKey, epicCache);
          epicName = en || summary;
          labelsArr = ll;
        } else {
          epicName = summary;
        }
      }

      const label = labelsArr.join(", ");
      const hoursLogged = await getTotalHoursLogged(auth, key);
      if (hoursLogged === 0) continue;

      const hoursLeft = Math.max(0, Math.round((originalEstimate - hoursLogged) * 100) / 100);

      rows.push({
        project_name: projectName,
        project_key: projectKey,
        epic_name: epicName,
        story_name: storyName,
        label,
        status: statusName,
        original_estimate: originalEstimate,
        hours_logged: hoursLogged,
        hours_left: hoursLeft,
      });
    }

    return NextResponse.json({ rows, issueCount: issues.length });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
