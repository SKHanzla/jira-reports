import { NextResponse } from "next/server";

const EMAIL = process.env.JIRA_EMAIL!;
const API_TOKEN = process.env.JIRA_USER_API_TOKEN!;
const DOMAIN = process.env.JIRA_DOMAIN!;

const JQL = "worklogDate >= -30d";
const MAX_RESULTS = 100;

const FIELDS = [
  "summary", "project", "issuetype", "status",
  "parent", "customfield_10014", "customfield_10011", "labels", "worklog", "timespent",
];

export interface UserReportRow {
  user: string;
  epic_name: string;
  story_name: string;
  project_key: string;
  label: string;
  project_name: string;
  status: string;
  hours: number;
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

async function getWorklogsByUser(
  auth: string,
  issueKey: string
): Promise<{ user: string; hours: number }[]> {
  const headers = { Authorization: `Basic ${auth}`, Accept: "application/json" };
  try {
    const res = await fetch(`https://${DOMAIN}/rest/api/3/issue/${issueKey}/worklog`, { headers });
    if (!res.ok) return [];
    const data = await res.json();
    return (data.worklogs || [])
      .filter((wl: Record<string, unknown>) => (wl.timeSpentSeconds as number) > 0)
      .map((wl: Record<string, unknown>) => ({
        user: (wl.author as Record<string, unknown>)?.displayName as string || "Unknown",
        hours: Math.round(((wl.timeSpentSeconds as number) / 3600) * 100) / 100,
      }));
  } catch {
    return [];
  }
}

export async function GET() {
  try {
    const auth = getAuth();
    const issues = await fetchAllIssues(auth);
    const epicCache: Record<string, { isEpic: boolean; name: string; labels: string[] }> = {};
    const rows: UserReportRow[] = [];

    for (const issue of issues) {
      const key = issue.key as string;
      const f = issue.fields as Record<string, unknown>;

      const project = (f.project as Record<string, unknown>) || {};
      const projectKey = (project.key as string) || "UNKNOWN";
      const projectName = (project.name as string) || projectKey;
      const itype = ((f.issuetype as Record<string, unknown>)?.name as string || "").toLowerCase();
      const summary = (f.summary as string) || "No Summary";
      const statusName = ((f.status as Record<string, unknown>)?.name as string) || "Unknown";

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
      const worklogs = await getWorklogsByUser(auth, key);
      if (!worklogs.length) continue;

      const userHours: Record<string, number> = {};
      for (const wl of worklogs) {
        userHours[wl.user] = (userHours[wl.user] || 0) + wl.hours;
      }

      for (const [user, hours] of Object.entries(userHours)) {
        rows.push({
          user,
          epic_name: epicName,
          story_name: storyName,
          project_key: projectKey,
          label,
          project_name: projectName,
          status: statusName,
          hours: Math.round(hours * 100) / 100,
        });
      }
    }

    return NextResponse.json({ rows, issueCount: issues.length });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
