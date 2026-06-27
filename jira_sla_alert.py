import os
import json
import requests
from datetime import datetime, timezone

# ── CONFIG ───────────────────────────────────────────────────────────────────
TEAM_ID    = "7a748f03-3b67-4bd2-992b-fda5a8ac4c30"
CHANNEL_ID = "19:8c4730fcee4a47b7b7346ff685504431@thread.tacv2"
JIRA_BASE  = "https://axso-tim.atlassian.net"

# Read from GitHub Secrets
JIRA_EMAIL        = os.environ["JIRA_EMAIL"]
JIRA_API_TOKEN    = os.environ["JIRA_API_TOKEN"]
TEAMS_WEBHOOK_URL = os.environ["TEAMS_WEBHOOK_URL"]

WARNING_BUFFER_HOURS = 40

# ── JIRA ──────────────────────────────────────────────────────────────────────
def get_jira_tickets():
    resp = requests.post(f"{JIRA_BASE}/rest/api/3/search/jql",
        auth=(JIRA_EMAIL, JIRA_API_TOKEN),
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
        json={
            "jql": "project = SR AND labels = MI_DataOps AND statusCategory != Done",
            "fields": ["summary", "priority", "labels", "created", "status", "assignee", "customfield_10306"],
            "maxResults": 50
        },
        timeout=15
    )
    resp.raise_for_status()
    return resp.json().get("issues", [])

def check_sla_breach(ticket):
    fields   = ticket["fields"]
    priority = fields["priority"]["name"]

    assignee_field = fields.get("assignee")
    assignee = assignee_field["displayName"] if assignee_field else "Unassigned"

    sla_field        = fields.get("customfield_10306", {})
    ongoing          = sla_field.get("ongoingCycle", {})
    breached         = ongoing.get("breached", False)
    rem_millis       = ongoing.get("remainingTime", {}).get("millis", 0)
    rem_friendly     = ongoing.get("remainingTime", {}).get("friendly", "N/A")
    goal_friendly    = ongoing.get("goalDuration", {}).get("friendly", "N/A")
    elapsed_friendly = ongoing.get("elapsedTime", {}).get("friendly", "N/A")
    remaining_hours  = rem_millis / 3600000

    completed = sla_field.get("completedCycles", [])
    if not ongoing and completed:
        breached = any(c.get("breached", False) for c in completed)

    return {
        "key":           ticket["key"],
        "summary":       fields["summary"],
        "priority":      priority,
        "assignee":      assignee,
        "sla_goal":      goal_friendly,
        "elapsed":       elapsed_friendly,
        "remaining":     rem_friendly,
        "remaining_hrs": round(remaining_hours, 1),
        "breached":      breached,
        "warning":       not breached and 0 < remaining_hours <= WARNING_BUFFER_HOURS,
    }

# ── TEAMS ─────────────────────────────────────────────────────────────────────
def send_teams_alert(alerts):
    breached = [a for a in alerts if a["breached"]]
    warnings = [a for a in alerts if a["warning"]]

    today_str = datetime.now().strftime('%d %b %Y, %I:%M %p')

    def ticket_block(a):
        color = "attention" if a["breached"] else "warning"
        status_text = "🔴 BREACHED" if a["breached"] else "🟡 WARNING"
        remaining_str = "OVERDUE" if a["breached"] else a["remaining"]

        # Truncate summary more tightly so "KEY — summary" reliably fits on one line
        # even on narrower Teams panes (mobile / side panel)
        max_summary_len = 38
        summary = a["summary"]
        if len(summary) > max_summary_len:
            summary = summary[:max_summary_len - 1].rstrip() + "…"

        return {
            "type": "Container",
            "style": color,
            "bleed": True,
            "spacing": "Small",
            "items": [
                {
                    "type": "ColumnSet",
                    "spacing": "None",
                    "columns": [
                        {
                            "type": "Column",
                            "width": "stretch",
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "text": f"**{a['key']}** — {summary}",
                                    "wrap": False,
                                    "weight": "Bolder",
                                    "size": "Small",
                                    "spacing": "None"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": f"👤 {a['assignee']}  |  🎯 {a['priority']}  |  ⏱ SLA: {a['sla_goal']}",
                                    "wrap": False,
                                    "size": "Small",
                                    "isSubtle": True,
                                    "spacing": "None"
                                }
                            ]
                        },
                        {
                            "type": "Column",
                            "width": "auto",
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "text": status_text,
                                    "weight": "Bolder",
                                    "size": "Small",
                                    "color": "Attention" if a["breached"] else "Warning",
                                    "horizontalAlignment": "Right",
                                    "spacing": "None"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": remaining_str,
                                    "size": "Small",
                                    "isSubtle": True,
                                    "horizontalAlignment": "Right",
                                    "spacing": "None"
                                }
                            ],
                            "horizontalAlignment": "Right"
                        }
                    ]
                }
            ],
            "selectAction": {
                "type": "Action.OpenUrl",
                "url": f"https://axso-tim.atlassian.net/browse/{a['key']}"
            }
        }

    body_blocks = []
    if breached:
        body_blocks.append({
            "type": "TextBlock",
            "text": f"🔴 BREACHED ({len(breached)})",
            "weight": "Bolder",
            "color": "Attention",
            "size": "Medium",
            "spacing": "Medium"
        })
        for a in breached:
            body_blocks.append(ticket_block(a))

    if warnings:
        body_blocks.append({
            "type": "TextBlock",
            "text": f"🟡 WARNING ({len(warnings)})",
            "weight": "Bolder",
            "color": "Warning",
            "size": "Medium",
            "spacing": "Medium"
        })
        for a in warnings:
            body_blocks.append(ticket_block(a))

    adaptive_card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": [
            {
                "type": "Container",
                "style": "emphasis",
                "bleed": True,
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "⚠️ MI DataOps — Jira SLA Alert",
                                        "weight": "Bolder",
                                        "size": "Large",
                                        "wrap": True
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"🕐 {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  📋 {len(alerts)} ticket(s) flagged",
                                        "isSubtle": True,
                                        "size": "Small",
                                        "spacing": "None"
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": f"🔴 {len(breached)} Breached",
                                        "weight": "Bolder",
                                        "color": "Attention",
                                        "size": "Small"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"🟡 {len(warnings)} Warning",
                                        "weight": "Bolder",
                                        "color": "Warning",
                                        "size": "Small",
                                        "spacing": "None"
                                    }
                                ],
                                "horizontalAlignment": "Right"
                            }
                        ]
                    }
                ]
            },
            *body_blocks
        ],
        "actions": [
            {
                "type": "Action.OpenUrl",
                "title": "📋 Open Jira Board",
                "url": "https://axso-tim.atlassian.net/jira/servicedesk/projects/SR/queues"
            }
        ]
    }

    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": adaptive_card
            }
        ]
    }

    resp = requests.post(TEAMS_WEBHOOK_URL, json=payload)
    resp.raise_for_status()
    print(f"✅ Teams alert sent — {len(alerts)} ticket(s) flagged.")

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Checking Jira SLA status...")
    tickets = get_jira_tickets()
    print(f"Tickets fetched: {len(tickets)}")

    alerts = []
    for ticket in tickets:
        result = check_sla_breach(ticket)
        if result["breached"] or result["warning"]:
            alerts.append(result)

    if not alerts:
        print("✅ No SLA breaches or warnings found.")
        return

    print(f"⚠️  {len(alerts)} ticket(s) need attention. Sending Teams alert...")
    send_teams_alert(alerts)

if __name__ == "__main__":
    main()
