import os
import requests
from datetime import datetime, timezone
from msal import PublicClientApplication, SerializableTokenCache

# ── CONFIG ───────────────────────────────────────────────────────────────────
CLIENT_ID        = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
TENANT_ID        = "8619c67c-945a-48ae-8e77-35b1b71c9b98"
TEAM_ID          = "7a748f03-3b67-4bd2-992b-fda5a8ac4c30"
CHANNEL_ID       = "19:8c4730fcee4a47b7b7346ff685504431@thread.tacv2"
TOKEN_CACHE_FILE = r"C:\Scripts\msal_token_cache.bin"

JIRA_BASE   = "https://axso-tim.atlassian.net"
JIRA_COOKIE = "tenant.session.token=eyJraWQiOiJzZXNzaW9uLXNlcnZpY2UvcHJvZC0xNzM4Nzk0ODc0IiwiYWxnIjoiUlMyNTYifQ.eyJhc3NvY2lhdGlvbnMiOltdLCJzdWIiOiI3MTIwMjA6Yjg0NmU1ZTAtZDk0Ni00NjIxLTgxYjEtYzcyNGMyNGMxZTg2IiwiZW1haWxEb21haW4iOiJheHBvLmNvbSIsImltcGVyc29uYXRpb24iOltdLCJjcmVhdGVkIjoxNzgyNTMyMTQyLCJyZWZyZXNoVGltZW91dCI6MTc4MjUzMzY5MiwidmVyaWZpZWQiOnRydWUsImlzcyI6InNlc3Npb24tc2VydmljZSIsInNlc3Npb25JZCI6IjhiZGExMGVjLWM5ZGEtNDRhOS05YTk0LTg4OTAzZTc1M2RkNyIsInN0ZXBVcHMiOltdLCJvcmdJZCI6ImE2Njc5OWIzLWJkYmotMTQ2MS1qYWtiLWs4NWE1Y2Rra2Q2YSIsImF1ZCI6ImF0bGFzc2lhbiIsIm5iZiI6MTc4MjUzMzA5MiwiZXhwIjoxNzgyNjE5NDkyLCJpYXQiOjE3ODI1MzMwOTIsImVtYWlsIjoibmFpZHUucmFnaHVuYXRoQGF4cG8uY29tIiwianRpIjoiOGJkYTEwZWMtYzlkYS00NGE5LTlhOTQtODg5MDNlNzUzZGQ3In0.IRXAJ_qlaQBVmq7UNGREu5Hg9s2L8h34cRRfFkUExbi7iXbpmZXIB0U_qX4rm9hMFFwjnlyL-0LLcL17VWB4K4MELgpeWcLnAEo-p-_Bo9_dXWRZr9X8lRYUfuyAJOKQhF8lRWBoJt5fbmy38JJ7Lai7C8eFce6f8caMmbZxU_caxiqLMec-DSbiFhk4MRF9cwe2z879kKO8902y35abSgGrewwJdnsaabYITgFmjL1YJtA7XsX1zs_SBaa9FA7mFI8TwpDccTL56QBRmOtiEnDZVKgv8-3D0x41Nh2fF2CFOSgvQb__o1g5msRvNJ65ZT7gZJgTDbOq_A9bBBxDxA"

SCOPES = ["https://graph.microsoft.com/ChannelMessage.Send"]

# Alert when remaining time is under this threshold
WARNING_BUFFER_HOURS = 40

# ── MSAL TOKEN CACHE ──────────────────────────────────────────────────────────
def load_cache():
    cache = SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())

def get_teams_token():
    cache = load_cache()
    app = PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result:
        print("No cached token. Starting device flow login...")
        flow = app.initiate_device_flow(scopes=SCOPES)
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)
    save_cache(cache)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Token error: {result.get('error_description')}")

# ── JIRA ──────────────────────────────────────────────────────────────────────
def get_jira_tickets():
    resp = requests.post(f"{JIRA_BASE}/rest/api/3/search/jql",
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Cookie": JIRA_COOKIE,
            "X-Atlassian-Token": "no-check"
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

    # Read actual Jira SLA — customfield_10306 = Time to resolution
    sla_field    = fields.get("customfield_10306", {})
    ongoing      = sla_field.get("ongoingCycle", {})
    breached     = ongoing.get("breached", False)
    goal_millis  = ongoing.get("goalDuration", {}).get("millis", 0)
    rem_millis   = ongoing.get("remainingTime", {}).get("millis", 0)
    rem_friendly = ongoing.get("remainingTime", {}).get("friendly", "N/A")
    goal_friendly= ongoing.get("goalDuration", {}).get("friendly", "N/A")
    elapsed_friendly = ongoing.get("elapsedTime", {}).get("friendly", "N/A")

    remaining_hours = rem_millis / 3600000

    # Also flag if breached in completedCycles
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
def send_teams_alert(access_token, alerts):
    url = f"https://graph.microsoft.com/v1.0/teams/{TEAM_ID}/channels/{CHANNEL_ID}/messages"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    breached = [a for a in alerts if a["breached"]]
    warnings = [a for a in alerts if a["warning"]]

    def ticket_block(a):
        color = "attention" if a["breached"] else "warning"
        status_text = "🔴 BREACHED" if a["breached"] else "🟡 WARNING"
        remaining_str = "OVERDUE" if a["breached"] else a["remaining"]
        return {
            "type": "Container",
            "style": color,
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
                                    "text": f"**{a['key']}** — {a['summary'][:100]}",
                                    "wrap": False,
                                    "weight": "Bolder",
                                    "size": "Small"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": f"👤 {a['assignee']}  |  🎯 Priority: {a['priority']}  |  ⏱ SLA: {a['sla_goal']}",
                                    "wrap": True,
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
                                    "color": "Attention" if a["breached"] else "Warning"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": remaining_str,
                                    "size": "Small",
                                    "isSubtle": True,
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
        "body": {
            "contentType": "html",
            "content": "<attachment id=\"adaptivecard1\"></attachment>"
        },
        "attachments": [
            {
                "id": "adaptivecard1",
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": str(adaptive_card).replace("'", '"').replace("True", "true").replace("False", "false").replace("None", "null")
            }
        ]
    }

    resp = requests.post(url, headers=headers, json=payload)
    resp.raise_for_status()
    print(f"✅ Teams Adaptive Card sent — {len(alerts)} ticket(s) flagged.")

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
    token = get_teams_token()
    send_teams_alert(token, alerts)

if __name__ == "__main__":
    main()
