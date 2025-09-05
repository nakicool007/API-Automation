from jira import JIRA
import matplotlib.pyplot as plt
import win32com.client as win32
import requests
from requests.auth import HTTPBasicAuth

# ---------------- Jira Configuration ----------------
JIRA_URL = "https://jira.amer.company_name.com"
BOARD_ID = "ID"
JIRA_USERNAME = "email"
JIRA_API_TOKEN = "xxxxxx" 

# Auth for REST calls
auth = HTTPBasicAuth(JIRA_USERNAME, JIRA_API_TOKEN)
headers = {"Accept": "application/json"}

# ---------------- Fetch Velocity Chart Data ----------------
velocity_url = f"{JIRA_URL}/rest/agile/1.0/board/{BOARD_ID}/velocity"
response = requests.get(velocity_url, headers=headers, auth=auth)

if response.status_code != 200:
    raise Exception(f"Failed to fetch velocity data: {response.status_code} - {response.text}")

velocity_data = response.json()

committed = []
completed = []
sprint_names = []

for sprint in velocity_data["sprints"]:
    sprint_id = sprint["id"]
    sprint_name = sprint["name"]
    sprint_names.append(sprint_name)
    
    # velocityStatEntries might be missing some sprint IDs
    entry = velocity_data["velocityStatEntries"].get(str(sprint_id), {})
    committed.append(entry.get("estimated", {}).get("value", 0))
    completed.append(entry.get("completed", {}).get("value", 0))

# ---------------- Fetch Sprint Report Data ----------------
jira = JIRA(server=JIRA_URL, basic_auth=(JIRA_USERNAME, JIRA_API_TOKEN))
sprints = jira.sprints(BOARD_ID)
latest_sprint = sprints[-1]

sprint_report_url = f"{JIRA_URL}/rest/agile/1.0/sprint/{latest_sprint.id}/issue?maxResults=1000"
resp = requests.get(sprint_report_url, headers=headers, auth=auth)

if resp.status_code != 200:
    raise Exception(f"Failed to fetch sprint report: {resp.status_code} - {resp.text}")

sprint_report = resp.json()

# Count completed vs not completed issues
completed_issues = 0
not_completed_issues = 0

for issue in sprint_report.get("issues", []):
    status = issue["fields"]["status"]["name"].lower()
    if status in ["done", "closed", "resolved"]:
        completed_issues += 1
    else:
        not_completed_issues += 1

# ---------------- Plot Velocity Chart ----------------
plt.figure(figsize=(8, 5))
plt.plot(sprint_names, committed, marker="o", label="Committed")
plt.plot(sprint_names, completed, marker="o", label="Completed")
plt.title("Velocity Chart")
plt.xlabel("Sprint")
plt.ylabel("Story Points")
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig("velocity_chart.png")

# ---------------- Send Email ----------------
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "nakul.kumbria@thermofisher.com"
mail.Subject = "Jira Velocity & Sprint Report"
mail.Body = (
    f"Velocity Chart attached.\n\n"
    f"Latest Sprint ({latest_sprint.name}) summary:\n"
    f"- Completed Issues: {completed_issues}\n"
    f"- Not Completed: {not_completed_issues}\n"
)
mail.Attachments.Add("velocity_chart.png")
mail.Send()

print("Email sent successfully with Velocity + Sprint Report data!")
