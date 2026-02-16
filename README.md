# Hotel Risk Advisor Bot

Telegram bot for HUB International's Hotel Franchise Practice. Queries Airtable-based Sales and Consulting databases for claims data, task management, and executive reporting.

## Commands

| Command | Description |
|---------|-------------|
| `/start` | Welcome message |
| `/help` | Show available commands |
| `/consulting query` | Search Consulting System (Incidents/Claims) |
| `/report query` | Generate executive PDF report |
| `/sales query` | Search Sales System |
| `/update` | Get task list |
| `/status` | View progress |
| `/add Client \| Task \| Priority` | Add task |

## Consulting Query Examples

```
/consulting Jasmin                              # All claims
/consulting Jasmin open                         # Open claims only
/consulting Jasmin closed liability             # Closed liability claims
/consulting Jasmin open greater than 25000      # Open claims > $25K
/consulting Jasmin last 5 years                 # Claims from last 5 policy years
/consulting Ocean Partners closed property      # Closed property claims
```

## Environment Variables

| Variable | Description |
|----------|-------------|
| `TELEGRAM_TOKEN` | Telegram Bot API token |
| `AIRTABLE_PAT` | Airtable Personal Access Token |

## Deployment (Railway)

1. Push this repo to GitHub
2. Connect to Railway.app
3. Set environment variables in Railway dashboard
4. Railway auto-deploys from the Procfile
