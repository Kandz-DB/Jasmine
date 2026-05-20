# Jasmine — AI Scheduling Agent
## Risk 2 Solution Group

### What is this?
Jasmine is an Express.js server that:
- Watches training@risk2solution.com every 10 minutes for new emails
- Identifies known client emails by domain (VA, V/Line, DEECA, Wayss, Scentregroup)
- Passes them to Claude (Anthropic API) for processing
- Creates Outlook draft emails (client + trainer) automatically
- Creates calendar events when Diane approves
- Serves the Jasmine HTML interface at /

### Setup

1. Copy .env.example to .env and fill in your values
2. Run: npm install
3. Run: npm start
4. Open: http://localhost:3000
5. Login with the password in APP_PASSWORD

### Same Azure App Registration as Ariel
Uses the same client_id, tenant_id, client_secret from your Ariel Azure app.
Just add training@risk2solution.com access to the existing app registration.
Permissions needed: Mail.ReadWrite, Calendars.ReadWrite, User.Read

### Deploy to Azure App Service
Same process as Ariel:
1. Create Azure App Service (Node.js 18+)
2. Set environment variables in Configuration
3. Deploy via GitHub or Azure CLI
4. Set the URL as jasmine.risk2solution.com via custom domain

### Files
- server.js — main Express server
- public/index.html — Jasmine HTML interface
- jasmine_prompt.txt — Jasmine's AI system prompt
- package.json — dependencies
- .env.example — environment variable template
