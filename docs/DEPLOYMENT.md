# Deployment and Operations

## Local Python run

Requirements: Python 3.12 is recommended; openpyxl is the only package listed in requirements.txt.

From the project directory:

    pip install -r requirements.txt
    python platform_server.py 8780

Open http://127.0.0.1:8780. On Windows, start_dataprism.bat performs the local launch on port 8785 and opens the browser.

The application must be served by platform_server.py. Opening platform.html directly with a file URL will prevent its relative API requests from reaching the backend.

## Docker Compose

From the project directory:

    docker compose up --build

Open http://127.0.0.1:8780. Compose mounts a named volume at /app/data, so the local SQLite database persists across container restarts. Stop the service with Ctrl+C; to run it in the background, use docker compose up -d.

## Render Blueprint

The repository is configured for Render’s Docker runtime:

- Blueprint: render.yaml
- Image build: Dockerfile
- Service name: dataprism
- Health check: /api/health
- Host binding: 0.0.0.0
- Port: the PORT environment variable supplied by Render
- SQLite path in the current Blueprint: /tmp/dataprism/platform.sqlite3

### Initial deployment

1. Push the project to GitHub.
2. In Render, create a Blueprint and connect the repository containing render.yaml.
3. Review the proposed web service and deploy it.
4. Open the service’s URL after the deploy is marked Live.
5. Visit /api/health on that host; a healthy response includes ok: true and storage information.

### Publish a code update

1. Save the changed files in the repository.
2. Commit and push them to the branch connected to the Render service.
3. Watch the service’s Deploys page until the latest commit is Live.
4. Refresh the site. If it still serves an older version, confirm the deployed commit and perform a hard refresh.

A dashboard manual deploy can deploy the latest linked-branch commit. Avoid deploying a specific older commit unless that is intentional.

## Persistence and demo safety

The Render Blueprint currently puts SQLite under /tmp. That filesystem is ephemeral on the free service plan, so user accounts, sessions, and uploaded datasets can be lost on restart, sleep, or redeploy. Do not treat the hosted demo as durable storage.

For durable hosted storage, attach a persistent disk or move the database to a managed database and update DATAPRISM_DB accordingly. That requires a deployment and storage decision beyond the current demo configuration. Keep local Docker data in the named volume when you need it between restarts.

Use sample or non-sensitive datasets in a public portfolio deployment. The app is designed for small exploratory data, not a hardened multi-tenant production analytics service.

## Troubleshooting

| Symptom | First check |
| --- | --- |
| Deploy failed during image build | Open that deploy’s build logs. Check Dockerfile, dependency installation, and repository contents. |
| Service is live but page shows an error | Open the service’s runtime Logs and Events. Check whether the process started and whether health checks pass. |
| Render loading screen or slow first response | A free service may be waking after inactivity. Wait for it to start, then refresh. |
| 502 or health-check failure | Confirm the server binds to 0.0.0.0 and uses Render’s PORT; inspect startup logs and /api/health. |
| Site displays an older version | Check which Git branch Render tracks and which commit is Live; then hard-refresh the browser. |
| Upload says “Failed to fetch” locally | Start the app with start_dataprism.bat or platform_server.py and use its localhost URL, not a file:// URL. |
| Sign-in or uploaded data disappeared after redeploy | The hosted database is currently on ephemeral /tmp storage. This is expected for the configured demo. |

When asking for help, include the Render service URL (if public), the deploy status, and the relevant error lines from the Render build or runtime logs. Never share passwords, session cookies, or environment secrets.