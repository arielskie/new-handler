# Handler Monitoring Dashboard (Intest 2)

A real-time semiconductor handler monitoring, repair tracking, and spare parts inventory dashboard.

- **Live URL**: [https://arielskie.github.io/new-handler/](https://arielskie.github.io/new-handler/)
- **Power Automate Integration**: Real-time read, write, and delete operations connected to cloud workflows.

---

## Deploying to GitHub Pages (Fixing Blank Screen)

If your GitHub Pages displays a blank white screen, it is because GitHub Pages was trying to serve raw unbundled TypeScript files instead of the compiled build, or asset paths were pointing to the root domain.

We have resolved this with two deployment options:

### Option 1: GitHub Actions (Recommended — Automated)
1. Go to your repository on GitHub: `https://github.com/arielskie12345/new-handler`
2. Navigate to **Settings** &rarr; **Pages**.
3. Under **Build and deployment** &gt; **Source**, select **GitHub Actions**.
4. That's it! GitHub will automatically trigger the included workflow (`.github/workflows/deploy.yml`) on every push to `main` and deploy the application.

---

### Option 2: Deploy from Branch (`main` / `docs`)
If you prefer deploying from the `main` branch without GitHub Actions:
1. Navigate to **Settings** &rarr; **Pages**.
2. Under **Build and deployment** &gt; **Source**, keep **Deploy from a branch**.
3. Set **Branch**: `main` and select folder: **`/docs`** (instead of `/ (root)`).
4. Click **Save**. The pre-bundled production files in `/docs` with relative asset paths will load immediately.

---

## Local Development

```bash
# Install dependencies
npm install

# Start development server
npm run dev

# Build for production (outputs to dist/ and docs/)
npm run build
```
