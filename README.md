# Ionto Live Dashboard — PowerPoint Content Add-in

A minimal PowerPoint **Content Add-in** (`xsi:type="ContentApp"`) that embeds the
live Ionto Streamlit dashboard in an iframe on a slide, so it updates in real
time during a presentation — including in actual Slide Show mode.

## Why hand-written instead of `yo office`

The task called for scaffolding this with `yo office` (`generator-office`).
I tried that first, but as of `generator-office@3.0.2` the Yeoman generator's
project-type menu only offers Task Pane / SSO / Custom Functions projects —
**Content Add-ins have been dropped from the scaffolding tool**, even though
Microsoft's own docs and manifest schema still fully support
`xsi:type="ContentApp"` for PowerPoint (verified against the current
[Content Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/content-add-ins)
doc and a real content add-in manifest from Microsoft's official samples repo).
So `manifest.xml` below was written directly from that schema and validated
with Microsoft's own validator:

```
npx office-addin-manifest validate manifest.xml
```

It passes with "The manifest is valid." and lists PowerPoint (Windows, Mac,
web, iPad) as supported targets.

## Files

- `manifest.xml` — the Content Add-in manifest for PowerPoint (`Host
  Name="Presentation"`).
- `index.html` — the iframe wrapper page that loads `office.js` and embeds
  the dashboard.
- `assets/icon-32.png`, `assets/icon-64.png`, `assets/icon-80.png` — add-in
  icons referenced by the manifest.

## Changing the embedded URL

Two ways, no redeploy required for the second:

1. **Permanent default** — edit `DEFAULT_DASHBOARD_URL` at the top of the
   `<script>` block in `index.html`, then redeploy.
2. **Per-slide override** — append `?src=<url-encoded-url>` to the add-in's
   `SourceLocation` in `manifest.xml` (or when embedding manually). Example:

   ```
   https://matteomeisterdev.github.io/web-viewer/index.html?src=https%3A%2F%2Fionto-dashboard.fly.dev%2F%3Fpage%3Dmicrofluidic%26mf_run%3D36
   ```

   Only `https://` URLs are accepted; anything else silently falls back to
   the default.

## Hosting: GitHub Pages vs. the Fly app

Two static files (`index.html`, `manifest.xml`) plus three tiny icons need
HTTPS hosting. Two realistic options:

| Option | Effort | Tradeoff |
|---|---|---|
| **GitHub Pages** (recommended, in use) | Push this folder to a repo, flip on Pages in settings. Done. | Fully decoupled from the Streamlit deploy — editing the add-in never touches the running dashboard container, and there's no Dockerfile/static-route wiring to get right. Free, HTTPS by default. **Caveat:** GitHub Pages requires a *public* repo unless you're on GitHub Pro/Team/Enterprise — `matteomeisterdev/web-viewer` was flipped from private to public for this reason. Nothing sensitive lives in it (just the iframe wrapper pointing at the already-public dashboard URL). |
| **Static route on the existing Fly app** | Add a route/static-file handler to the Streamlit container (Streamlit itself doesn't serve arbitrary static files well — you'd need a small nginx sidecar or a separate Fly process in the same app) and redeploy the whole app for any add-in tweak. | Only worth it if you want everything under one domain/deploy pipeline, or need the files to stay in a private repo; otherwise it's strictly more work for no benefit here. |

**Live at:**

```
https://matteomeisterdev.github.io/web-viewer/index.html
https://matteomeisterdev.github.io/web-viewer/manifest.xml
```

`manifest.xml` already points at these URLs (`AppDomain`, `IconUrl`,
`HighResolutionIconUrl`, `SourceLocation`) — no placeholder swap needed.

## Redeploying after edits

GitHub Pages redeploys automatically on every push to `main` (usually live
within ~30–60 seconds):

```bash
git add -A
git commit -m "Update dashboard embed"
git push
```

PowerPoint fetches `manifest.xml` fresh each time you insert the add-in, and
`index.html` is fetched live inside the iframe host on every slide load — no
need to re-sideload unless you change the manifest's `Id`, icons, or
`AppDomains`/`SourceLocation` values.

## Streamlit iframe config — double-check this

By default Streamlit does **not** send `X-Frame-Options` or a restrictive
`frame-ancestors` CSP, so iframing should just work. Before the real demo,
confirm your Fly deployment hasn't added either header (e.g. via a reverse
proxy in front of Streamlit, or `st.set_page_config` extensions):

```bash
curl -sI "https://ionto-dashboard.fly.dev/?page=microfluidic&mf_run=35&mf_view=events" \
  | grep -i "x-frame-options\|content-security-policy"
```

No output = safe to iframe. If either header shows up and blocks framing,
either remove it at the proxy, or (Streamlit config) make sure
`server.enableXsrfProtection`/any custom headers middleware isn't injecting
`frame-ancestors 'none'` / `SAMEORIGIN` restricting the parent to something
that excludes PowerPoint's add-in host process.

If the dashboard is blocked, `index.html`'s fallback UI appears after 6
seconds instead of showing a blank box, with a clickable link to the target
URL.

## Sideloading into PowerPoint Desktop

No Office Store submission needed — this is "sideloading" a manifest file
directly.

**Important:** this is *not* the same as `Tools → Add-Ins` (the legacy
dialog that lists things like `SaveAsAdobePDF` with a `+`/`−` file browser —
that manager only loads old native/VBA add-in packages and cannot load a
web add-in's `manifest.xml` at all). Use the paths below instead.

### Windows

1. Open PowerPoint desktop.
2. **Insert** tab (ribbon) → **Add-ins** → **My Add-ins**.
3. In the dialog, look for **Upload My Add-in** (small link, usually
   top-right of the "My Add-ins" panel).
4. Browse to your local copy of `manifest.xml` → **Upload**.
5. The dashboard content add-in appears; insert it onto a slide from the
   ribbon or the My Add-ins list, then resize/position as needed.

If your build/tenant hides **Upload My Add-in**, use a
[network shared folder catalog](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
instead (Windows-only method): create/share a folder, drop `manifest.xml`
into it, add that folder as a trusted catalog under
**File → Options → Trust Center → Trust Center Settings → Trusted Add-in
Catalogs**, restart PowerPoint, then find the add-in under **Insert →
Add-ins → My Add-ins → SHARED FOLDER**.

### Mac

The **Upload My Add-in** dialog is not available on PowerPoint for Mac.
Mac sideloading instead works by dropping the manifest into a special
folder that PowerPoint watches ([source: Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)):

1. Open **Finder**.
2. Press **Cmd+Shift+G** ("Go to folder").
3. Enter this path (create the `wef` folder if it doesn't exist):

   ```
   /Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   ```

4. Copy your local `manifest.xml` into that `wef` folder.
5. Open (or restart) PowerPoint and open any presentation.
6. Go to **Home** tab → **Add-ins**, and select the dashboard add-in from
   the menu that appears (it's listed by the `DisplayName` from the
   manifest — "Ionto Live Dashboard").
7. Insert it onto the current slide and resize/position as needed.

To remove it later, delete the manifest from the `wef` folder and clear
the Office cache (`~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/`),
then restart PowerPoint.

## Testing in actual Slide Show mode

Content add-ins render as objects embedded in the slide itself (unlike a
task pane, which lives outside the slide canvas), so they should keep
rendering when you start the show. To verify:

1. Insert the add-in onto a slide in Normal/Edit view and confirm the
   dashboard renders live.
2. Start **Slide Show** (F5 or "From Beginning").
3. Navigate to the slide with the add-in — the iframe should still be live
   and interactive (Streamlit auto-refresh, etc. should keep working).
4. If it goes blank in Slide Show but works in Edit view, it's most likely
   the 6-second fallback catching a CSP/X-Frame-Options block that only
   manifests differently under the Slide Show renderer — re-run the `curl`
   header check above and check PowerPoint's version (older builds have
   weaker content add-in support in Slide Show; Microsoft 365 current
   channel is most reliable).
