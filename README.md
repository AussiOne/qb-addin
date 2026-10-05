# Quickbase Live Dashboard – PowerPoint add-in

A PowerPoint **content add-in** that sits on a slide as a box you can move and resize, and shows a live Quickbase dashboard inside it. Each box keeps its own settings, and the data reloads whenever the presentation opens or the slideshow starts.

Works in **PowerPoint for Windows, Mac and the web** (Microsoft 365).

## Features

| | |
|---|---|
| **Move / resize the box** | PowerPoint's own handles. It's a normal object on the slide. |
| **Zoom (browser-style, default)** | `−` / `+` / `0` work like Ctrl +/− in a browser: Quickbase re-lays out the page, so charts and tables grow and refit the box. It uses the same zoom steps as Chrome/Edge (25%–500%). Layout is identical in edit view and slideshow. |
| **Magnify mode (optional)** | Choose it in ⚙ › *Zoom behaves like*. The page is laid out at a fixed size and scaled like a picture, with *Fit width*, *Fit whole area*, *Fill frame* and *Pan*. |
| **Crop** | `✂ Crop` shows the whole page, and you drag a rectangle around the part you want. `⊘` removes the crop. You can also type exact values in ⚙ Settings. |
| **Window width** | How wide the box acts as a browser window at 100% (default 1600px). Bigger = more columns at the same zoom. |
| **Scrolling** | In browser mode, scroll inside the dashboard as you would in a browser. |
| **Live data** | Reloads every time the add-in loads, when the slideshow starts, and on an optional auto-refresh timer. Refreshes load in the background, so the slide never goes blank. |
| **Snap position** | **Position** (3×3 grid icon) › 3×3 grid snaps the content to any edge, corner or the centre when it doesn't fill the box (for example after cropping). |
| **Background colour** | **Position** (3×3 grid icon) › colour picker, or the **eyedropper (Sample)**: click it, then click the dashboard's edge to copy that colour (Windows/web, Chromium-based). **Reset** goes back to white. |
| **Enlarge while editing** | If the box is small, opening Settings, Position or Crop (or ☰) makes the box bigger on the slide, then puts it back when you're done. **⤢ Enlarge** does this on demand, and the green **✓ Done** button shrinks it back. Needs a current Microsoft 365 PowerPoint (PowerPointApi 1.5). |
| **Small boxes** | The toolbar shrinks to icons only, then to a single ☰ button. Where enlarging isn't supported, ☰ opens a drop-down menu instead. |
| **Lock** | Stops clicks and scrolling inside the dashboard during a talk. |
| **Per-slide settings** | Stored inside the .pptx, so save the file after you configure it. |

The toolbar shows up when you move the mouse to the top edge of the box. It hides during the slideshow unless you turn it on in Settings.

---

## 1. Check Quickbase first (2 minutes)

1. **Embedding must be allowed.** A Quickbase realm admin must make sure the realm security policy **"Prevent embedding in iframes"** is **off**. If it's on, the box stays blank. No add-in can get around this.
2. **Get your dashboard URLs.** Open each dashboard in your browser and copy the address bar URL.

## 2. Host the files (free, about 5 minutes)

Office add-ins must be served over **HTTPS**. Any static host works. GitHub Pages is the easiest:

1. Create a GitHub repo, for example `qb-addin`.
2. Upload the **contents of the `src/` folder** to the repo root (`index.html`, `app.js`, `styles.css`, `dialog.html`, `assets/`).
3. Go to repo **Settings › Pages**, pick Source = *Deploy from branch*, choose `main` and `/ (root)`.
4. Your add-in URL is `https://<your-github-name>.github.io/qb-addin/index.html`.

(Azure Static Web Apps, Netlify, SharePoint-hosted files and an internal IIS server also work.)

## 3. Create your manifest

Needs Node.js:

```bash
node configure.js https://<your-github-name>.github.io/qb-addin <yourrealm>
```

This writes `manifest.prod.xml`. `<yourrealm>` is the part before `.quickbase.com`. Without Node, open `manifest.xml` and replace `https://YOUR-HOST/qb-addin` and `YOURREALM` by hand.

## 4. Install in PowerPoint

### PowerPoint on the web, or any version signed in with a Microsoft 365 account
**Insert › Add-ins › My Add-ins › Upload My Add-in** and choose `manifest.prod.xml`.
(In some builds this is under **Home › Add-ins › More Add-ins › My Add-ins › Upload My Add-in**.)

### PowerPoint desktop on Windows (shared-folder catalog)
1. Make a folder, for example `C:\OfficeAddins`, and copy `manifest.prod.xml` into it.
2. Right-click the folder › **Properties › Sharing › Share…** and copy the network path (`\\YOURPC\OfficeAddins`).
3. In PowerPoint: **File › Options › Trust Center › Trust Center Settings › Trusted Add-in Catalogs**. Paste the path, click **Add catalog**, tick **Show in Menu**, click OK, and restart PowerPoint.
4. **Insert › Add-ins › My Add-ins › SHARED FOLDER** › *Quickbase Live Dashboard*.

### PowerPoint on Mac
Copy the manifest to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (create the `wef` folder if it isn't there), restart PowerPoint, then use **Insert › Add-ins › My Add-ins**.

### Whole organization
A Microsoft 365 admin can deploy `manifest.prod.xml` to everyone from **Microsoft 365 admin center › Settings › Integrated apps › Upload custom apps**. That way nobody needs to sideload.

## 5. Build the deck

1. On a slide: **Insert › My Add-ins › Quickbase Live Dashboard**. A box appears.
2. Click **Set up dashboard**, paste the URL, and click **Save & load**.
3. If you see a Quickbase sign-in page, open **⚙ › Sign in to Quickbase…**, sign in, then close that window. The box reloads.
4. Resize or move the box with PowerPoint's handles. Then use **Fit W / Fit / ✂ Crop / ✥ Pan / zoom** to frame the content.
5. To add more dashboards, **duplicate the slide** (the copy keeps the settings) and change the URL in ⚙. Or insert the add-in again.
6. **Save the .pptx.**

Next time you open the deck, or start the slideshow, every box loads current data.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| Blank box, or "refused to connect" | The realm's *Prevent embedding in iframes* policy is on. Ask your Quickbase admin. |
| Sign-in page inside the box, or SSO login won't load | Use **⚙ › Sign in to Quickbase…**. SSO providers (Okta, Entra ID, etc.) block being shown inside frames, so sign-in has to happen in its own window. |
| Works on desktop but asks for sign-in on PowerPoint web | The browser is blocking third-party cookies. Allow cookies for `[*.]quickbase.com` in your browser's site settings. Safari blocks them by default, so use Edge or Chrome. |
| Shows old data | Press ⟳, or set an auto-refresh interval. Check that *Reload when the slideshow starts* is ticked. |
| Settings disappear | The .pptx wasn't saved after you configured the box. |
| Add-in won't load after you changed hosted files | Clear the Office cache. On Windows, delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. On the web, hard-refresh the page. |

## Local development (optional)

```bash
npm run certs                                  # one-time: trusted localhost HTTPS cert
node configure.js https://localhost:3000 <yourrealm>   # writes manifest.dev.xml (separate "(Dev)" add-in)
npm start                                      # serves src/ at https://localhost:3000
```

Sideload `manifest.dev.xml` using the steps above. You can also open `https://localhost:3000/index.html` in a normal browser to test the UI. Outside PowerPoint, settings are saved to localStorage.

## Files

```
manifest.xml      template manifest (placeholders)
configure.js      fills in host + realm -> manifest.prod.xml / manifest.dev.xml
serve.js          local HTTPS server for development
src/index.html    the add-in UI
src/app.js        zoom / crop / pan / refresh / settings logic
src/styles.css
src/dialog.html   sign-in window launcher (forwards to *.quickbase.com only)
src/assets/       icons
```
