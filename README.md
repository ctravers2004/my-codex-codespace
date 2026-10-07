# my-codex-codespace

A simple counter app, plus a pelican illustration and sample PowerPoint slides.

## Counter app

Open `counter-app/dist/index.html` in a browser to use the counter locally.
It includes increase, decrease, and reset buttons. Reloading resets the count.

## Publish with GitHub Pages

1. Open this repository on GitHub and go to **Settings → Pages**.
2. Under **Build and deployment**, select **Deploy from a branch**.
3. Choose the **main** branch and **/docs** folder, then click **Save**.
4. Wait for the Pages deployment to finish. GitHub will show the published URL:
   `https://ctravers2004.github.io/my-codex-codespace/`.

The `docs` folder contains the ready-to-publish app. After changing the app,
copy `counter-app/dist/index.html` to `docs/index.html` and push the changes.

## Files

- `pelican-on-a-bike.svg` and `pelican-on-a-bike.pdf`: pelican illustration.
- `sample-slides/sample-powerpoint.pptx`: six sample slides.
- `sample-slides/build.py`: slide generation script.
- `sample-slides/pelican.png`: illustration used in the slides.
- `counter-app/dist/index.html`: app source.
- `counter-app/.openai/hosting.json`: existing Sites hosting configuration.
