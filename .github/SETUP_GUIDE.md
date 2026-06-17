# Bright Metalwork — Android App Setup Guide

## Method 1: GitHub Actions (Recommended — no Android Studio needed)

### What happens
Every time you push to GitHub, GitHub's servers automatically build the APK 
and make it downloadable. You never need Android Studio or any build tools.

### Steps

**Step 1: Add these files to your GitHub repo**

Copy the following into your repo root (same level as src/):
- `.github/workflows/build-android.yml`  ← the build workflow
- `capacitor.config.ts`
- `vite.config.js` (replace existing)
- `package.json` (replace existing)

**Step 2: Push to GitHub**

Either via GitHub Desktop or drag-and-drop upload on GitHub.

**Step 3: Watch the build**

1. Go to your repo on GitHub
2. Click the **Actions** tab at the top
3. You'll see "Build Android APK" running (takes about 5-8 minutes)
4. Click on it to watch the progress

**Step 4: Download your APK**

1. Once the build finishes (green tick ✅)
2. Click on the completed workflow run
3. Scroll to the bottom — you'll see **"Artifacts"**
4. Click **"bright-metalwork-debug-apk"** to download a zip
5. Unzip it — inside is `app-debug.apk`

**Step 5: Install on Android phone**

1. Send the APK file to your phone (email, WhatsApp, Google Drive)
2. Open it on your phone
3. You'll be asked to "Allow from this source" — tap Allow
4. Tap Install
5. Done — Bright Metalwork appears in your app drawer

---

## Method 2: Build locally with Android Studio

### What you need
- Android Studio (free): https://developer.android.com/studio

### Steps

1. Install Android Studio
2. Open Terminal in this project folder
3. Run:
   ```
   npm install
   npm run build
   npx cap add android
   npx cap sync android
   ```
4. Open Android Studio → Open Project → select the `android/` folder
5. Wait for Gradle sync to finish
6. Click **Build → Build Bundle(s) / APK(s) → Build APK(s)**
7. APK saved to: `android/app/build/outputs/apk/debug/app-debug.apk`

---

## Distributing to your team (without app store)

### Option A: Direct download link
Upload the APK to Google Drive and share the link.
Workers tap the link on their Android → Install.

### Option B: WhatsApp / Email
Send the APK file directly. Workers open it to install.

### Option C: Firebase App Distribution (free)
1. Go to console.firebase.google.com
2. Create a project → App Distribution
3. Upload your APK
4. Invite workers by email — they get a link to install

---

## For iPhone (iOS) — different process

Android APKs don't run on iPhone. For iOS you need:
1. A Mac computer
2. Xcode (free from Mac App Store)
3. Apple Developer account ($99/year) to distribute outside TestFlight
4. Run: `npx cap add ios && npx cap open ios`

The cheapest iPhone option: PWA (installable from Safari, no app store needed)
See the manifest.json and vite.config.js with VitePWA for that setup.

---

## What the app has on Android (vs web)
- ✅ All features identical to web version
- ✅ Supabase database — same data, live sync
- ✅ GPS sign-in works with native accuracy
- ✅ Works offline (cached web view)
- ✅ Push notifications (route changes, approvals)
- ✅ Camera access for certificate photos
- ✅ Installs to home screen like a real app
- ✅ Runs without browser chrome (full screen)
