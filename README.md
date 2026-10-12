# Rentbook - phone app for your Rent Tracker sheet

Practice mode works right away: open `index.html`. To connect your real Google Sheet:

## 1. Put the new sheet in Google Sheets (once)
1. Upload `Apartment_Rent_Tracker.xlsx` to Google Drive, open it, then **File > Save as Google Sheets**.
2. In the new Google Sheet choose **File > Settings** and set the **Time zone** to Manila (GMT+8).
3. Clear the example people, payments and the one example rent change (green columns only) when you are ready to start for real.

## 2. Connect the sheet (10 minutes, once)
1. In the Google Sheet click **Extensions > Apps Script**.
2. Delete what is there, paste everything from `apps-script/Code.gs`.
3. Change `CHANGE-ME-1234` to a PIN only your family knows.
4. Click **Deploy > New deployment > Web app**. Execute as: **Me**. Who has access: **Anyone**. Click Deploy and allow access.
5. Copy the **Web app URL**.

Already deployed an older version? Paste the new `Code.gs`, then **Deploy > Manage deployments > Edit > New version**. The URL stays the same.

## 3. Put the app online (GitHub Pages)
1. Open `config.js` and paste the URL between the quotes after `API_URL:`.
2. Create a GitHub repository, upload every file in this folder (except `apps-script`, which is optional), then **Settings > Pages > Deploy from branch > main**.
3. Open the Pages link on each phone and type the PIN once.

## 4. Install it like an app
- **Android (Chrome):** menu (three dots) > Add to Home screen.
- **iPhone (Safari):** Share > Add to Home Screen.

## What the app does
- **Today:** who owes money, who is due soon, and who is all paid up. No totals, just names and what each person owes.
- **Add payment:** pick the tenant, the amount fills in. The next due date is worked out by itself.
- **Change rent:** open a tenant, tap Change rent, type the new rent and the date it starts. Earlier months keep the old rent.
- **Tenants:** add, edit, mark as moved out. Moved-out tenants move to Old tenants, and their payment history stays.
- **History:** look back by year, month and day.
- **Settings:** removing payments is **off** until you turn it on. Also the warning days and the text size for this phone.

## 5. Phone reminders (new)
Already set up the older version? Do NOT redo the sheet. Paste the new `Code.gs` (keep your PIN on line 1), run `setupReminders` once, redeploy a new version, and upload `index.html`, `sw.js` and `push-config.js` (not `config.js`). Then follow Part 2 of the guide to add the free Firebase keys.
- One message a day at most, around 8 AM Manila time, only if someone is late, due today, or due soon (at your warning-days mark and the day before).
- Settings > **Reminders on this phone** turns it on or off for that phone, and has a **Send me a test message** button.
- Settings > **Reminder times**: the hour, and when to stop reminding about someone who is very late (still red in the app).
- Android: works in Chrome. iPhone: iOS 16.4+, and the app must be opened from its Home Screen icon.
- The sheet gets a **Devices** tab by itself. Set a phone's Reminders to No there to stop it.
- Editor helpers: `previewReminder` (shows what would be sent), `sendReminderNow`.

## Notes
- Dates follow Manila time (`TIMEZONE` in `config.js`). The app refreshes when it is opened again after midnight.
- **Screenshot upload:** in Add payment, tap *Upload payment screenshot*. The phone reads the picture itself (free, no account). The first time it downloads a small reading tool, so use Wi-Fi once. Always check the amount, date and reference number before saving.
- The same reference number cannot be saved twice.
- The sheet stays the database: every payment is a new row on the Payments tab, every rent change is a row on Rent Changes. Edit the sheet by hand any time, then tap **Settings > Refresh now** in the app.
- Anyone with the web app URL **and** the PIN can read and add data. Use a PIN that is hard to guess.
- Tenant names must be different from each other. Do not rename a tenant directly in the sheet; use the app's Edit button so history stays attached.
