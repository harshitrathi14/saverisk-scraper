// login.js — opens a real browser so you can log in to Saverisk once (OTP/SSO/reCAPTCHA).
// It AUTO-DETECTS a successful login (when you leave the flogin page) and saves the
// session to ./.session and storageState.json. No need to press Enter.
//
// Run:  npm run login   (or: node login.js)

const { chromium } = require('playwright');
const path = require('path');

const USER_DATA_DIR = path.join(__dirname, '.session');
const LOGIN_URL = 'https://www.saverisk.com/flogin.aspx';
const MAX_WAIT_MS = 20 * 60 * 1000; // 20 minutes to complete login
const POLL_MS = 1500;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function isLoggedInUrl(u) {
  if (!u) return false;
  if (u === 'about:blank') return false;
  // Logged in == navigated away from the login page, still on saverisk
  return /saverisk\.com/i.test(u) && !/flogin\.aspx/i.test(u) && !/login/i.test(u.split('/').pop());
}

(async () => {
  console.log('Launching browser… a window will appear on your desktop.');
  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless: false,
    viewport: { width: 1366, height: 850 },
    args: ['--start-maximized'],
  });

  const page = context.pages()[0] || (await context.newPage());
  await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded' }).catch(() => {});

  console.log('\n=================================================');
  console.log(' Please LOG IN in the browser window (number + OTP, or SSO).');
  console.log(' This script will detect success automatically and save your session.');
  console.log('=================================================\n');

  const start = Date.now();
  let lastUrl = '';
  let saved = false;

  while (Date.now() - start < MAX_WAIT_MS) {
    // Look across all open tabs for a logged-in URL
    const urls = context.pages().map((p) => p.url());
    const loggedIn = urls.find(isLoggedInUrl);
    const cur = urls.join(' | ');
    if (cur !== lastUrl) { console.log('URL:', cur); lastUrl = cur; }

    if (loggedIn) {
      console.log('\nDetected login ->', loggedIn);
      await sleep(4000); // let the dashboard settle / cookies finalize
      await context.storageState({ path: path.join(__dirname, 'storageState.json') });
      console.log('Session saved to ./.session and storageState.json');
      saved = true;
      break;
    }
    await sleep(POLL_MS);
  }

  if (!saved) console.log('\nTimed out waiting for login. Re-run `npm run login` and try again.');
  console.log('Closing browser. You can re-run the scraper now.');
  await context.close();
  process.exit(saved ? 0 : 1);
})().catch((err) => {
  console.error('Login script error:', err);
  process.exit(1);
});
