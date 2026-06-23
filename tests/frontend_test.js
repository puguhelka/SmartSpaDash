// Frontend smoke test for Lelap Mom Baby Care
// Verify Flutter web app loads correctly

const URL = 'https://significantly-endless-lite-toxic.trycloudflare.com';

async function test() {
  // 1. Navigate to the app
  await page.goto(URL, { waitUntil: 'networkidle2', timeout: 30000 });
  
  // 2. Verify page loaded
  const title = await page.title();
  console.log('Page title:', title);
  
  // 3. Wait for Flutter app to render
  await page.waitForSelector('flutter-view', { timeout: 15000 });
  console.log('Flutter view detected');
  
  // 4. Check for content
  const text = await page.evaluate(() => document.body.innerText.substring(0, 500));
  console.log('Page text sample:', text);
  
  return { success: true, message: 'Frontend loads correctly' };
}

module.exports = { test };
