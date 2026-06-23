/**
 * Instagram Feed Scraper for @lelap.salatiga
 * Uses puppeteer-core + existing Chrome to extract posts from Instagram internal API.
 * Run: node scripts/ig_scraper.js
 * Output: data/ig_feed.json
 */
const puppeteer = require('puppeteer-core');
const https = require('https');
const fs = require('fs');
const path = require('path');

const PROFILE = 'lelap.salatiga';
const DATA_DIR = path.join(__dirname, '..', 'data');
const OUTPUT = path.join(DATA_DIR, 'ig_feed.json');
const IMG_DIR = path.join(DATA_DIR, 'ig_images');
const CHROME_PATH = 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const APP_ID = '936619743392459';
const MAX_POSTS = 20;

// Download image from URL to local file
function downloadImage(url, filepath) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(filepath);
    https.get(url, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400) {
        // Follow redirect
        https.get(res.headers.location, (r2) => {
          r2.pipe(file);
          file.on('finish', () => { file.close(); resolve(); });
        }).on('error', reject);
      } else {
        res.pipe(file);
        file.on('finish', () => { file.close(); resolve(); });
      }
    }).on('error', reject);
  });
}

async function scrape() {
  console.log(`[IG Scraper] Starting scrape for @${PROFILE}...`);
  
  const browser = await puppeteer.launch({
    executablePath: CHROME_PATH,
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage',
           '--disable-gpu', '--window-size=1280,720']
  });

  try {
    const page = await browser.newPage();
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36');
    
    // Navigate to profile (establishes cookies)
    await page.goto(`https://www.instagram.com/${PROFILE}/`, { waitUntil: 'networkidle2', timeout: 30000 });
    console.log('[IG Scraper] Page loaded');

    // Use internal API to get posts
    const posts = await page.evaluate(async (appId, maxPosts) => {
      const all = [];
      let cursor = null;
      
      for (let page = 0; page < 6; page++) {
        let url = `/api/v1/users/web_profile_info/?username=lelap.salatiga`;
        if (cursor) url += '&after=' + cursor;
        
        const resp = await fetch(url, { headers: { 'x-ig-app-id': appId } });
        const data = await resp.json();
        const media = data.data?.user?.edge_owner_to_timeline_media;
        const edges = media?.edges || [];
        
        for (const e of edges) {
          const node = e.node;
          all.push({
            shortcode: node.shortcode,
            thumbnail: node.thumbnail_src,
            display_url: node.display_url,
            caption: node.edge_media_to_caption?.edges?.[0]?.node?.text || '',
            taken_at: node.taken_at_timestamp,
            is_video: node.is_video || false,
            likes: node.edge_liked_by?.count || 0
          });
        }
        
        cursor = media?.page_info?.end_cursor;
        if (!media?.page_info?.has_next_page || all.length >= maxPosts * 2) break;
      }
      
      // Deduplicate by shortcode
      const seen = new Set();
      return all.filter(p => {
        if (seen.has(p.shortcode)) return false;
        seen.add(p.shortcode);
        return true;
      }).slice(0, maxPosts);
    }, APP_ID, MAX_POSTS);

    const igPosts = posts;
    
    // Ensure image directory exists
    if (!fs.existsSync(IMG_DIR)) fs.mkdirSync(IMG_DIR, { recursive: true });
    
    // Download images locally (Instagram CDN URLs expire quickly)
    console.log(`[IG Scraper] Downloading ${igPosts.length} images...`);
    for (let i = 0; i < igPosts.length; i++) {
      const p = igPosts[i];
      const localFile = path.join(IMG_DIR, `${p.shortcode}.jpg`);
      try {
        const imgUrl = p.thumbnail || p.display_url;
        await downloadImage(imgUrl, localFile);
        p.local_image = `data/ig_images/${p.shortcode}.jpg`;
        console.log(`[IG Scraper]   [${i+1}/${igPosts.length}] ${p.shortcode} ✓`);
      } catch (e) {
        console.log(`[IG Scraper]   [${i+1}/${igPosts.length}] ${p.shortcode} ✗ (${e.message})`);
        // Keep CDN URL as fallback
      }
    }

    // Build feed object
    const feed = {
      profile: PROFILE,
      fetched_at: new Date().toISOString(),
      post_count: igPosts.length,
      posts: igPosts.map(p => ({
        shortcode: p.shortcode,
        thumbnail: p.local_image || p.thumbnail,
        display_url: p.display_url,
        caption: p.caption,
        taken_at: p.taken_at,
        date: new Date(p.taken_at * 1000).toISOString().split('T')[0],
        is_video: p.is_video,
        url: `https://www.instagram.com/p/${p.shortcode}/`,
        likes: p.likes
      }))
    };

    fs.writeFileSync(OUTPUT, JSON.stringify(feed, null, 2));
    console.log(`[IG Scraper] Done! ${feed.post_count} posts saved to ${OUTPUT}`);
    return feed;

  } finally {
    await browser.close();
  }
}

scrape()
  .then(f => {
    console.log(`[IG Scraper] Success — ${f.post_count} posts`);
    process.exit(0);
  })
  .catch(err => {
    console.error(`[IG Scraper] Failed:`, err.message);
    process.exit(1);
  });
