// --- CONFIGURATION ---
const SHEET_NAME = "Sheet1", TRIGGER_STATUS = "Generate", OUTPUT_STATUS_DRAFT = "Draft", OUTPUT_STATUS_PUBLISHED = "Published";
const GEMINI_API_KEY_PROPERTY_NAME = 'GEMINI_API_KEY', PEXELS_API_KEY_PROPERTY_NAME_V1 = 'PEXELS_API_KEY_V1';
const BLOG_ID = 'BLOGGER_ID'; // --- REPLACE WITH YOUR BLOGGER ID ---
const [COL_A_TITLE_KEYWORD, COL_B_HTML_CONTENT, COL_C_LABELS, COL_D_STATUS, COL_E_POSTED_URL] = [0, 1, 2, 3, 4];
const GEMINI_API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"; // Or gemini-1.5-flash-latest
const MAX_ROWS_PER_GENERATION_RUN = 3;

// --- MENU ---
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Blog Automation')
    .addItem('Set Gemini API Key', 'setGeminiApiKeyMenu')
    .addItem('Set Pexels API Key', 'setPexelsApiKeyMenu').addSeparator()
    .addItem('1. Generate Blog Drafts (Batch)', 'processSheetForBlogGeneration')
    .addItem('2. Publish Drafts to Blogger', 'publishDraftsToBlogger')
    .addItem('4. Sync Posts & Comments', 'syncBloggerPostsAndComments')
    .addToUi();
}

// --- API KEY HANDLERS ---
function setGeminiApiKeyMenu() { setApiKeyDialog('Set Gemini API Key', 'Enter your Gemini API Key:', GEMINI_API_KEY_PROPERTY_NAME); }
function setPexelsApiKeyMenu() { setApiKeyDialog('Set Pexels API Key', 'Enter your Pexels API Key:', PEXELS_API_KEY_PROPERTY_NAME_V1); }
function setApiKeyDialog(title, promptMsg, propertyName) {
  const ui = SpreadsheetApp.getUi(), result = ui.prompt(title, promptMsg, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() == ui.Button.OK) {
    const apiKey = result.getResponseText().trim();
    if (apiKey) PropertiesService.getUserProperties().setProperty(propertyName, apiKey), ui.alert('API Key saved.');
    else ui.alert('API Key cannot be empty.');
  }
}
function getApiKeyFromProperties(propertyName, keyDisplayName) {
  const apiKey = PropertiesService.getUserProperties().getProperty(propertyName);
  if (!apiKey) return SpreadsheetApp.getUi().alert(keyDisplayName + " not set. Configure via menu."), null;
  return apiKey;
}
const getGeminiApiKey = () => getApiKeyFromProperties(GEMINI_API_KEY_PROPERTY_NAME, 'Gemini API Key');
const getPexelsApiKey = () => getApiKeyFromProperties(PEXELS_API_KEY_PROPERTY_NAME_V1, 'Pexels API Key');

// --- API CALLS ---
function callGeminiApi(apiKey, prompt) {
  if (!apiKey) return Logger.log("Gemini API key missing."), null;
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
  const endpoint = `${GEMINI_API_ENDPOINT}?key=${encodeURIComponent(apiKey)}`;
  try {
    const response = UrlFetchApp.fetch(endpoint, options), code = response.getResponseCode(), body = response.getContentText();
    if (code === 200) {
      const data = JSON.parse(body);
      if (data.candidates && data.candidates[0]?.content?.parts?.[0]?.text) return data.candidates[0].content.parts[0].text.trim();
      Logger.log(`Gemini: No content. FinishReason: ${data.candidates?.[0]?.finishReason}. Safety: ${JSON.stringify(data.candidates?.[0]?.safetyRatings)}`);
      return null;
    }
    Logger.log(`Gemini API Error ${code}: ${body}`);
    SpreadsheetApp.getUi().alert(`Gemini API Error: ${code}. Check Logs.`);
    return null;
  } catch (e) { Logger.log(`Gemini call exception: ${e}`); return null; }
}

function fetchImageFromPexels(apiKey, query) {
  if (!apiKey || !query?.trim()) return Logger.log(!apiKey ? "Pexels API key missing." : "Pexels query empty."), null;
  const mainQuery = query.split(':')[0].split(' - ')[0].trim();
  const url = `https://api.pexels.com/v1/search?query=${encodeURIComponent(mainQuery)}&per_page=1&orientation=landscape`;
  const options = { method: 'get', headers: { Authorization: apiKey }, muteHttpExceptions: true };
  try {
    const response = UrlFetchApp.fetch(url, options), code = response.getResponseCode(), body = response.getContentText();
    if (code === 200) {
      const photo = JSON.parse(body)?.photos?.[0];
      if (photo) return { imageUrl: photo.src.large, photographer: photo.photographer, photographerUrl: photo.photographer_url, pexelsUrl: photo.url };
      Logger.log(`Pexels: No photo for '${mainQuery}'.`); return null;
    }
    Logger.log(`Pexels API Error ${code}: ${body}`); return null;
  } catch (e) { Logger.log(`Pexels fetch exception: ${e}`); return null; }
}

// --- HELPERS ---
const extractH1Content = (html) => html?.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i)?.[1]?.trim() ?? null;

// --- CONTENT GENERATION ---
function generateOptimizedContentWithImage(geminiApiKey, pexelsApiKey, topic) {
  Logger.log(`Generating content for: ${topic}`);
  const blogPrompt = `You are an HTML blog post generator. Create an optimized HTML content for an article about "${topic}".
Instructions:
1.  Response MUST be ONLY raw HTML for article body (starts with <h1>). NO <!DOCTYPE>, <html>, <head>, <body> tags.
2.  Use one <h1> for article title (based on "${topic}"), <h2> for subheadings, <p> for paragraphs, <ul>/<ol> for lists.
3.  Content: ~500-700 words.
4.  NO extra text, explanations, or markdown fences.
5.  Integrate keywords naturally. Tone: informative, engaging.
Generate HTML for "${topic}":`;
  let htmlBodyContent = callGeminiApi(geminiApiKey, blogPrompt);
  if (!htmlBodyContent) return Logger.log(`Failed to generate body for ${topic}.`), null;
  htmlBodyContent = htmlBodyContent.replace(/^```html\s*|\s*```$/g, '').trim();

  const blogTitle = extractH1Content(htmlBodyContent) || topic;
  let fullHtml = `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${blogTitle.replace(/<[^>]+>/g, '')}</title><meta name="description" content="Blog post about ${blogTitle.replace(/<[^>]+>/g, '')}.">
<style>body{font-family:'Helvetica Neue',Arial,sans-serif;line-height:1.7;margin:0;padding:0;color:#333;background-color:#f9f9f9}article{max-width:750px;margin:30px auto;padding:25px 30px;background-color:#fff;border-radius:10px;box-shadow:0 4px 15px rgba(0,0,0,.08)}h1{color:#222;text-align:left;margin-bottom:25px;font-size:2.2em;line-height:1.2}h2{color:#333;margin-top:35px;margin-bottom:15px;border-bottom:2px solid #e0e0e0;padding-bottom:8px;font-size:1.6em}p{margin-bottom:18px;font-size:1.05em}ul,ol{margin-bottom:18px;padding-left:25px}li{margin-bottom:8px}figure{margin:30px 0;text-align:center}figcaption{font-size:.95em;color:#555;margin-top:10px;font-style:italic}img{max-width:100%;height:auto;border-radius:6px;box-shadow:0 5px 12px rgba(0,0,0,.12);display:block;margin:auto}a{color:#0066cc;text-decoration:none}a:hover{text-decoration:underline;color:#004c99}</style>
</head><body><article>${htmlBodyContent}</article></body></html>`;

  const imageData = fetchImageFromPexels(pexelsApiKey, blogTitle);
  if (imageData?.imageUrl) {
    const altTextPrompt = `Generate concise, descriptive alt text (under 125 chars) for an image about "${blogTitle}".
Response MUST be ONLY the alt text phrase. Example for "Backyard Birdwatching": "Colorful blue jay on bird feeder in sunny backyard."
Image about: "${blogTitle}". Alt text:`;
    let altText = callGeminiApi(geminiApiKey, altTextPrompt);
    altText = altText?.trim()?.replace(/^["']|["']$/g, '')?.replace(/\.$/, '')?.replace(/"/g, '"')?.replace(/'/g, '') || blogTitle;
    const imageHtml = `\n<figure><img src='${imageData.imageUrl}' alt='${altText}'><figcaption>Photo by <a href='${imageData.photographerUrl}' target='_blank' rel='noopener noreferrer'>${imageData.photographer}</a> on <a href='${imageData.pexelsUrl}' target='_blank' rel='noopener noreferrer'>Pexels</a></figcaption></figure>\n`;
    
    const h1Match = htmlBodyContent.match(/<\/h1>/i);
    if (h1Match) {
      const insertIdx = h1Match.index + h1Match[0].length;
      const updatedBody = htmlBodyContent.slice(0, insertIdx) + imageHtml + htmlBodyContent.slice(insertIdx);
      fullHtml = fullHtml.replace(htmlBodyContent, updatedBody);
    } else { // Fallback: insert image at the beginning of article if H1 is missing in body (should not happen)
      const articleMatch = fullHtml.match(/<article[^>]*>/i);
      if (articleMatch && typeof articleMatch.index === 'number') {
         const insertIdx = articleMatch.index + articleMatch[0].length;
         fullHtml = fullHtml.slice(0, insertIdx) + imageHtml + fullHtml.slice(insertIdx);
      } else { // Last resort
         const updatedBody = imageHtml + htmlBodyContent;
         fullHtml = fullHtml.replace(htmlBodyContent, updatedBody);
      }
    }
  } else Logger.log(`No image for ${blogTitle}.`);
  return fullHtml;
}

function generateOptimizedLabels(apiKey, topicOrTitle) {
  if (!topicOrTitle?.trim()) return Logger.log("No topic for labels."), "";
  const labelPrompt = `Generate 3-5 relevant, comma-separated tags for blog post: "${topicOrTitle}".
Response MUST be ONLY comma-separated tags. Example: "tag1, tag2, final tag"
Tags for: "${topicOrTitle}"`;
  let labels = callGeminiApi(apiKey, labelPrompt);
  return labels?.replace(/^labels:\s*|\.$/gi, '')?.trim() || "";
}

// --- SHEET PROCESSING ---
function processSheetForBlogGeneration() {
  const ui = SpreadsheetApp.getUi(), geminiKey = getGeminiApiKey(), pexelsKey = getPexelsApiKey();
  if (!geminiKey) return ui.alert("Gemini API Key not set.");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return ui.alert(`Sheet '${SHEET_NAME}' not found.`);
  
  const values = sheet.getDataRange().getValues();
  let generatedCount = 0, processedRows = 0;
  SpreadsheetApp.getActiveSpreadsheet().toast("Starting blog generation...", "Processing", 5);

  for (let i = 1; i < values.length && processedRows < MAX_ROWS_PER_GENERATION_RUN; i++) {
    const row = values[i], status = String(row[COL_D_STATUS]).trim().toLowerCase();
    const keyword = String(row[COL_A_TITLE_KEYWORD]).trim();
    if (status === TRIGGER_STATUS.toLowerCase() && keyword) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Generating for: ${keyword.substring(0,30)}...`, `Row ${i+1}`, 10);
      const fullHtml = generateOptimizedContentWithImage(geminiKey, pexelsKey, keyword);
      if (fullHtml) {
        const title = extractH1Content(fullHtml) || keyword;
        const labels = generateOptimizedLabels(geminiKey, title);
        sheet.getRange(i + 1, COL_A_TITLE_KEYWORD + 1).setValue(title);
        sheet.getRange(i + 1, COL_B_HTML_CONTENT + 1).setValue(fullHtml);
        sheet.getRange(i + 1, COL_C_LABELS + 1).setValue(labels);
        sheet.getRange(i + 1, COL_D_STATUS + 1).setValue(OUTPUT_STATUS_DRAFT);
        generatedCount++;
      } else sheet.getRange(i + 1, COL_D_STATUS + 1).setValue("Error: Content Gen Failed");
      processedRows++;
    }
  }
  if (generatedCount > 0) ui.alert(`${generatedCount} blog draft(s) generated.`);
  else if (processedRows === 0) ui.alert(`No rows to process with status '${TRIGGER_STATUS}'.`);
  else ui.alert(`Processed ${processedRows} row(s), 0 drafts generated. Check logs.`);
  Logger.log(`Blog generation finished. Drafts: ${generatedCount}`);
}

function publishDraftsToBlogger() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return ui.alert(`Sheet '${SHEET_NAME}' not found.`);

  const values = sheet.getDataRange().getValues(), token = ScriptApp.getOAuthToken();
  let publishedCount = 0;
  SpreadsheetApp.getActiveSpreadsheet().toast("Starting to publish drafts...", "Blogger Publishing", 5);

  for (let i = 1; i < values.length; i++) {
    const row = values[i], status = String(row[COL_D_STATUS]).trim().toLowerCase();
    if (status === OUTPUT_STATUS_DRAFT.toLowerCase()) {
      const title = String(row[COL_A_TITLE_KEYWORD]).trim() || extractH1Content(String(row[COL_B_HTML_CONTENT]).trim());
      const htmlContent = String(row[COL_B_HTML_CONTENT]).trim();
      const labels = String(row[COL_C_LABELS]).trim().split(',').map(t => t.trim()).filter(t => t);
      if (!title || !htmlContent) {
        sheet.getRange(i + 1, COL_D_STATUS + 1).setValue(`Error: ${!title ? 'Title' : 'Content'} Missing`); continue;
      }
      SpreadsheetApp.getActiveSpreadsheet().toast(`Publishing: ${title.substring(0,30)}...`, `Row ${i+1}`, 10);
      const payload = { kind: "blogger#post", blog: { id: BLOG_ID }, title, content: htmlContent, labels };
      const options = { method: 'post', contentType: 'application/json', headers: { Authorization: `Bearer ${token}` }, payload: JSON.stringify(payload), muteHttpExceptions: true };
      try {
        const response = UrlFetchApp.fetch(`https://www.googleapis.com/blogger/v3/blogs/${BLOG_ID}/posts/`, options);
        const code = response.getResponseCode(), resultText = response.getContentText();
        if (code === 200) {
          const result = JSON.parse(resultText);
          if (result.url) {
            sheet.getRange(i + 1, COL_D_STATUS + 1).setValue(OUTPUT_STATUS_PUBLISHED);
            sheet.getRange(i + 1, COL_E_POSTED_URL + 1).setValue(result.url);
            publishedCount++;
          } else sheet.getRange(i + 1, COL_D_STATUS + 1).setValue("Error: Publish No URL");
        } else sheet.getRange(i + 1, COL_D_STATUS + 1).setValue(`Error: Publish Failed (${code})`);
      } catch (e) { Logger.log(`Publish exception: ${e}`); sheet.getRange(i + 1, COL_D_STATUS + 1).setValue("Error: Publish Exception"); }
    }
  }
  if (publishedCount > 0) ui.alert(`${publishedCount} post(s) published.`);
  else ui.alert("No drafts to publish or all failed. Check logs.");
  Logger.log(`Blogger publishing finished. Published: ${publishedCount}`);
}

function syncBloggerPostsAndComments() {
  const blogId = BLOG_ID;
  const token = ScriptApp.getOAuthToken();

  const postsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Posts") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Posts");
  const commentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Comments") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Comments");

  // --- Clear existing content ---
  postsSheet.clear().appendRow(["Post ID", "Title", "Labels", "Published", "Updated", "Status", "URL", "Content"]);
  commentsSheet.clear().appendRow(["Comment ID", "Post ID", "Author Name", "Author Email", "Published", "Content"]);

  const postUrl = `https://www.googleapis.com/blogger/v3/blogs/${blogId}/posts?maxResults=100`;
  const options = {
    method: "get",
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  };

  try {
    const postResp = UrlFetchApp.fetch(postUrl, options);
    const postData = JSON.parse(postResp.getContentText());
    const posts = postData.items || [];

    for (const post of posts) {
      // Add post metadata to Posts sheet
      postsSheet.appendRow([
        post.id,
        post.title,
        (post.labels || []).join(", "),
        post.published,
        post.updated,
        post.status,
        post.url,
        post.content
      ]);

      // --- Fetch Comments for this Post ---
      const commentsUrl = `https://www.googleapis.com/blogger/v3/blogs/${blogId}/posts/${post.id}/comments`;
      const commentResp = UrlFetchApp.fetch(commentsUrl, options);
      const commentData = JSON.parse(commentResp.getContentText());
      const comments = commentData.items || [];

      for (const comment of comments) {
        commentsSheet.appendRow([
          comment.id,
          post.id,
          comment.author?.displayName || "",
          comment.author?.email || "",
          comment.published,
          comment.content
        ]);
      }
    }

    SpreadsheetApp.getUi().alert(`✅ Synced ${posts.length} posts and associated comments.`);
  } catch (e) {
    Logger.log("❌ Error syncing Blogger data: " + e);
    SpreadsheetApp.getUi().alert("Failed to sync Blogger data. Check logs.");
  }
}
