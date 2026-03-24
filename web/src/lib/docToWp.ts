import * as mammoth from 'mammoth';

export function fixMojibake(text: string): string {
  return text
    .replace(/â€œ/g, '\u201C')
    .replace(/â€\u009d/g, '\u201D')
    .replace(/â€™/g, '\u2019')
    .replace(/â€˜/g, '\u2018')
    .replace(/â€¦/g, '\u2026')
    .replace(/â€/g, '\u201D')
    .replace(/Â /g, ' ');
}

export function htmlToWpBlocks(html: string): string {
  const fixed = fixMojibake(html);
  const parser = new DOMParser();
  const doc = parser.parseFromString(fixed, 'text/html');
  const nonEmpty = Array.from(doc.body.querySelectorAll('p')).filter(
    p => (p.textContent ?? '').trim() !== ''
  );
  if (nonEmpty.length === 0) return '(No content found)';

  let idx = 0;
  let conferenceLabel = '';
  let articleTitle = '';
  let authorName = '';

  if (nonEmpty[idx]?.querySelector('strong') && /conference/i.test(nonEmpty[idx].textContent ?? '')) {
    conferenceLabel = (nonEmpty[idx].textContent ?? '').trim();
    idx++;
  }
  if (idx < nonEmpty.length && nonEmpty[idx]?.querySelector('strong') && !/^by\s/i.test((nonEmpty[idx].textContent ?? '').trim())) {
    articleTitle = (nonEmpty[idx].textContent ?? '').trim().replace(/^[\u201C"]|[\u201D"]$/g, '');
    idx++;
  }
  if (idx < nonEmpty.length && /^by\s+/i.test((nonEmpty[idx].textContent ?? '').trim())) {
    authorName = (nonEmpty[idx].textContent ?? '').trim().replace(/^by\s+/i, '');
    idx++;
  }

  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const parts: string[] = [];

  if (conferenceLabel) {
    parts.push(
      `<!-- wp:paragraph {"style":{"typography":{"fontSize":"20px"}}} -->\n` +
      `<p style="font-size:20px"><mark style="background-color:rgba(0, 0, 0, 0);color:#cf2e2e" class="has-inline-color"><strong>${esc(conferenceLabel)}\u00a0</strong></mark></p>\n` +
      `<!-- /wp:paragraph -->`
    );
  }

  if (articleTitle) {
    const titleJson = JSON.stringify(articleTitle).slice(1, -1);
    parts.push(
      `<!-- wp:ultimate-post/heading {"blockId":"c901aa","currentPostId":"","headingText":"${titleJson}","headingStyle":"style1","headingAlign":"center","headingTypo":{"openTypography":1,"size":{"lg":42,"unit":"px","xs":"30"},"height":{"lg":"","unit":"px"},"decoration":"none","transform":"","family":"","weight":"400"},"headingColor":"var(\\u002d\\u002dpostx_preset_Contrast_3_color)","headingSpacing":{"lg":"10","sm":10,"unit":"px","ulg":"px","uxs":"px","xs":"10"},"wrapBg":{"openColor":1,"type":"color","color":"rgba(234,240,240,1)","gradient":"linear-gradient(90deg, rgb(6, 147, 227) 0%, rgb(155, 81, 224) 100%)","clip":false}} /-->`
    );
  }

  parts.push(
    `<!-- wp:image {"sizeSlug":"full","linkDestination":"none","align":"wide"} -->\n` +
    `<figure class="wp-block-image alignwide size-full"><img src="TODO_featured_image_url" alt="" /></figure>\n` +
    `<!-- /wp:image -->`
  );

  const today = new Date().toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
  parts.push(
    `<!-- wp:preformatted {"style":{"typography":{"fontSize":"12px"}}} -->\n` +
    `<pre class="wp-block-preformatted" style="font-size:12px">Author\uff1a${esc(authorName) || 'TODO'}<br>Date: ${today}<br>Category\uff1a<a href="https://enyu.acccn.org/mission-en/" target="_blank" rel="noreferrer noopener">Missions Beyond Borders</a></pre>\n` +
    `<!-- /wp:preformatted -->`
  );

  for (const p of nonEmpty.slice(idx)) {
    const inner = (p.innerHTML ?? '').replace(/^[\t ]+/, '').trim();
    if (!inner) continue;
    parts.push(
      `<!-- wp:paragraph {"editorskit":{"tabletAlignment":"justify","mobileAlignment":"justify","devices":false,"desktop":true,"tablet":true,"mobile":true,"loggedin":true,"loggedout":true,"acf_visibility":"","acf_field":"","acf_condition":"","acf_value":"","migrated":false,"unit_test":false}} -->\n` +
      `<p class="has-tablet-text-align-justify has-mobile-text-align-justify">${inner}</p>\n` +
      `<!-- /wp:paragraph -->`
    );
  }

  return parts.join('\n\n');
}

export type DocxImage = {
  name: string;
  mimeType: string;
  b64: string;
};

export type DocxConvertResult = {
  wpBlocks: string;
  previewHtml: string;
  images: DocxImage[];
};

export async function convertDocxToHtml(arrayBuffer: ArrayBuffer): Promise<DocxConvertResult> {
  const images: DocxImage[] = [];
  let imgIdx = 0;
  const result = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      convertImage: mammoth.images.imgElement(async (img) => {
        const b64 = await img.read('base64');
        const mime = img.contentType || 'image/jpeg';
        const ext = mime.split('/')[1] ?? 'jpg';
        images.push({ name: `image-${++imgIdx}.${ext}`, mimeType: mime, b64 });
        return { src: `data:${mime};base64,${b64}`, 'data-wp-img': String(imgIdx - 1) };
      }),
    }
  );
  return {
    wpBlocks: htmlToWpBlocks(result.value),
    previewHtml: fixMojibake(result.value),
    images,
  };
}

export function buildWpOutputPage(result: DocxConvertResult, wpToken?: string): string {
  const { wpBlocks, previewHtml, images } = result;
  const escaped = wpBlocks.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const imagesJson = JSON.stringify(images);
  const tokenJson = JSON.stringify(wpToken || '');
  const uploadBtn = images.length > 0
    ? `<button id="btn-upload" onclick="uploadImages()">Upload Images (${images.length})</button>`
    : '';

  const scriptBody = [
    'var IMAGES = ' + imagesJson + ';',
    'var WP_TOKEN = ' + tokenJson + ';',
    'function copyCode() {',
    '  navigator.clipboard.writeText(document.getElementById("b").textContent)',
    '    .then(function() {',
    '      var btn = document.getElementById("btn-copy");',
    '      btn.textContent = "Copied \u2713";',
    '      setTimeout(function() { btn.textContent = "Copy All"; }, 2000);',
    '    }).catch(function() { document.getElementById("btn-copy").textContent = "Error"; });',
    '}',
    'function show(id) {',
    '  document.querySelectorAll(".pane").forEach(function(p) { p.classList.remove("active"); });',
    '  document.getElementById(id).classList.add("active");',
    '}',
    'async function uploadImages() {',
    '  var statusEl = document.getElementById("upload-status");',
    '  if (!WP_TOKEN) {',
    '    statusEl.innerHTML = "<span style=\\"color:red\\">No WP token configured. Set enyu_wp_token in FreedCamp Secrets.</span>";',
    '    return;',
    '  }',
    '  var authHeader = "Basic " + btoa(WP_TOKEN);',
    '  statusEl.innerHTML = "";',
    '  var btn = document.getElementById("btn-upload");',
    '  if (btn) btn.disabled = true;',
    '  for (var i = 0; i < IMAGES.length; i++) {',
    '    var img = IMAGES[i];',
    '    var div = document.createElement("div");',
    '    div.textContent = "Uploading " + img.name + "...";',
    '    statusEl.appendChild(div);',
    '    try {',
    '      var binaryStr = atob(img.b64);',
    '      var bytes = new Uint8Array(binaryStr.length);',
    '      for (var j = 0; j < binaryStr.length; j++) bytes[j] = binaryStr.charCodeAt(j);',
    '      var blob = new Blob([bytes], { type: img.mimeType });',
    '      var res = await fetch("https://enyu.acccn.org/wp-json/wp/v2/media", {',
    '        method: "POST",',
    '        headers: {',
    '          "Authorization": authHeader,',
    '          "Content-Disposition": "attachment; filename=\\"" + img.name + "\\"",',
    '          "Content-Type": img.mimeType',
    '        },',
    '        body: blob',
    '      });',
    '      var data = await res.json();',
    '      if (data.source_url) {',
    '        document.querySelectorAll("[data-wp-img=\\"" + i + "\\"]").forEach(function(el) { el.setAttribute("src", data.source_url); });',
    '        div.innerHTML = "\u2713 <a href=\\"" + data.source_url + "\\" target=\\"_blank\\">" + img.name + "</a>";',
    '        div.style.color = "green";',
    '      } else {',
    '        div.textContent = "\u2717 " + img.name + ": " + (data.message || JSON.stringify(data));',
    '        div.style.color = "red";',
    '      }',
    '    } catch(e) {',
    '      div.textContent = "\u2717 " + img.name + ": " + String(e);',
    '      div.style.color = "red";',
    '    }',
    '  }',
    '  if (btn) btn.disabled = false;',
    '}',
  ].join('\n');

  return [
    '<!DOCTYPE html><html><head><meta charset="utf-8"><title>WP Blocks</title><style>',
    '*,*::before,*::after{box-sizing:border-box}',
    'body{font-family:sans-serif;margin:0;background:#f0f0f0}',
    '.toolbar{position:sticky;top:0;z-index:10;background:#fff;border-bottom:1px solid #ddd;padding:.6rem 1rem;display:flex;gap:.5rem;align-items:center;flex-wrap:wrap}',
    '.toolbar button{padding:.4rem 1.1rem;border:none;border-radius:4px;cursor:pointer;font-size:.9rem;color:#fff}',
    '#btn-copy{background:#6f42c1}#btn-preview{background:#0078d4}#btn-code{background:#495057}#btn-upload{background:#e91e63}',
    '#upload-status{padding:.5rem 1rem;font-size:.85rem}',
    '#upload-status a{color:#0078d4}',
    '.pane{display:none;padding:1rem}.pane.active{display:block}',
    '#preview-pane{max-width:780px;margin:0 auto;background:#fff;padding:2rem;border-radius:6px;margin-top:1rem;line-height:1.7;font-size:1rem}',
    '#preview-pane p{margin:.9em 0}#preview-pane strong{font-weight:700}#preview-pane img{max-width:100%;height:auto}',
    'pre{background:#fff;padding:1rem;border:1px solid #ddd;border-radius:4px;white-space:pre-wrap;word-break:break-word;font-size:.82rem}',
    '</style></head><body>',
    '<div class="toolbar">',
    '  <button id="btn-preview" onclick="show(\'preview\')">Preview</button>',
    '  <button id="btn-code" onclick="show(\'code\')">WP Code</button>',
    '  <button id="btn-copy" onclick="copyCode()">Copy All</button>',
    '  ' + uploadBtn,
    '</div>',
    '<div id="upload-status"></div>',
    '<div id="preview" class="pane active"><div id="preview-pane">' + previewHtml + '</div></div>',
    '<div id="code" class="pane"><pre id="b">' + escaped + '</pre></div>',
    '<script>',
    scriptBody,
    '<\/script>',
    '</body></html>',
  ].join('\n');
}
