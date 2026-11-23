import { createHash } from 'node:crypto';
import { access, copyFile, mkdir, stat, unlink } from 'node:fs/promises';
import { extname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { AuthRequiredError, CommandExecutionError, ArgumentError } from '@jackwener/opencli/errors';
import { formatCookieHeader } from '@jackwener/opencli/download';
import { downloadMedia } from '@jackwener/opencli/download/media-download';
import { cli, Strategy } from '@jackwener/opencli/registry';

const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const packageEntry = fileURLToPath(import.meta.resolve('@jackwener/opencli'));
const packageRoot = resolve(packageEntry, '..', '..', '..');
const { extractAssetsForInput } = await import(pathToFileURL(join(packageRoot, 'clis', '1688', 'assets.js')).href);

const READ_PRODUCT = `(() => {
  const text = document.body?.innerText || '';
  if (/请拖动下方滑块完成验证|请按住滑块|通过验证以确保正常访问|验证码拦截|访问验证|滑动验证|_____tmd_____|访问过于频繁/.test(text) || location.href.includes('/_____tmd_____/punish')) return { blocked: true, text: text.slice(0, 240) };
  if (/登录后|请先登录|账号登录/.test(text) && !document.querySelector('p[title]')) return { login: true };
  const lines = text.split(/\\r?\\n/).map((line) => line.replace(/\\u00a0/g, ' ').trim()).filter(Boolean);
  const section = (heading, stops) => {
    let start = -1;
    for (let index = 0; index < lines.length; index++) if (lines[index] === heading) start = index;
    if (start < 0) return [];
    let end = lines.length;
    for (let index = start + 1; index < lines.length; index++) if (stops.includes(lines[index])) { end = index; break; }
    return lines.slice(start + 1, end).flatMap((line) => line.split(/\\t+/).map((cell) => cell.trim())).filter(Boolean);
  };
  const parsePairs = (items, structural) => {
    const rows = [];
    for (let index = 0; index < items.length;) {
      const key = items[index++];
      if (!key || structural.has(key)) continue;
      const value = items[index];
      if (!value || structural.has(value)) continue;
      index++;
      if (!rows.some((row) => row.key === key && row.value === value)) rows.push({ key, value });
    }
    return rows;
  };
  const attrTokens = section('商品属性', ['包装信息', '商品详情', '商品资质证书']);
  const packTokens = section('包装信息', ['商品详情', '商品评价', '商品属性', '商品资质证书']);
  const certTokens = section('商品资质证书', ['包装信息', '商品详情', '商品评价', '商品属性']);
  const structural = new Set(['全部参数', '商品属性', '商品件重尺', '包装信息', '商品详情', '商品资质证书', '商品评价', '商品参数']);
  return {
    href: location.href,
    title: document.title || '',
    visible_attributes: parsePairs(attrTokens, structural),
    packaging_info: parsePairs(packTokens, structural),
    qualification_info: parsePairs(certTokens, structural),
    page_sections_found: { attributes: attrTokens.length > 0, packaging: packTokens.length > 0, qualification: certTokens.length > 0 },
  };
})()`;

function mediaExt(url, fallback) {
  try {
    const extension = extname(new URL(url).pathname).toLowerCase();
    return /^\.[a-z0-9]{1,7}$/.test(extension) ? extension : fallback;
  } catch { return fallback; }
}

async function fileExists(path) {
  try {
    const info = await stat(path);
    return info.isFile() && info.size > 0;
  } catch { return false; }
}

cli({
  site: '1688',
  name: 'xudaai-product',
  access: 'read',
  description: '读取商品属性和包装信息，并下载素材；相同视频URL复用本地缓存',
  domain: '1688.com',
  strategy: Strategy.COOKIE,
  browser: true,
  navigateBefore: false,
  args: [
    { name: 'input', type: 'string', positional: true, required: true, help: '1688 商品 URL 或 offer ID' },
    { name: 'output', type: 'string', required: true, help: '本商品素材输出目录' },
    { name: 'videoCache', type: 'string', required: true, help: '共享视频缓存目录' },
  ],
  columns: ['offer_id', 'title', 'attribute_count', 'packaging_count', 'image_count', 'video_count'],
  func: async (page, args) => {
    const input = String(args.input || '').trim();
    const match = input.match(/(?:offer\/|offerId=)?(\d{8,})/);
    if (!match) throw new ArgumentError('input must contain a valid 1688 offer ID');
    const offerId = match[1];
    const itemUrl = `https://detail.1688.com/offer/${offerId}.html`;
    const outputDir = resolve(String(args.output || ''));
    const videoCache = resolve(String(args.videoCache || ''));
    if (!String(args.output || '').trim() || !String(args.videoCache || '').trim()) throw new ArgumentError('output and videoCache are required');
    try { await page.goto(itemUrl, { waitUntil: 'domcontentloaded', timeout: 45000 }); }
    catch (error) { throw new CommandExecutionError(`Unable to open 1688 item: ${error.message}`); }
    await wait(2500);
    let product = null;
    for (let attempt = 0; attempt < 8; attempt++) {
      product = await page.evaluate(READ_PRODUCT);
      if (product.blocked) throw new CommandExecutionError('1688 verification/rate-limit detected; stopped before extracting product data.');
      if (product.login) throw new AuthRequiredError('1688.com', 'Log in to 1688 in the connected Chrome profile, then retry.');
      if (product.page_sections_found?.attributes || product.page_sections_found?.packaging) break;
      await wait(1000);
    }
    if (!product?.page_sections_found?.attributes && !product?.page_sections_found?.packaging) {
      throw new CommandExecutionError('商品属性和包装信息标签未加载，已停止，避免生成空档案。');
    }

    const assets = await extractAssetsForInput(page, offerId);
    const cookies = await page.getCookies({ domain: '1688.com' });
    const cookieHeader = formatCookieHeader(cookies);
    await mkdir(outputDir, { recursive: true });
    await mkdir(videoCache, { recursive: true });

    const groups = [
      ['main', assets.main_images],
      ['sku', assets.sku_images],
      ['detail', assets.detail_images],
      ['other', assets.other_images],
    ];
    const imageItems = [];
    const seenImages = new Set();
    for (const [group, urls] of groups) {
      for (let index = 0; index < (urls || []).length; index++) {
        const url = urls[index];
        if (seenImages.has(url)) continue;
        seenImages.add(url);
        imageItems.push({ type: 'image', url, filename: `${offerId}_${group}_${String(index + 1).padStart(2, '0')}${mediaExt(url, '.jpg')}` });
      }
    }
    const imageResults = imageItems.length
      ? await downloadMedia(imageItems, { output: outputDir, cookies: cookieHeader, browserCookies: cookies, timeout: 60000, filenamePrefix: offerId, verbose: false })
      : [];

    const videoResults = [];
    const seenVideos = new Set();
    for (let index = 0; index < (assets.videos || []).length; index++) {
      const url = assets.videos[index];
      if (seenVideos.has(url)) continue;
      seenVideos.add(url);
      const urlHash = createHash('sha256').update(url).digest('hex');
      const extension = mediaExt(url, '.mp4');
      const cacheFile = join(videoCache, `${urlHash}${extension}`);
      const targetFile = join(outputDir, `${offerId}_video_${String(index + 1).padStart(2, '0')}${extension}`);
      let reused = await fileExists(cacheFile);
      if (!reused) {
        const rows = await downloadMedia([{ type: 'video', url, filename: `${urlHash}${extension}` }], {
          output: videoCache, cookies: cookieHeader, browserCookies: cookies, timeout: 120000,
          filenamePrefix: urlHash, verbose: false,
        });
        if (rows.some((row) => row.status !== 'success') || !(await fileExists(cacheFile))) {
          await unlink(cacheFile).catch(() => {});
          videoResults.push({ url_hash: urlHash, status: 'failed', error: rows.map((row) => row.error || row.size).filter(Boolean).join('; ') || '下载未生成有效文件' });
          continue;
        }
      }
      await copyFile(cacheFile, targetFile);
      videoResults.push({ url_hash: urlHash, status: 'success', reused_cache: reused, filename: targetFile });
    }

    return [{
      offer_id: offerId,
      title: product.title.replace(/\s*-\s*阿里巴巴$/, '').trim(),
      item_url: product.href,
      visible_attributes: product.visible_attributes,
      packaging_info: product.packaging_info,
      qualification_info: product.qualification_info,
      page_sections_found: product.page_sections_found,
      assets: { main_count: assets.main_count, sku_count: assets.sku_count, detail_count: assets.detail_count,
        video_count: assets.video_count, image_results: imageResults, video_results: videoResults },
      attribute_count: product.visible_attributes.length,
      packaging_count: product.packaging_info.length,
      image_count: imageResults.filter((row) => row.status === 'success').length,
      video_count: videoResults.filter((row) => row.status === 'success').length,
    }];
  },
});
