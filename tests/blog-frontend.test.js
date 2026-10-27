const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const html = fs.readFileSync('blog/index.html', 'utf8');
const script = html.match(/<script id="blog-script">([\s\S]*?)<\/script>/)[1];
const legacyHtml = fs.readFileSync('blog/article.html', 'utf8');
const redirectScript = legacyHtml.match(/<script id="blog-redirect">([\s\S]*?)<\/script>/)[1];

function makeNode(tag) {
  return {
    tag, children: [], textContent: '', className: '', attributes: {}, hidden: true,
    append(...nodes) { this.children.push(...nodes); },
    replaceChildren() { this.children = []; },
    setAttribute(name, value) { this.attributes[name] = value; }
  };
}

function makePage(search, response, ok = true) {
  // Both views exist in the actual shared page; startup must select only one.
  const nodes = new Map([...html.matchAll(/\bid="([^"]+)"/g)]
    .map(([, id]) => [id, makeNode(id)]));
  const metadata = { content: '' };
  const document = {
    title: '',
    getElementById(id) { return nodes.get(id) || null; },
    createElement: makeNode,
    querySelector(selector) { return selector === 'meta[name="description"]' ? metadata : null; }
  };
  const saved = new Map();
  const requests = [];
  const context = vm.createContext({
    document, location: { search }, URLSearchParams, URL,
    sessionStorage: { getItem: key => saved.get(key) || null, setItem: (key, value) => saved.set(key, value) },
    fetch: async url => {
      requests.push(String(url));
      const data = typeof response === 'function' ? await response() : response;
      return { ok, json: async () => data };
    },
    console
  });
  const ready = vm.runInContext(script, context, { filename: 'blog/index.html#blog-script' });
  return { context, nodes, requests, document, metadata, ready, saved };
}

function redirect(search, hash = '') {
  const href = `https://example.test/blog/article.html${search}${hash}`;
  const calls = [];
  const context = vm.createContext({
    URL, location: { href, search, hash, replace: url => calls.push(url) }
  });
  vm.runInContext(redirectScript, context);
  assert.equal(calls.length, 1, 'legacy links redirect once with location.replace');
  return new URL(calls[0]);
}

(async () => {
  assert.deepEqual(fs.readdirSync('blog').sort(), ['article.html', 'index.html']);
  assert.doesNotMatch(html, /(?:src|href)=["'][^"']*blog\.(?:js|css)/);
  assert.match(html, /<main id="blogList"[^>]*hidden>/);
  assert.match(html, /<main id="blogArticle"[^>]*hidden>/);

  const posts = Array.from({ length: 3 }, (_, index) => ({
    slug: `story-${index + 1}`, category: 'farmers', title: `Story ${index + 1}`,
    excerpt: `Summary ${index + 1}`, coverImageUrl: 'https://example.test/photo.jpg'
  }));
  const listResponse = { status: 'success', page: 1, pages: 2, total: 4, posts };
  const page = makePage('?category=farmers&page=1', listResponse);
  await page.ready;
  assert.equal(page.nodes.get('blogList').hidden, false);
  assert.equal(page.nodes.get('blogArticle').hidden, true);
  assert.equal(page.requests.length, 1, 'startup fetches only the selected view');
  assert.match(page.requests[0], /action=getBlogPosts/);
  assert.equal(page.nodes.get('categoryNav').children.length, 3);
  assert.equal(page.nodes.get('categoryNav').children[1].attributes['aria-current'], 'page');
  assert.equal(page.nodes.get('postGrid').children.length, 3);
  assert.equal(page.nodes.get('postGrid').children[0].children[0].href, '/blog/?category=farmers&slug=story-1');
  assert.equal(page.nodes.get('pagination').children[0].textContent, 'Previous');
  assert.equal(page.nodes.get('pagination').children[0].tag, 'span');
  assert.equal(page.nodes.get('pagination').children.at(-1).href, '/blog/?category=farmers&page=2');
  assert.equal(page.document.title, 'The Farmers | Oriberry Journal');

  const second = makePage('?category=business&page=2', { ...listResponse, page: 2, posts: [posts[0]] });
  await second.ready;
  assert.equal(second.nodes.get('postGrid').children.length, 1);
  assert.equal(second.nodes.get('pagination').children[0].href, '/blog/?category=business&page=1');
  assert.equal(second.nodes.get('pagination').children.at(-1).tag, 'span');
  for (const category of ['business', 'farmers', 'processing']) {
    const categoryPage = makePage(`?category=${category}&page=1`, listResponse);
    await categoryPage.ready;
    assert.equal(new URL(categoryPage.requests[0]).searchParams.get('category'), category);
  }
  const fallback = makePage('?category=unknown&page=-2', listResponse);
  await fallback.ready;
  assert.equal(new URL(fallback.requests[0]).searchParams.get('category'), 'farmers');
  assert.equal(new URL(fallback.requests[0]).searchParams.get('page'), '1');

  const detailResponse = { status: 'success', post: {
    title: 'Processing Story', excerpt: 'A short introduction',
    coverImageUrl: 'https://example.test/photo.jpg',
    content: '## Heading\n- First point\nPlain paragraph'
  } };
  const detail = makePage('?category=processing&slug=processing-story', detailResponse);
  await detail.ready;
  assert.equal(detail.nodes.get('blogList').hidden, true);
  assert.equal(detail.nodes.get('blogArticle').hidden, false);
  assert.equal(detail.requests.length, 1);
  assert.match(detail.requests[0], /action=getBlogPost&/);
  assert.equal(detail.nodes.get('backToCategory').href, '/blog/?category=processing&page=1');
  assert.equal(detail.nodes.get('articleTitle').textContent, 'Processing Story');
  assert.equal(detail.document.title, 'Processing Story | Oriberry Journal');
  assert.equal(detail.nodes.get('articleCover').hidden, false);
  assert.deepEqual(detail.nodes.get('articleBody').children.map(node => node.tag), ['h2', 'ul', 'p']);
  assert.equal(detail.metadata.content, 'A short introduction');

  const missing = makePage('?category=processing&slug=missing', {
    status: 'error', code: 'NOT_FOUND', message: 'Không tìm thấy bài viết.'
  });
  await missing.ready;
  assert.equal(missing.nodes.get('articleBody').children[0].className, 'state');
  assert.match(missing.nodes.get('articleBody').children[0].textContent, /Không tìm thấy/);
  const emptySlug = makePage('?category=processing&slug=', detailResponse);
  await emptySlug.ready;
  assert.equal(emptySlug.nodes.get('blogArticle').hidden, false);
  assert.equal(emptySlug.requests.length, 0);
  assert.match(emptySlug.nodes.get('articleBody').children[0].textContent, /Thiếu đường dẫn/);
  assert.equal(emptySlug.document.title, 'Bài viết | Oriberry Journal');

  for (const [total, message] of [[0, /chưa có bài viết được xuất bản/], [4, /Trang này chưa có/]]) {
    const empty = makePage('?category=farmers&page=9', { ...listResponse, posts: [], total });
    await empty.ready;
    assert.match(empty.nodes.get('postGrid').children[0].textContent, message);
  }
  const offline = makePage('?category=farmers', () => { throw new Error('Network offline'); });
  await offline.ready;
  assert.match(offline.nodes.get('postGrid').children[0].textContent, /Network offline/);
  const httpError = makePage('?slug=story-1', detailResponse, false);
  await httpError.ready;
  assert.match(httpError.nodes.get('articleBody').children[0].textContent, /Không thể kết nối/);

  let resolveResponse;
  const loading = makePage('?category=farmers', () => new Promise(resolve => { resolveResponse = resolve; }));
  assert.equal(loading.nodes.get('blogList').hidden, false, 'view is visible while fetch is pending');
  assert.equal(loading.nodes.get('blogArticle').hidden, true);
  resolveResponse(listResponse);
  await loading.ready;

  const query = '?category=processing&slug=story%2D1&page=2&extra=a%26b';
  const target = redirect(query, '#articleBody');
  assert.equal(target.origin, 'https://example.test');
  assert.equal(target.pathname, '/blog/');
  assert.equal(target.search, query);
  assert.equal(target.hash, '#articleBody');
  for (const search of ['', '?category=processing', '?category=farmers&slug=']) {
    const missingTarget = redirect(search, '#articleBody');
    assert.equal(missingTarget.searchParams.has('slug'), true);
    assert.equal(missingTarget.searchParams.get('slug'), '');
    assert.equal(missingTarget.hash, '#articleBody');
    const errorPage = makePage(missingTarget.search, detailResponse);
    await errorPage.ready;
    assert.match(errorPage.nodes.get('articleBody').children[0].textContent, /Thiếu đường dẫn/);
  }
  assert.equal(new URL(detail.context.blogArticleUrl('farmers', 'a&b'), 'https://example.test').searchParams.get('slug'), 'a&b');
  console.log('Blog startup, views, navigation, pagination, errors and legacy redirect tests passed (mock DOM/API)');
})().catch(error => { console.error(error); process.exitCode = 1; });
