const BLOG_API_URL = 'https://script.google.com/macros/s/AKfycbx6HqvTfvToQnn61_8d2dtIIf57SlKvL_-XCi32O-HVWWCIwlQ6f9kvbas04guhiGt-/exec';
const BLOG_CATEGORIES = { business: 'The Business', farmers: 'The Farmers', processing: 'The Processing' };
const blogParams = new URLSearchParams(location.search);
const blogCategory = Object.hasOwn(BLOG_CATEGORIES, blogParams.get('category')) ? blogParams.get('category') : 'farmers';

function blogListUrl(category, page = 1) {
  return `/blog/?category=${encodeURIComponent(category)}&page=${page}`;
}
function blogArticleUrl(category, slug) {
  return `/blog/article.html?category=${encodeURIComponent(category)}&slug=${encodeURIComponent(slug)}`;
}
function setText(id, value) { document.getElementById(id).textContent = value; }
function blogState(element, message) { element.replaceChildren(); const box = document.createElement('div'); box.className = 'state'; box.textContent = message; element.append(box); }
function imageUrl(value) { return /^https:\/\/[^\s"'<>]+$/i.test(String(value || '')) ? value : ''; }

async function blogFetch(action, params) {
  const url = new URL(BLOG_API_URL);
  url.searchParams.set('action', action);
  Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));
  // [SỬA] Cache bài công khai hai phút trong phiên; dữ liệu nháp không được lưu.
  const cacheKey = `oriberry_blog_v1:${url.search}`;
  try {
    const cached = JSON.parse(sessionStorage.getItem(cacheKey) || 'null');
    if (cached && Date.now() - cached.savedAt < 120000) return cached.data;
  } catch (_) { /* Storage có thể bị trình duyệt chặn. */ }
  const response = await fetch(url.toString());
  if (!response.ok) throw new Error('Không thể kết nối máy chủ.');
  const data = await response.json();
  if (data.status !== 'success') throw new Error(data.message || 'Không tải được bài viết.');
  try { sessionStorage.setItem(cacheKey, JSON.stringify({savedAt:Date.now(),data})); } catch (_) { /* Optional cache. */ }
  return data;
}

function addCategoryNavigation() {
  const nav = document.getElementById('categoryNav');
  if (!nav) return;
  Object.entries(BLOG_CATEGORIES).forEach(([id, name]) => {
    const link = document.createElement('a');
    link.href = blogListUrl(id);
    link.textContent = name;
    if (id === blogCategory) link.setAttribute('aria-current', 'page');
    nav.append(link);
  });
}

function addPagination(data) {
  const nav = document.getElementById('pagination');
  nav.replaceChildren();
  if (data.pages < 2) return;
  const add = (label, page, current = false) => {
    const node = document.createElement(page ? 'a' : 'span');
    node.textContent = label;
    if (page) node.href = blogListUrl(blogCategory, page);
    if (current) { node.className = 'current'; node.setAttribute('aria-current', 'page'); }
    if (!page && !current) node.className = 'disabled';
    nav.append(node);
  };
  add('Previous', data.page > 1 ? data.page - 1 : null);
  for (let page = 1; page <= data.pages; page++) add(String(page), page === data.page ? null : page, page === data.page);
  add('Next', data.page < data.pages ? data.page + 1 : null);
}

async function loadBlogList() {
  addCategoryNavigation();
  const title = BLOG_CATEGORIES[blogCategory];
  setText('categoryTitle', title);
  document.title = `${title} | Oriberry Journal`;
  const container = document.getElementById('postGrid');
  const rawPage = Number(blogParams.get('page') || 1);
  const page = Number.isInteger(rawPage) && rawPage > 0 ? rawPage : 1;
  try {
    const data = await blogFetch('getBlogPosts', { category: blogCategory, page });
    container.replaceChildren();
    if (!data.posts.length) blogState(container, data.total ? 'Trang này chưa có bài viết.' : 'Chuyên mục này chưa có bài viết được xuất bản.');
    data.posts.forEach(post => {
      const card = document.createElement('article'); card.className = 'post-card';
      const link = document.createElement('a'); link.href = blogArticleUrl(blogCategory, post.slug);
      const img = document.createElement('img'); img.src = imageUrl(post.coverImageUrl); img.alt = post.title; img.loading = 'lazy';
      const body = document.createElement('div'); body.className = 'post-card-content';
      const heading = document.createElement('h2'); heading.textContent = post.title;
      const summary = document.createElement('p'); summary.textContent = post.excerpt;
      body.append(heading, summary); link.append(img, body); card.append(link); container.append(card);
    });
    addPagination(data);
  } catch (error) { blogState(container, error.message || 'Không thể tải bài viết. Vui lòng thử lại.'); }
}

// [SỬA] Main_Content là văn bản/Markdown giới hạn; chỉ tạo DOM bằng textContent, không chèn HTML từ Sheets.
function renderArticleContent(container, source) {
  container.replaceChildren();
  let list = null;
  String(source || '').split(/\r?\n/).forEach(raw => {
    const line = raw.trim();
    if (!line) { list = null; return; }
    const image = /^!\[([^\]]*)\]\((https:\/\/[^\s"'<>]+)\)$/.exec(line);
    if (image) {
      const img = document.createElement('img'); img.src = imageUrl(image[2]); img.alt = image[1]; img.loading = 'lazy'; container.append(img); list = null;
    } else if (/^###? /.test(line)) {
      const heading = document.createElement(line.startsWith('### ') ? 'h3' : 'h2');
      heading.textContent = line.replace(/^###? /, ''); container.append(heading); list = null;
    } else if (/^- /.test(line)) {
      if (!list) { list = document.createElement('ul'); container.append(list); }
      const item = document.createElement('li'); item.textContent = line.slice(2); list.append(item);
    } else {
      const paragraph = document.createElement('p'); paragraph.textContent = line; container.append(paragraph); list = null;
    }
  });
}

async function loadBlogArticle() {
  const back = document.getElementById('backToCategory'); back.href = blogListUrl(blogCategory);
  back.textContent = `← ${BLOG_CATEGORIES[blogCategory]}`;
  const slug = blogParams.get('slug') || '';
  const container = document.getElementById('articleBody');
  if (!slug) return blogState(container, 'Thiếu đường dẫn bài viết.');
  try {
    const { post } = await blogFetch('getBlogPost', { category: blogCategory, slug });
    setText('articleCategory', BLOG_CATEGORIES[blogCategory]);
    setText('articleTitle', post.title);
    setText('articleExcerpt', post.excerpt);
    document.title = `${post.title} | Oriberry Journal`;
    document.querySelector('meta[name="description"]').content = post.excerpt;
    const cover = document.getElementById('articleCover'); cover.src = imageUrl(post.coverImageUrl); cover.alt = post.title; cover.hidden = !cover.src;
    renderArticleContent(container, post.content);
  } catch (error) { blogState(container, error.message || 'Không tìm thấy bài viết.'); }
}

if (document.getElementById('postGrid')) loadBlogList();
if (document.getElementById('articleBody')) loadBlogArticle();
