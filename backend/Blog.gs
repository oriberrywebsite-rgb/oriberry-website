// [SỬA] Ba mục Stories lấy bài từ đúng ba tab CMS hiện có; không đọc tab đơn hàng.
var BLOG_CATEGORY_SHEETS = {
  business: "Block3_Shops_Menu",
  farmers: "Block1_Farmers_Philosophy",
  processing: "Block2_Roastery_Products"
};

function blogJson_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function blogRows_(category) {
  var sheetName = BLOG_CATEGORY_SHEETS[category];
  if (!sheetName) return null;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var header = rows.shift().map(function(value) { return String(value || "").trim().toLowerCase(); });
  var column = function(name) { return header.indexOf(name.toLowerCase()); };
  var posts = [];
  var seenSlugs = {};
  rows.forEach(function(row) {
    var read = function(name) { var index = column(name); return index < 0 ? "" : row[index]; };
    var status = String(read("Status") || "").trim().toLowerCase();
    var slug = String(read("Slug") || "").trim().toLowerCase();
    var title = String(read("Title") || "").trim();
    var excerpt = String(read("Excerpt") || "").trim();
    var content = String(read("Main_Content") || "").trim();
    var cover = String(read("Cover_Image_URL") || "").trim();
    // [SỬA] Hàng nháp hoặc dữ liệu mẫu "test" thiếu URL ảnh/slug không lên website.
    if (["active", "published"].indexOf(status) < 0 || !/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(slug) ||
        seenSlugs[slug] || !title || !excerpt || !content ||
        [title, excerpt, content].some(function(value) { return /^test\s*$/i.test(value); }) ||
        !/^https:\/\/[^\s"'<>]+$/i.test(cover)) return;
    seenSlugs[slug] = true;
    var published = read("Published_At");
    var publishedAt = published instanceof Date && !isNaN(published.getTime()) ? published.toISOString() : String(published || "").trim();
    var sort = read("Sort_Order");
    posts.push({ id: String(read("ID") || ""), slug: slug, category: category, title: title,
      excerpt: excerpt, content: content, coverImageUrl: cover, publishedAt: publishedAt,
      sortOrder: typeof sort === "number" && isFinite(sort) ? sort : null });
  });
  posts.sort(function(a, b) {
    if (a.sortOrder !== null || b.sortOrder !== null) return (a.sortOrder === null ? Infinity : a.sortOrder) - (b.sortOrder === null ? Infinity : b.sortOrder);
    return b.publishedAt.localeCompare(a.publishedAt) || a.slug.localeCompare(b.slug);
  });
  return posts;
}

function getBlogPostsData(category, requestedPage) {
  var posts = blogRows_(String(category || "").trim().toLowerCase());
  if (posts === null) return { status: "error", code: "INVALID_CATEGORY", message: "Chuyên mục không hợp lệ." };
  var page = Number(requestedPage || 1);
  if (!Number.isInteger(page) || page < 1 || page > 100000) return { status: "error", code: "INVALID_PAGE", message: "Số trang không hợp lệ." };
  var pages = Math.ceil(posts.length / 3);
  return { status: "success", category: category, page: page, pages: pages, total: posts.length,
    posts: posts.slice((page - 1) * 3, page * 3).map(function(post) {
      return { slug: post.slug, category: post.category, title: post.title, excerpt: post.excerpt,
        coverImageUrl: post.coverImageUrl, publishedAt: post.publishedAt };
    }) };
}

function getBlogPostData(category, slug) {
  var posts = blogRows_(String(category || "").trim().toLowerCase());
  if (posts === null) return { status: "error", code: "INVALID_CATEGORY", message: "Chuyên mục không hợp lệ." };
  var target = String(slug || "").trim().toLowerCase();
  var post = posts.filter(function(item) { return item.slug === target; })[0];
  return post ? { status: "success", post: post } : { status: "error", code: "NOT_FOUND", message: "Không tìm thấy bài viết." };
}
