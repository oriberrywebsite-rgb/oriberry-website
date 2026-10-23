const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function makeNode(tag) {
  return {
    tag, children: [], textContent: '', className: '', attributes: {},
    append(...nodes) { this.children.push(...nodes); },
    replaceChildren() { this.children = []; },
    setAttribute(name, value) { this.attributes[name] = value; }
  };
}

function makePage(search, mode, response) {
  const nodes = new Map();
  const ids = mode === 'list'
    ? ['categoryNav','categoryTitle','postGrid','pagination']
    : ['backToCategory','articleBody','articleCategory','articleTitle','articleExcerpt','articleCover'];
  ids.forEach(id => nodes.set(id, makeNode(id)));
  const metadata = {content:''};
  let ready = false;
  const document = {
    title: '',
    getElementById(id) { return ready ? (nodes.get(id) || null) : null; },
    createElement: makeNode,
    querySelector(selector) { return selector === 'meta[name="description"]' ? metadata : null; }
  };
  const saved = new Map();
  const requests = [];
  const context = vm.createContext({
    document, location:{search}, URLSearchParams, URL,
    sessionStorage:{getItem:key=>saved.get(key)||null,setItem:(key,value)=>saved.set(key,value)},
    fetch:async url=>{requests.push(String(url));return {ok:true,json:async()=>response};},
    console
  });
  vm.runInContext(fs.readFileSync('blog/blog.js','utf8'),context,{filename:'blog/blog.js'});
  ready = true;
  return {context,nodes,requests,document,metadata};
}

(async () => {
  const posts = Array.from({length:3},(_,index)=>({slug:`story-${index+1}`,category:'farmers',title:`Story ${index+1}`,excerpt:`Summary ${index+1}`,coverImageUrl:'https://example.test/photo.jpg'}));
  const page = makePage('?category=farmers&page=1','list',{status:'success',page:1,pages:2,total:4,posts});
  await page.context.loadBlogList();
  assert.equal(page.requests.length,1);
  assert.match(page.requests[0],/action=getBlogPosts/);
  assert.equal(page.nodes.get('categoryNav').children.length,3);
  assert.equal(page.nodes.get('postGrid').children.length,3,'list renders exactly three posts');
  assert.equal(page.nodes.get('postGrid').children[0].children[0].href,'/blog/article.html?category=farmers&slug=story-1');
  assert.equal(page.nodes.get('pagination').children[0].textContent,'Previous');
  assert.equal(page.nodes.get('pagination').children[0].tag,'span');
  assert.equal(page.nodes.get('pagination').children.at(-1).href,'/blog/?category=farmers&page=2');
  assert.equal(page.document.title,'The Farmers | Oriberry Journal');

  const second = makePage('?category=business&page=2','list',{status:'success',page:2,pages:2,total:4,posts:[posts[0]]});
  await second.context.loadBlogList();
  assert.equal(second.nodes.get('postGrid').children.length,1);
  assert.equal(second.nodes.get('pagination').children[0].href,'/blog/?category=business&page=1');
  assert.equal(second.nodes.get('pagination').children.at(-1).tag,'span');

  const detail = makePage('?category=processing&slug=processing-story','article',{status:'success',post:{title:'Processing Story',excerpt:'A short introduction',coverImageUrl:'https://example.test/photo.jpg',content:'## Heading\n- First point\nPlain paragraph'}});
  await detail.context.loadBlogArticle();
  assert.match(detail.requests[0],/action=getBlogPost/);
  assert.equal(detail.nodes.get('backToCategory').href,'/blog/?category=processing&page=1');
  assert.equal(detail.nodes.get('articleTitle').textContent,'Processing Story');
  assert.deepEqual(detail.nodes.get('articleBody').children.map(node=>node.tag),['h2','ul','p']);
  assert.equal(detail.metadata.content,'A short introduction');

  const missing = makePage('?category=processing&slug=missing','article',{status:'error',code:'NOT_FOUND',message:'Không tìm thấy bài viết.'});
  await missing.context.loadBlogArticle();
  assert.equal(missing.nodes.get('articleBody').children[0].className,'state');
  assert.match(missing.nodes.get('articleBody').children[0].textContent,/Không tìm thấy/);
  console.log('Blog list, pagination, direct article and error UI tests passed (mock DOM/API)');
})().catch(error=>{console.error(error);process.exitCode=1;});
