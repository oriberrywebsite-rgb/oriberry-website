const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const headers = ['ID','Category','Title','Excerpt','Main_Content','Cover_Image_URL','Target_Page','Status','Slug','Published_At','Sort_Order'];
const cover = 'https://example.test/coffee.jpg';
const rows = [headers];
for (let n = 1; n <= 7; n++) rows.push([n,'farmers',`Story ${n}`,`Summary ${n}`,`## Section ${n}\n- Detail ${n}`,cover,'','Published',`story-${n}`,`2026-10-${String(n).padStart(2,'0')}`,n]);
rows.push([8,'farmers','Draft','Hidden','Hidden',cover,'','Draft','draft','2026-10-08',8]);
rows.push([9,'farmers','No cover','Hidden','Hidden','test','','Active','no-cover','2026-10-09',9]);
const sheet = values => ({getLastRow:()=>values.length,getLastColumn:()=>values[0].length,getRange:()=>({getValues:()=>values.map(row => [...row])})});
const sheets = {
  Block1_Farmers_Philosophy: sheet(rows),
  Block2_Roastery_Products: sheet([headers,[1,'processing','Processing','Summary','Full story',cover,'','Active','processing-story','2026-10-01',1]]),
  Block3_Shops_Menu: sheet([headers,[1,'business','Business','Summary','Full story',cover,'','Active','business-story','2026-10-01',1]])
};
const context = vm.createContext({
  SpreadsheetApp: {getActiveSpreadsheet:()=>({getSheetByName:name=>sheets[name]})},
  ContentService: {MimeType:{JSON:'JSON'},createTextOutput(content){return {content,setMimeType(){return this;}}}}
});
vm.runInContext(fs.readFileSync('backend/Blog.gs','utf8'),context);
vm.runInContext(fs.readFileSync('backend/Code.gs','utf8'),context);
const get = params => JSON.parse(context.doGet({parameter:params}).content);
for (const category of ['business','farmers','processing']) assert.equal(get({action:'getBlogPosts',category}).status,'success');
const first = get({action:'getBlogPosts',category:'farmers',page:'1'});
assert.equal(first.total,7); assert.equal(first.pages,3); assert.equal(first.posts.length,3);
assert.deepEqual(first.posts.map(p=>p.slug),['story-1','story-2','story-3']);
assert.equal(get({action:'getBlogPosts',category:'farmers',page:'2'}).posts.length,3);
assert.equal(get({action:'getBlogPosts',category:'farmers',page:'3'}).posts.length,1);
assert.equal(get({action:'getBlogPosts',category:'farmers',page:'4'}).posts.length,0);
assert.equal(get({action:'getBlogPosts',category:'farmers',page:'0'}).code,'INVALID_PAGE');
assert.equal(get({action:'getBlogPosts',category:'unknown'}).code,'INVALID_CATEGORY');
assert.match(get({action:'getBlogPost',category:'farmers',slug:'story-1'}).post.content,/Detail 1/);
assert.equal(get({action:'getBlogPost',category:'farmers',slug:'draft'}).code,'NOT_FOUND');
assert.equal(get({action:'getBlogPost',category:'farmers',slug:'missing'}).code,'NOT_FOUND');

const node = tag => ({tag,textContent:'',children:[],append(...items){this.children.push(...items);},replaceChildren(){this.children=[];}});
const document = {getElementById:()=>null,createElement:node};
const frontend = vm.createContext({document,location:{search:'?category=farmers'},URLSearchParams,URL,fetch(){throw Error('No network')},console});
vm.runInContext(fs.readFileSync('blog/blog.js','utf8'),frontend);
const container = node('section');
frontend.renderArticleContent(container,'## Heading\n- item\n<script>alert(1)</script>\n![photo](https://example.test/photo.jpg)');
assert.deepEqual(container.children.map(item=>item.tag),['h2','ul','p','img']);
assert.equal(container.children[2].textContent,'<script>alert(1)</script>');
assert.equal(frontend.imageUrl('javascript:alert(1)'),'');
console.log('Blog CMS and safe rendering tests passed (mock Sheets and DOM)');
