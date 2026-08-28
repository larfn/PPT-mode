// 预置图源模板（纯数据，无依赖）：默认启用源 + 预置模板库
// 字段约定：
//   endpoint 支持占位符 {query} {count} {page} {start} {key}（自动 URL 编码；{key} 在 headers 值中原样替换）
//   resultsPath / fields 支持点分路径与数组下标，如 'data.items'、'imageinfo[0].url'
'use strict';

// 首次安装即默认启用的源（随配置保存，可编辑/删除）
const sogouDef = {
  id: 'sogou_page',
  name: '搜狗图片',
  note: '免 Key；接口偶尔要求 Cookie，若返回 forbid 可在「请求头」或「Cookies」中补充（如 SNUID）',
  keyRequired: false,
  endpoint: 'https://pic.sogou.com/napi/pc/searchList?mode=1&start={start}&xml_len={count}&query={query}',
  headers: { Referer: 'https://pic.sogou.com/' },
  cookies: {},
  resultsPath: 'data.items',
  fields: { imageUrl: 'picUrl', thumbnailUrl: 'thumbUrl', width: 'width', height: 'height', title: 'title', sourceUrl: 'pageUrl' },
  enabled: true
};

const defaultEnabledSources = [sogouDef];

// 预置模板库（图源管理页「一键添加」，默认不启用）；搜狗同时作为模板，便于删除后重新添加
const presets = [sogouDef,
  {
    id: 'openverse',
    name: 'Openverse（CC 图库，免 Key）',
    note: '免 Key；需能访问外网（国内网络可能不可达）',
    keyRequired: false,
    endpoint: 'https://api.openverse.org/v1/images/?q={query}&page_size={count}&page={page}',
    headers: {},
    cookies: {},
    resultsPath: 'results',
    fields: { imageUrl: 'url', thumbnailUrl: 'thumbnail', width: 'width', height: 'height', title: 'title', sourceUrl: 'foreign_landing_url', author: 'creator', license: 'license' },
    enabled: false
  },
  {
    id: 'wikimedia',
    name: '维基共享资源（免 Key）',
    note: '免 Key；需能访问外网（国内网络可能不可达）',
    keyRequired: false,
    endpoint: 'https://commons.wikimedia.org/w/api.php?action=query&generator=search&gsrsearch={query}&gsrnamespace=6&gsrlimit={count}&prop=imageinfo&iiprop=url|size|mime&iiurlwidth=480&format=json&formatversion=2&origin=*',
    headers: {},
    cookies: {},
    resultsPath: 'query.pages',
    fields: { imageUrl: 'imageinfo[0].url', thumbnailUrl: 'imageinfo[0].thumburl', width: 'imageinfo[0].width', height: 'imageinfo[0].height', title: 'title', mimeType: 'imageinfo[0].mime' },
    enabled: false
  },
  {
    id: 'pixabay',
    name: 'Pixabay（需 Key）',
    note: '需 Pixabay API Key（免费注册）',
    keyRequired: true,
    endpoint: 'https://pixabay.com/api/?key={key}&q={query}&per_page={count}&page={page}',
    headers: {},
    cookies: {},
    resultsPath: 'hits',
    fields: { imageUrl: 'webformatURL', thumbnailUrl: 'previewURL', width: 'imageWidth', height: 'imageHeight', title: 'tags', sourceUrl: 'pageURL', author: 'user' },
    enabled: false
  },
  {
    id: 'unsplash',
    name: 'Unsplash（需 Key）',
    note: '需 Unsplash API Key（Access Key）',
    keyRequired: true,
    endpoint: 'https://api.unsplash.com/search/photos?query={query}&per_page={count}&page={page}',
    headers: { Authorization: 'Client-ID {key}' },
    cookies: {},
    resultsPath: 'results',
    fields: { imageUrl: 'urls.regular', thumbnailUrl: 'urls.thumb', width: 'width', height: 'height', title: 'alt_description', sourceUrl: 'links.html', author: 'user.name' },
    enabled: false
  },
  {
    id: 'pexels',
    name: 'Pexels（需 Key）',
    note: '需 Pexels API Key（免费注册）',
    keyRequired: true,
    endpoint: 'https://api.pexels.com/v1/search?query={query}&per_page={count}&page={page}',
    headers: { Authorization: '{key}' },
    cookies: {},
    resultsPath: 'photos',
    fields: { imageUrl: 'src.large2x', thumbnailUrl: 'src.medium', width: 'width', height: 'height', title: 'alt', sourceUrl: 'url', author: 'photographer' },
    enabled: false
  }
];

module.exports = { defaultEnabledSources, presets };