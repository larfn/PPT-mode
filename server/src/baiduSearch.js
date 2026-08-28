// 兼容层：转发到 BaiduPageProvider（统一 ImageResult）
'use strict';
const { searchImages } = require('./providers/baiduPageProvider.js');
module.exports = { searchImages };
