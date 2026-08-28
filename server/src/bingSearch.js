// 兼容层：转发到 BingPageProvider（统一 ImageResult）
'use strict';
const { searchImages } = require('./providers/bingPageProvider.js');
module.exports = { searchImages };
