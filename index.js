// index.js

import { modifythefilenameFunction } from './filename_tools/modifythefilename.js'
import { characterFunction } from './filename_tools/character.js'
import { imageFunction } from './filename_tools/image.js'
import { sortFunction } from './filename_tools/sort.js'
import { exportFunction } from './filename_tools/export.js'

import { spliceFunction } from './xlsx_xls_tools/splice.js';
import { subtotalsFunction } from './xlsx_xls_tools/subtotals.js';

import { single_sort_exportFunction } from './data_analysis_tools/single_sort_export.js';
import { bank_statement_sortFunction } from './data_analysis_tools/bank_statement_sort.js';

document.querySelectorAll('.sidebar > ul > li').forEach(item => {
    item.addEventListener('click', function(e) {
        // 检查是否有子菜单
        const sublist = this.querySelector('ul');
        if (sublist) {
            e.stopPropagation();
            sublist.classList.toggle('active');
        }
    });
});

document.querySelectorAll('.sidebar ul ul li').forEach(item => {
    item.addEventListener('click', function(e) {
        e.stopPropagation();
    });
});

// 页面加载后自动显示的页面
window.addEventListener('DOMContentLoaded', () => {
    modifythefilenameFunction();
});

window.filename_tools_modifythefilenameFunction = function() {
    modifythefilenameFunction();
}

window.filename_tools_characterFunction = function() {
    characterFunction();
}

window.filename_tools_imageFunction = function() {
    imageFunction();
}

window.filename_tools_exportFunction = function() {
    exportFunction();
}

window.filename_tools_sortFunction = function() {
    sortFunction();
}








window.excel_tools_spliceFunction = function() {
    spliceFunction();
}

window.excel_tools_subtotalsFunction = function() {
    subtotalsFunction();
}

window.excel_tools_single_sort_exportFunction = function() {
    single_sort_exportFunction();
}

window.excel_tools_bank_statement_sortFunction = function() {
    bank_statement_sortFunction();
}