// index.js

import { modifythefilenameFunction } from './file_tools/modifythefilename.js'
import { characterFunction } from './file_tools/character.js'
import { imageFunction } from './file_tools/image.js'
import { sortFunction } from './file_tools/sort.js'
import { exportFunction } from './file_tools/export.js'
import { collect_fileFunction } from './file_tools/collect_file.js'
import { copy_folderFunction } from './file_tools/copy_folder.js'
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

window.file_tools_modifythefilenameFunction = function() {
    modifythefilenameFunction();
}

window.file_tools_characterFunction = function() {
    characterFunction();
}

window.file_tools_imageFunction = function() {
    imageFunction();
}

window.file_tools_exportFunction = function() {
    exportFunction();
}

window.file_tools_sortFunction = function() {
    sortFunction();
}

window.file_tools_collect_fileFunction = function() {
    collect_fileFunction();
}

window.file_tools_copy_folderFunction = function() {
    copy_folderFunction();
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