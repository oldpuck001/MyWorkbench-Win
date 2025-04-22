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
import { fillFunction } from './xlsx_xls_tools/fill.js';
import { regexFunction } from './xlsx_xls_tools/regex.js'
import { bank_statement_sortFunction } from './xlsx_xls_tools/bank_statement_sort.js';
import { generate_chronological_accountFunction } from './xlsx_xls_tools/generate_chronological_account.js';
import { data_cleaning_Function } from './data_analysis_tools/data_cleaning.js'
import { sql_sqlite_Function } from './data_analysis_tools/sql_sqlite.js'
import { text_comparisonFunction } from './other_tools/text_comparison.js'

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

document.addEventListener('DOMContentLoaded', () => {
    const sidebar = document.getElementById('sidebar');
    const toggleBtn = document.getElementById('toggleSidebarBtn');
    const content = document.getElementById('content');

    window.toggleSidebar = function () {
        if (sidebar.classList.contains('hidden')) {
            sidebar.classList.remove('hidden');
            sidebar.style.width = '250px';
            toggleBtn.style.left = '265px';
            content.style.marginLeft = '250px';
        } else {
            sidebar.classList.add('hidden');
            sidebar.style.width = '0';
            toggleBtn.style.left = '10px';
            content.style.marginLeft = '0';
        }
    };
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

window.excel_tools_fillFunction = function() {
    fillFunction();
}

window.excel_tools_regexFunction = function() {
    regexFunction();
}

window.excel_tools_bank_statement_sortFunction = function() {
    bank_statement_sortFunction();
}
window.excel_tools_generate_chronological_accountFunction = function() {
    generate_chronological_accountFunction();
}

window.data_analysis_tools_data_cleaningFunction = function() {
    data_cleaning_Function();
}

window.data_analysis_tools_sql_sqliteFunction = function() {
    sql_sqlite_Function();
}

window.other_tools_text_comparisonFunction = function() {
    text_comparisonFunction();
}