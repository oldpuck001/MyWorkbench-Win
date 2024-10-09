// index.js

import { dashboardFunction, settingsFunction } from './countdown/countdown.js';

import { spliceFunction } from './xlsx_xls_tools/splice.js';
import { subtotalsFunction } from './xlsx_xls_tools/subtotals.js';
import { single_sort_exportFunction } from './xlsx_xls_tools/single_sort_export.js';

//PDF

import { china_mainland_annual_auditFunction } from './accounting_audit/china_mainland_annual_audit/china_mainland_annual_audit.js';

import { modifythefilenameFunction } from './filename_tools/modifythefilename.js'
import { characterFunction } from './filename_tools/character.js'
import { imageFunction } from './filename_tools/image.js'
import { sortFunction } from './filename_tools/sort.js'
import { exportFunction } from './filename_tools/export.js'

import { vocabulary_to_audioFunction } from './study_language_tools/vocabulary_to_audio.js';

document.querySelectorAll('.sidebar > ul > li').forEach(item => {
    item.addEventListener('click', function(e) {
        // 检查是否有子菜单
        const sublist = this.querySelector('ul');
        if (sublist) {
            // 如果有子菜单，则展开/隐藏子菜单
            e.stopPropagation();
            sublist.classList.toggle('active');
            sublist.style.display = sublist.style.display === 'block' ? 'none' : 'block';
        } else {
            // 如果没有子菜单，则加载相应的页面
            const src = this.getAttribute('data-src');
            if (src) {
                e.stopPropagation();
                document.getElementById('mainFrame').setAttribute('src', src);
            }
        }
    });
});

document.querySelectorAll('.sidebar ul ul li').forEach(item => {
    item.addEventListener('click', function(e) {
        e.stopPropagation();
        const src = this.getAttribute('data-src');
        if (src) {
            document.getElementById('mainFrame').setAttribute('src', src);
        }
    });
});

const sidebar = document.querySelector('.sidebar');
const content = document.querySelector('.content');
const toggleBtn = document.querySelector('.toggle-btn');
toggleBtn.addEventListener('click', () => {
    sidebar.classList.toggle('hidden');
});

window.instructionFunction = function() {
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.innerHTML = `<h1 style="text-align: center; width: 100%;">使用说明</h1>`;
}

window.countdown_dashboardFunction = function() {
    dashboardFunction();
}

window.countdown_settingsFunction = function() {
    settingsFunction();
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

window.pdf_tools_splitFunction = function() {
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.innerHTML = `<h1 style="text-align: center; width: 100%;">PDF文件分割功能开发中</h1>`;
}

window.accounting_audit_china_mainland_annual_auditFunction = function() {
    china_mainland_annual_auditFunction();
}

window.filename_tools_modifythefilenameFunction = function() {
    modifythefilenameFunction();
}

window.filename_tools_characterFunction = function() {
    characterFunction();
}

window.filename_tools_imageFunction = function() {
    imageFunction();
}

window.filename_tools_sortFunction = function() {
    sortFunction();
}

window.filename_tools_exportFunction = function() {
    exportFunction();
}

window.study_language_tools_vocabulary_to_audioFunction = function() {
    vocabulary_to_audioFunction();
}

window.video_editing_tools_text_to_audioFunction = function() {
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.innerHTML = `<h1 style="text-align: center; width: 100%;">文本转语音功能开发中</h1>`;
}

window.video_editing_tools_cc_toolFunction = function() {
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.innerHTML = `<h1 style="text-align: center; width: 100%;">CC字幕辅助功能开发中</h1>`;
}