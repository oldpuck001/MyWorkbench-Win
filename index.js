// index.js

import { dashboardFunction, settingsFunction } from './countdown/script.js';

import { spliceFunction } from './xlsx_xls_tools/splice.js';
import { subtotalsFunction } from './xlsx_xls_tools/subtotals.js';

import { vocabulary_to_audio } from './study_language_tools/script.js';

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

window.pdf_tools_splitFunction = function() {
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.innerHTML = `<h1 style="text-align: center; width: 100%;">PDF文件分割功能开发中</h1>`;
}

window.vocabulary_to_audioFunction = function() {
    vocabulary_to_audio();
}