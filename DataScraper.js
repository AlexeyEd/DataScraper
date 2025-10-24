// ==UserScript==
// @name         Google Sheets Data Scraper
// @namespace    http://tampermonkey.net/
// @version      6.4
// @description  @description RU: Последовательный сбор данных с сайтов: открывает сайты по порядку, автоматически следует по переадресациям, извлекает нужные элементы по CSS-селекторам и копирует их в буфер обмена. Поддерживает динамические regex-замены в URL и в результатах (редактируются прямо в интерфейсе), мгновенное сохранение, применение к конкретным сайтам, перетаскивание сайтов и секций для изменения порядка. Работает в Google Таблицах — запускается одной кнопкой, показывает прогресс, защищён от зависаний. EN: Sequential data collection from websites: opens sites in order, automatically follows redirects, extracts required elements via CSS selectors, and copies them to the clipboard. Supports dynamic regex replacements in URLs and results (editable directly in the interface), instant saving, application to specific sites, drag-and-drop reordering of sites and sections. Runs in Google Sheets — starts with one button, shows progress, protected from hangs.
// @author       You
// @match        https://docs.google.com/spreadsheets/*
// @match        https://*/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_openInTab
// @grant        GM_setClipboard
// @grant        GM_addValueChangeListener
// @grant        GM_removeValueChangeListener
// ==/UserScript==

(function() {
    'use strict';

    const isGoogleSheets = window.location.href.includes('docs.google.com/spreadsheets');

    /* ----------------- Helpers for parsing sites string ----------------- */
    function parseSites(str) {
        if (!str || !str.trim()) return [];
        const parts = str.split(',').map(p => p.trim()).filter(p => p);
        const nums = [];
        for (let part of parts) {
            if (part.includes('-')) {
                const [startStr, endStr] = part.split('-');
                const start = parseInt(startStr.trim());
                const end = parseInt(endStr.trim());
                if (!isNaN(start) && !isNaN(end) && start <= end && start > 0) {
                    for (let i = start; i <= end; i++) nums.push(i);
                }
            } else {
                const n = parseInt(part);
                if (!isNaN(n) && n > 0) nums.push(n);
            }
        }
        return [...new Set(nums)]; // unique
    }

    /* ----------------- Generic Replacements Manager ----------------- */
    function createReplacementsManager(storageKey, applyWarningPrefix) {
        return {
            get() {
                const raw = GM_getValue(storageKey);
                try { return raw ? JSON.parse(raw) : []; } catch (e) { GM_setValue(storageKey, JSON.stringify([])); return []; }
            },
            save(rules) {
                GM_setValue(storageKey, JSON.stringify(rules));
            },
            add(findPattern = '', replaceValue = '') {
                const rules = this.get();
                rules.push({ find: findPattern, replace: replaceValue, enabled: true, sites: '' });
                this.save(rules);
                return rules.length - 1;
            },
            update(index, findPattern, replaceValue, enabled, sites = '') {
                const rules = this.get();
                if (index < 0 || index >= rules.length) return;
                rules[index].find = findPattern;
                rules[index].replace = replaceValue;
                rules[index].enabled = !!enabled;
                rules[index].sites = sites;
                this.save(rules);
            },
            delete(index) {
                const rules = this.get();
                if (index < 0 || index >= rules.length) return;
                rules.splice(index, 1);
                this.save(rules);
            },
            apply(original, siteNum) {
                let text = original;
                const rules = this.get();
                for (let i = 0; i < rules.length; i++) {
                    const r = rules[i];
                    if (!r || !r.enabled || typeof r.find !== 'string' || r.find.trim() === '') continue;
                    const applicable = !r.sites || r.sites.trim() === '' || parseSites(r.sites).includes(siteNum);
                    if (!applicable) continue;
                    try {
                        const re = new RegExp(r.find, 'g');
                        text = text.replace(re, r.replace == null ? '' : String(r.replace));
                    } catch (e) {
                        console.warn(applyWarningPrefix + ' replacement rule #' + i + ' invalid RegExp "' + r.find + '" — skipped', e);
                    }
                }
                return text;
            }
        };
    }

    const urlReplacements = createReplacementsManager('urlReplacements', 'URL');
    const resultReplacements = createReplacementsManager('resultReplacements', 'Result');

    /* ----------------- Sites storage ----------------- */
    function getDefaultSites() {
        return [
            {
                name: 'example.com',
                url: 'https://example.com/',
                enabled: true,
                collapsed: false,
                redirectSelectors: [],
                selectors: [
                    {enabled: true, selector: 'body > div > h1'}
                ]
            }
        ];
    }
    function getSites() {
        const saved = GM_getValue('sites');
        try {
            let sites = saved ? JSON.parse(saved) : getDefaultSites();
            sites.forEach(site => {
                if (site.enabled === undefined) site.enabled = true;
                if (site.collapsed === undefined) site.collapsed = true;
                if (site.redirectSelector) {
                    site.redirectSelectors = [site.redirectSelector];
                    delete site.redirectSelector;
                }
                if (!site.redirectSelectors) site.redirectSelectors = [];
                // Migrate selectors to objects if they are strings
                if (Array.isArray(site.selectors) && site.selectors.every(s => typeof s === 'string')) {
                    site.selectors = site.selectors.map(s => ({enabled: true, selector: s}));
                }
                // Ensure selectors are array of objects
                if (!Array.isArray(site.selectors)) site.selectors = [];
                site.selectors.forEach(sel => {
                    if (typeof sel === 'string') {
                        sel = {enabled: true, selector: sel};
                    } else if (!sel.selector) {
                        sel.selector = '';
                    }
                });
            });
            return sites;
        } catch (e) { GM_setValue('sites', JSON.stringify(getDefaultSites())); return getDefaultSites(); }
    }
    function saveSites(sites) {
        GM_setValue('sites', JSON.stringify(sites));
    }

    function generateUniqueSiteName(sites, index, url) {
        try {
            const hostname = new URL(url).hostname;
            const hostnameCounts = {};
            sites.forEach((site, i) => {
                if (i !== index && site.url) {
                    try {
                        const h = new URL(site.url).hostname;
                        hostnameCounts[h] = (hostnameCounts[h] || 0) + 1;
                    } catch (e) {}
                }
            });
            const count = hostnameCounts[hostname] || 0;
            return count > 0 ? hostname + ' (' + (count + 1) + ')' : hostname;
        } catch (e) {
            return 'Untitled';
        }
    }

    /* ----------------- Generic UI helpers ----------------- */
    function createInput(type, value, placeholder, style, events) {
        const input = document.createElement('input');
        input.type = type;
        input.value = value || '';
        input.placeholder = placeholder || '';
        Object.assign(input.style, style || {});
        if (events) {
            Object.entries(events).forEach(([event, handler]) => input.addEventListener(event, handler));
        }
        return input;
    }

    function createButton(text, className, style, onclick) {
        const btn = document.createElement('button');
        btn.textContent = text;
        btn.className = className || '';
        Object.assign(btn.style, style || {});
        if (onclick) btn.onclick = onclick;
        return btn;
    }

    function validateRegexInline(inputEl, errorDiv) {
        const value = inputEl && inputEl.value || '';
        const trimmed = value.trim();
        if (trimmed === '') {
            if (errorDiv) {
                errorDiv.textContent = 'Паттерн regex пустой';
                errorDiv.style.display = 'block';
            }
            inputEl.style.borderColor = '#b00020';
            return false;
        }
        try {
            new RegExp(value);
            if (errorDiv) errorDiv.style.display = 'none';
            inputEl.style.borderColor = '';
            return true;
        } catch (e) {
            if (errorDiv) { errorDiv.textContent = 'Некорректный regex: ' + e.message; errorDiv.style.display = 'block'; }
            inputEl.style.borderColor = '#b00020';
            return false;
        }
    }

    /* ----------------- Generic UI for Replacement Section ----------------- */
    function createReplacementSection(type, collapsedKey, title) {
        const block = document.createElement('div');
        block.id = type + '-replace-block';
        block.style.marginBottom = '5px';
        block.className = 'reorderable-section';

        const sectionHeader = document.createElement('div');
        sectionHeader.className = 'section-header';

        const toggle = document.createElement('button');
        toggle.className = 'section-toggle';
        const collapsed = GM_getValue(collapsedKey, true);
        toggle.textContent = collapsed ? '+' : '−';

        const sectionTitle = document.createElement('div');
        sectionTitle.className = 'section-title';
        sectionTitle.style.fontWeight = '700';
        sectionTitle.textContent = title;

        sectionHeader.appendChild(toggle);
        sectionHeader.appendChild(sectionTitle);

        const content = document.createElement('div');
        content.id = type + '-replace-content';
        content.style.display = collapsed ? 'none' : 'block';

        // Quick add row
        const quickRow = document.createElement('div');
        quickRow.className = type + '-rule-row';
        quickRow.style.flexWrap = 'wrap';
        quickRow.style.marginBottom = '14px';

        const quickFind = createInput('text', '', 'Найти (regex)', {flex: '1 1 50%', padding: '4px 8px', fontFamily: 'monospace', fontSize: '13px'}, {});
        quickFind.id = type + 'replace-find-quick';

        const quickReplace = createInput('text', '', 'Заменить на', {flex: '1 1 50%', padding: '4px 8px', fontFamily: 'monospace', fontSize: '13px'});
        quickReplace.id = type + 'replace-to-quick';

        const quickAddBtn = createButton('Добавить', 'rule-add-btn', {padding: '6px 8px'}, function() {
            const find = quickFind.value;
            const replace = quickReplace.value;
            if (!find.trim()) { alert('Паттерн regex обязателен'); quickFind.focus(); return; }
            try { new RegExp(find); } catch (e) { alert('Некорректный regex: ' + e.message); quickFind.focus(); return; }
            const manager = type === 'url' ? urlReplacements : resultReplacements;
            const newIndex = manager.add(find, replace);
            setTimeout(() => {
                const inlineFind = document.querySelector(`#${type}-rule-find-${newIndex}`);
                if (inlineFind) inlineFind.focus();
                renderReplacementsUI(type);
            }, 50);
            quickFind.value = '';
            quickReplace.value = '';
        });

        const actions = document.createElement('div');
        actions.className = 'rule-actions';
        actions.appendChild(quickAddBtn);

        quickRow.appendChild(quickFind);
        quickRow.appendChild(quickReplace);
        quickRow.appendChild(actions);

        const rulesList = document.createElement('div');
        rulesList.id = type + '-replace-rules';
        rulesList.style.display = 'flex';
        rulesList.style.flexDirection = 'column';
        rulesList.style.gap = '8px';

        content.appendChild(quickRow);
        content.appendChild(rulesList);

        block.appendChild(sectionHeader);
        block.appendChild(content);

        const toggleFn = function() {
            const isCollapsed = content.style.display === 'none';
            content.style.display = isCollapsed ? 'block' : 'none';
            toggle.textContent = isCollapsed ? '−' : '+';
            GM_setValue(collapsedKey, !isCollapsed);
        };
        toggle.addEventListener('click', function(e) {
            e.stopPropagation();
            toggleFn();
        });
        sectionHeader.addEventListener('click', toggleFn);

        return block;
    }

    /* ----------------- Generic Render Replacements UI ----------------- */
    function createRuleRow(type, idx, rule, manager) {
        const rowWrap = document.createElement('div');
        rowWrap.className = type + '-rule-row';
        rowWrap.style.flexWrap = 'wrap';

        const enabledCheckbox = createInput('checkbox', '', '', {flex: '0 0 18px'}, {
            change: function() {
                manager.update(idx, findInput.value, replaceInput.value, enabledCheckbox.checked, sitesInput.value);
                validateRegexInline(findInput, errorDiv);
            }
        });
        enabledCheckbox.checked = !!rule.enabled;
        enabledCheckbox.title = 'Включить/выключить правило';

        const sitesInput = createInput('text', rule.sites || '', '№ сайтов, прим: "2-4", "1, 3", пустой - глобальн.', {flex: '0 0 75%', fontFamily: 'monospace', fontSize: '12px'}, {
            blur: function() { manager.update(idx, findInput.value, replaceInput.value, enabledCheckbox.checked, sitesInput.value); },
            change: function() { sitesInput.blur(); }
        });
        sitesInput.id = type + '-rule-sites-' + idx;

        const findInput = createInput('text', rule.find || '', 'Найти (regex)', {flex: '1 1 50%', fontFamily: 'monospace'}, {
            blur: function() {
                const value = findInput.value;
                const trimmed = value.trim();
                if (trimmed === '') {
                    errorDiv.textContent = 'Паттерн regex пустой';
                    errorDiv.style.display = 'block';
                    findInput.style.borderColor = '#b00020';
                } else {
                    try {
                        new RegExp(value);
                        errorDiv.style.display = 'none';
                        findInput.style.borderColor = '';
                    } catch (e) {
                        errorDiv.textContent = 'Некорректный regex: ' + e.message;
                        errorDiv.style.display = 'block';
                        findInput.style.borderColor = '#b00020';
                    }
                }
                manager.update(idx, findInput.value, replaceInput.value, enabledCheckbox.checked, sitesInput.value);
            },
            change: function() { findInput.blur(); }
        });
        findInput.id = type + '-rule-find-' + idx;

        const replaceInput = createInput('text', rule.replace || '', 'Будет заменён на', {flex: '1 1 50%', fontFamily: 'monospace'}, {
            blur: function() { manager.update(idx, findInput.value, replaceInput.value, enabledCheckbox.checked, sitesInput.value); },
            change: function() { replaceInput.blur(); }
        });
        replaceInput.id = type + '-rule-rep-' + idx;

        const deleteBtn = createButton('Удалить', 'rule-delete-btn', {}, function() {
            if (confirm('Удалить правило?')) {
                manager.delete(idx);
                renderReplacementsUI(type);
            }
        });

        const actions = document.createElement('div');
        actions.className = 'rule-actions';
        actions.appendChild(deleteBtn);

        rowWrap.appendChild(enabledCheckbox);
        rowWrap.appendChild(sitesInput);
        rowWrap.appendChild(findInput);
        rowWrap.appendChild(replaceInput);
        rowWrap.appendChild(actions);

        const errorDiv = document.createElement('div');
        errorDiv.className = 'rule-error';
        errorDiv.style.display = 'none';

        validateRegexInline(findInput, errorDiv);

        return { rowWrap, errorDiv };
    }

    function renderReplacementsUI(type) {
        const manager = type === 'url' ? urlReplacements : resultReplacements;
        const containerId = type + '-replace-rules';
        const container = document.getElementById(containerId);
        if (!container) return;
        container.textContent = '';
        const rules = manager.get();

        if (rules.length === 0) {
            const hint = document.createElement('div');
            hint.style.color = '#666';
            hint.style.fontSize = '13px';
            hint.textContent = 'Правил пока нет — добавьте через форму выше.';
            container.appendChild(hint);
            return;
        }

        rules.forEach((r, idx) => {
            const { rowWrap, errorDiv } = createRuleRow(type, idx, r, manager);
            container.appendChild(rowWrap);
            container.appendChild(errorDiv);
        });
    }

    /* ----------------- UI: main floating window ----------------- */
    function createFloatingWindow() {
        if (document.getElementById('data-scraper-window')) return;

        const container = document.createElement('div');
        container.id = 'data-scraper-window';

        const header = document.createElement('div');
        header.id = 'scraper-header';

        const headerTitle = document.createElement('span');
        headerTitle.textContent = 'Копирование данных';

        const toggleBtn = document.createElement('button');
        toggleBtn.id = 'scraper-toggle';
        toggleBtn.textContent = '−';

        header.appendChild(headerTitle);
        header.appendChild(toggleBtn);

        const content = document.createElement('div');
        content.id = 'scraper-content';

        const urlReplaceBlock = createReplacementSection('url', 'urlReplaceCollapsed', 'Замена в URL (Regex)');
        const resultReplaceBlock = createReplacementSection('result', 'resultReplaceCollapsed', 'Замена в итоговом результате (Regex)');

        // Sites section
        const sitesSection = document.createElement('div');
        sitesSection.id = 'sites-section';
        sitesSection.style.marginBottom = '5px';
        sitesSection.className = 'reorderable-section';

        const sitesSectionHeader = document.createElement('div');
        sitesSectionHeader.className = 'section-header';

        const sitesToggle = document.createElement('button');
        sitesToggle.className = 'section-toggle';
        const sitesCollapsed = GM_getValue('sitesCollapsed', true);
        sitesToggle.textContent = sitesCollapsed ? '+' : '−';

        const sitesTitle = document.createElement('div');
        sitesTitle.className = 'section-title';
        sitesTitle.style.fontWeight = '700';
        sitesTitle.textContent = 'Сайты';

        sitesSectionHeader.appendChild(sitesToggle);
        sitesSectionHeader.appendChild(sitesTitle);

        const sitesContainer = document.createElement('div');
        sitesContainer.id = 'sites-container';
        sitesContainer.style.display = sitesCollapsed ? 'none' : 'block';

        sitesSection.appendChild(sitesSectionHeader);
        sitesSection.appendChild(sitesContainer);

        const toggleSites = function() {
            const isCollapsed = sitesContainer.style.display === 'none';
            sitesContainer.style.display = isCollapsed ? 'block' : 'none';
            sitesToggle.textContent = isCollapsed ? '−' : '+';
            GM_setValue('sitesCollapsed', !isCollapsed);
        };
        sitesToggle.addEventListener('click', function(e) {
            e.stopPropagation();
            toggleSites();
        });
        sitesSectionHeader.addEventListener('click', toggleSites);

        const startBtn = document.createElement('button');
        startBtn.id = 'scraper-start';
        startBtn.textContent = 'Запустить копирование';

        const status = document.createElement('div');
        status.id = 'scraper-status';
        status.style.display = 'none';

        content.appendChild(startBtn);
        content.appendChild(status);

        // Append sections in saved order
        const savedOrder = GM_getValue('sectionOrder', JSON.stringify(['url-replace-block', 'result-replace-block', 'sites-section']));
        const order = JSON.parse(savedOrder);
        const sections = {
            'url-replace-block': urlReplaceBlock,
            'result-replace-block': resultReplaceBlock,
            'sites-section': sitesSection
        };
        order.forEach(id => {
            if (sections[id]) {
                content.appendChild(sections[id]);
            }
        });

        container.appendChild(header);
        container.appendChild(content);

        const style = document.createElement('style');
        style.textContent = '#data-scraper-window { position: fixed; bottom: 20px; right: 20px; width: 480px; height: 500px; background: #ffffff; border: 1px solid #d0d0d0; border-radius: 8px; box-shadow: 0 6px 18px rgba(0,0,0,0.15); z-index: 999999; font-family: Inter, Arial, sans-serif; font-size: 13px; display:flex; flex-direction:column; resize: both; overflow: auto; min-width: 220px; min-height: 50px; } #scraper-header { background:#f5f5f5; padding:10px 12px; border-bottom:1px solid #e6e6e6; border-radius:8px 8px 0 0; display:flex; justify-content:space-between; align-items:center; cursor:move; font-weight:700; } #scraper-toggle { background:none;border:none;font-size:20px;cursor:pointer;padding:0;width:24px;height:24px;line-height:1; } #scraper-content { padding:12px; overflow-y:auto; flex:1; display:flex; flex-direction:column; gap:4px; } #scraper-content.collapsed { display: none; } .secondary-btn { padding:8px 10px; background:#fff; color:#1a73e8; border:1px solid #1a73e8; border-radius:6px; font-size:13px; font-weight:600; cursor:pointer; } .rule-add-btn { background: transparent; color:#1a73e8; border:1px solid #1a73e8; border-radius:6px; font-size:13px; font-weight:600; cursor:pointer; } #scraper-start { padding:10px; background:#1a73e8; color:white; border:none; border-radius:6px; font-size:14px; font-weight:700; cursor:pointer; margin-bottom: 2px; } #scraper-status { padding:3px 8px; background:#f8f9fa; border-radius:6px; min-height:20px; font-size:12px; color:#5f6368; } .site-block { padding:10px 0px; display:flex; flex-direction:column; gap:8px; border: 2px solid transparent; border-radius: 4px; } .site-block.drag-over { border: 2px dashed #1a73e8; background-color: #f0f7ff; } .site-block.dragging { opacity: 0.5; } .site-name-display { font-weight: bold; margin-bottom: 5px; cursor: grab; padding: 4px 6px; border-radius: 4px; user-select: none; background-color: #f2f2f2;} .site-name-display:hover { background-color: #dddddd; } .site-name-display:active { cursor: grabbing; } .site-name-display.dragging { cursor: grabbing; opacity: 0.7; background-color: #e0e0e0; } .sites-divider { border-top: 1px solid #e6e6e6; margin: -5px 0; } .url-rule-row, .result-rule-row { display:flex; gap:8px; align-items:center; } .url-rule-row input[type="text"], .result-rule-row input[type="text"] { padding: 4px 8px; font-family:monospace; font-size:13px; } .url-rule-row input[type="checkbox"], .result-rule-row input[type="checkbox"] { width:18px; height:18px; } .rule-actions { display:flex; gap:6px; } .rule-delete-btn { background: transparent; color:#f44336; border:1px solid #f44336; padding:6px 8px; border-radius:6px; cursor:pointer; } .rule-save-btn { background:#0d873e; color:white; border:none; padding:6px 8px; border-radius:6px; cursor:pointer; } .rule-error { color:#b00020; font-size:12px; margin-top:4px; } .selector-item { display:flex; gap:6px; align-items:center; } .remove-selector-btn { background: transparent; color:#f44336; border:1px solid #f44336; padding:4px 8px; border-radius:6px; cursor:pointer; font-size:12px; white-space:nowrap; } .add-selector-btn { background: transparent; color:#1a73e8; border:1px solid #1a73e8; padding:8px 10px; border-radius:6px; cursor:pointer; font-size:13px; } .site-delete-btn { padding:8px 10px; background:#fff; color:#f44336; border:1px solid #f44336; border-radius:6px; font-size:13px; font-weight:600; cursor:pointer; } .section-header { display: flex; align-items: center; margin-bottom: 8px; cursor: pointer; user-select: none; } .section-toggle { background: none; border: none; font-size: 20px; cursor: pointer; padding: 0; width: 24px; height: 24px; line-height: 1; margin-right: 8px; } .reorderable-section {  } .reorderable-section.drag-over { border: 2px dashed #1a73e8; background-color: #f0f7ff; } .reorderable-section.dragging { opacity: 0.5; } .section-title { font-weight: bold; cursor: grab; padding: 4px 6px; border-radius: 4px; user-select: none; font-size: 14px} .section-title:hover { background-color: #f0f0f0; } .section-title:active { cursor: grabbing; } .section-title.dragging { cursor: grabbing; opacity: 0.7; background-color: #e0e0e0; } .url-row { display: flex; align-items: center; gap: 6px; } .url-row input[type="checkbox"] { width: 18px; height: 18px; flex: 0 0 18px; } .site-header { padding-left: 20px; } .site-right-actions { display: flex; align-items: center; gap: 6px; } .site-delete-btn-small { padding: 4px 6px; font-size: 12px; }';
        document.head.appendChild(style);
        document.body.appendChild(container);

        // Drag and drop for reordering sections
        let draggedSection = null;
        const reorderableSections = [urlReplaceBlock, resultReplaceBlock, sitesSection];
        const dragHandles = reorderableSections.map(sec => sec.querySelector('.section-title'));

        dragHandles.forEach((handle, i) => {
            handle.draggable = true;
            handle.addEventListener('dragstart', (e) => {
                draggedSection = reorderableSections[i];
                e.dataTransfer.effectAllowed = 'move';
                draggedSection.classList.add('dragging');
                handle.classList.add('dragging');
            });
            handle.addEventListener('dragend', (e) => {
                if (draggedSection) {
                    draggedSection.classList.remove('dragging');
                    handle.classList.remove('dragging');
                }
                reorderableSections.forEach(sec => sec.classList.remove('drag-over'));
                draggedSection = null;
            });
        });

        reorderableSections.forEach(section => {
            section.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                if (draggedSection !== section) {
                    section.classList.add('drag-over');
                }
            });
            section.addEventListener('dragleave', (e) => {
                if (!section.contains(e.relatedTarget)) {
                    section.classList.remove('drag-over');
                }
            });
            section.addEventListener('drop', (e) => {
                e.preventDefault();
                reorderableSections.forEach(sec => sec.classList.remove('drag-over'));
                if (draggedSection && draggedSection !== section) {
                    // Remove dragging classes immediately
                    draggedSection.classList.remove('dragging');
                    const draggedHandle = draggedSection.querySelector('.section-title');
                    if (draggedHandle) draggedHandle.classList.remove('dragging');

                    const defaultOrder = ['url-replace-block', 'result-replace-block', 'sites-section'];
                    const savedOrderStr = GM_getValue('sectionOrder', JSON.stringify(defaultOrder));
                    let order;
                    try {
                        order = JSON.parse(savedOrderStr);
                    } catch (e) {
                        order = defaultOrder;
                    }
                    const draggedId = draggedSection.id;
                    const targetId = section.id;
                    const draggedIdx = order.indexOf(draggedId);
                    const targetIdx = order.indexOf(targetId);
                    if (draggedIdx !== -1 && targetIdx !== -1 && draggedIdx !== targetIdx) {
                        order.splice(draggedIdx, 1);
                        order.splice(targetIdx, 0, draggedId);
                        GM_setValue('sectionOrder', JSON.stringify(order));
                        // Re-render sections in new order
                        const contentEl = document.getElementById('scraper-content');
                        const dropSectionMap = {
                            'url-replace-block': document.getElementById('url-replace-block'),
                            'result-replace-block': document.getElementById('result-replace-block'),
                            'sites-section': document.getElementById('sites-section')
                        };
                        // Remove existing sections
                        Object.values(dropSectionMap).forEach(sec => {
                            if (sec && contentEl.contains(sec)) {
                                contentEl.removeChild(sec);
                            }
                        });
                        // Append in new order
                        order.forEach(id => {
                            const sec = dropSectionMap[id];
                            if (sec) {
                                contentEl.appendChild(sec);
                            }
                        });
                    }
                }
                draggedSection = null;
            });
        });

        makeDraggable(container);
        renderReplacementsUI('url');
        renderReplacementsUI('result');
        renderSites();

        const originalHeight = getComputedStyle(container).height;

        toggleBtn.addEventListener('click', function() {
            const isCollapsed = content.classList.contains('collapsed');
            if (isCollapsed) {
                container.style.height = originalHeight;
                container.style.maxHeight = '';
                content.classList.remove('collapsed');
                this.textContent = '−';
            } else {
                container.style.height = '50px';
                container.style.maxHeight = '50px';
                content.classList.add('collapsed');
                this.textContent = '+';
            }
        });
        startBtn.onclick = startScraping;

        if (!window.__scraper_listener_registered) {
            window.__scraper_listener_registered = true;
            try { window.__scraper_listener_id = GM_addValueChangeListener('lastScrapedSite', onLastScrapedSite); } catch (e) { console.warn('GM_addValueChangeListener failed:', e); }
        }

        const prev = GM_getValue('scrapedDataArray');
        if (prev) {
            GM_setValue('scrapedDataArray', JSON.stringify([]));
        }
    }

    /* ----------------- Sites UI with drag-and-drop support for individual sites ----------------- */
    function renderSites() {
        const container = document.getElementById('sites-container');
        if (!container) return;
        container.textContent = '';
        const sites = getSites();

        sites.forEach((site, index) => {
            const siteBlock = createSiteBlock(site, index);
            container.appendChild(siteBlock);

            // Setup drag and drop for site block
            setupSiteDragAndDrop(siteBlock, index);

            if (index < sites.length - 1) {
                const divider = document.createElement('div');
                divider.className = 'sites-divider';
                container.appendChild(divider);
            }
        });

        let addSiteBtn = document.getElementById('add-site-btn');
        if (!addSiteBtn) {
            addSiteBtn = createButton('+ Добавить сайт', 'secondary-btn', {}, addNewSite);
            addSiteBtn.id = 'add-site-btn';
        }
        container.appendChild(addSiteBtn);
    }

    let draggedSiteBlock = null;

    function setupSiteDragAndDrop(siteBlock, index) {
        const handle = siteBlock.querySelector('.section-title');

        handle.draggable = true;

        handle.addEventListener('dragstart', (e) => {
            draggedSiteBlock = siteBlock;
            e.dataTransfer.effectAllowed = 'move';
            siteBlock.classList.add('dragging');
            handle.classList.add('dragging');
        });

        handle.addEventListener('dragend', (e) => {
            siteBlock.classList.remove('dragging');
            handle.classList.remove('dragging');
            const allBlocks = document.querySelectorAll('.site-block');
            allBlocks.forEach(b => b.classList.remove('drag-over'));
            draggedSiteBlock = null;
        });

        siteBlock.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            if (!siteBlock.classList.contains('dragging')) {
                siteBlock.classList.add('drag-over');
            }
        });

        siteBlock.addEventListener('dragleave', (e) => {
            siteBlock.classList.remove('drag-over');
        });

        siteBlock.addEventListener('drop', (e) => {
            e.preventDefault();
            siteBlock.classList.remove('drag-over');

            if (draggedSiteBlock && draggedSiteBlock !== siteBlock) {
                const container = document.getElementById('sites-container');
                const blocks = Array.from(container.querySelectorAll('.site-block'));
                const draggedIndex = blocks.indexOf(draggedSiteBlock);
                const targetIndex = blocks.indexOf(siteBlock);

                if (draggedIndex !== -1 && targetIndex !== -1) {
                    const sites = getSites();
                    const [draggedSite] = sites.splice(draggedIndex, 1);
                    sites.splice(targetIndex, 0, draggedSite);
                    saveSites(sites);
                    renderSites();
                }
            }
        });
    }

    function createSelectorRow(type, index, subIndex, value, enabled, updateFn, removeFn, isRedirect = false) {
        const row = document.createElement('div');
        row.className = 'selector-item';

        let checkbox;
        if (!isRedirect) {
            checkbox = createInput('checkbox', '', '', {flex: '0 0 18px', width: '18px', height: '18px'}, {
                change: (e) => updateFn(index, subIndex, e.target.checked)
            });
            checkbox.checked = !!enabled;
            checkbox.title = 'Включить/выключить селектор';
            row.appendChild(checkbox);
        }

        const input = createInput('text', value || '', isRedirect ? 'Поиск и открытие гиперссылки по селектору' : 'CSS селектор', {flex: '1'}, {
            change: (e) => updateFn(index, subIndex, e.target.value)
        });
        row.appendChild(input);

        const remBtn = createButton('✕', 'remove-selector-btn', {fontSize: '12px'}, () => removeFn(index, subIndex));
        row.appendChild(remBtn);

        return row;
    }

    function createSiteBlock(site, index) {
        const block = document.createElement('div');
        block.className = 'site-block';
        block.dataset.index = index;

        // Site header
        const siteHeader = document.createElement('div');
        siteHeader.className = 'section-header site-header';

        siteHeader.style.display = 'flex';
        siteHeader.style.alignItems = 'center';
        siteHeader.style.justifyContent = 'space-between';

        const toggle = document.createElement('button');
        toggle.className = 'section-toggle';
        const collapsed = site.collapsed !== undefined ? site.collapsed : true;
        toggle.textContent = collapsed ? '+' : '−';

        const siteTitle = document.createElement('div');
        siteTitle.className = 'section-title';
        siteTitle.textContent = (index + 1) + '. ' + (site.name || 'Untitled');

        const leftPart = document.createElement('div');
        leftPart.style.display = 'flex';
        leftPart.style.alignItems = 'center';
        leftPart.appendChild(toggle);
        leftPart.appendChild(siteTitle);

        const enabledCheckbox = createInput('checkbox', '', '', {width: '18px', height: '18px'}, {
            change: (e) => updateSiteEnabled(index, e.target.checked)
        });
        enabledCheckbox.checked = !!site.enabled;
        enabledCheckbox.title = 'Включить/выключить сайт';

        const deleteBtn = createButton('Удалить сайт', 'site-delete-btn site-delete-btn-small', {}, () => deleteSite(index));

        const rightActions = document.createElement('div');
        rightActions.className = 'site-right-actions';
        rightActions.appendChild(enabledCheckbox);
        rightActions.appendChild(deleteBtn);

        siteHeader.appendChild(leftPart);
        siteHeader.appendChild(rightActions);

        block.appendChild(siteHeader);

        // Content
        const content = document.createElement('div');
        content.style.display = collapsed ? 'none' : 'block';

        // URL input
        const urlInput = createInput('text', site.url || '', 'URL', {width: '100%', marginBottom: '8px', padding: '4px 8px', fontFamily: 'monospace', fontSize: '13px'}, {
            change: (e) => updateSiteUrl(index, e.target.value)
        });
        content.appendChild(urlInput);

        // Redirects
        const redirectLabel = document.createElement('div');
        redirectLabel.textContent = 'Переадресация:';
        redirectLabel.style.fontWeight = '600';

        const addRedBtn = createButton('+ Добавить уровень переадресации', 'add-selector-btn', {padding: '4px 8px', fontSize: '12px'}, () => addRedirect(index));

        const redirectHeader = document.createElement('div');
        redirectHeader.style.display = 'flex';
        redirectHeader.style.alignItems = 'center';
        redirectHeader.style.justifyContent = 'space-between';
        redirectHeader.style.marginBottom = '4px';
        redirectHeader.appendChild(redirectLabel);
        redirectHeader.appendChild(addRedBtn);

        const redirectsDiv = document.createElement('div');
        (site.redirectSelectors || []).forEach((sel, rIndex) => {
            const redRow = createSelectorRow('redirect', index, rIndex, sel, null, updateRedirect, removeRedirect, true);
            redirectsDiv.appendChild(redRow);
        });

        content.appendChild(redirectHeader);
        content.appendChild(redirectsDiv);

        // Selectors
        const selectorsLabel = document.createElement('div');
        selectorsLabel.textContent = 'Селекторы:';
        selectorsLabel.style.fontWeight = '600';

        const addSelBtn = createButton('+ Добавить селектор', 'add-selector-btn', {padding: '4px 8px', fontSize: '12px'}, () => addSelector(index));

        const selectorHeader = document.createElement('div');
        selectorHeader.style.display = 'flex';
        selectorHeader.style.alignItems = 'center';
        selectorHeader.style.justifyContent = 'space-between';
        selectorHeader.style.marginBottom = '4px';
        selectorHeader.appendChild(selectorsLabel);
        selectorHeader.appendChild(addSelBtn);

        const selectorsDiv = document.createElement('div');
        (site.selectors || []).forEach((selObj, sIndex) => {
            const selRow = createSelectorRow('selector', index, sIndex, selObj.selector, selObj.enabled, updateSelector, removeSelector);
            selectorsDiv.appendChild(selRow);
        });

        content.appendChild(selectorHeader);
        content.appendChild(selectorsDiv);

        block.appendChild(content);

        // Toggle functionality
        const toggleFn = function(e) {
            e.stopPropagation();
            const isCollapsed = content.style.display === 'none';
            content.style.display = isCollapsed ? 'block' : 'none';
            toggle.textContent = isCollapsed ? '−' : '+';
            const sites = getSites();
            sites[index].collapsed = !isCollapsed;
            saveSites(sites);
        };
        toggle.addEventListener('click', toggleFn);
        siteHeader.addEventListener('click', function(e) {
            if (!e.target.closest('.site-right-actions')) {
                toggleFn(e);
            }
        });

        return block;
    }

    function addNewSite() {
        const sites = getSites();
        const newName = generateUniqueSiteName(sites, sites.length, 'https://example.com');
        sites.push({ name: newName, url: 'https://', enabled: true, collapsed: false, redirectSelectors: [], selectors: [{enabled: true, selector: ''}] });
        saveSites(sites);
        renderSites();
    }

    function deleteSite(index) {
        const sites = getSites();
        if (sites.length === 1) { alert('Нельзя удалить последний сайт!'); return; }
        if (confirm('Удалить этот сайт?')) { sites.splice(index, 1); saveSites(sites); renderSites(); }
    }

    function updateSiteUrl(index, value) {
        const sites = getSites();
        sites[index].url = value;
        sites[index].name = generateUniqueSiteName(sites, index, value);
        saveSites(sites);
        renderSites();
    }

    function updateSiteEnabled(index, enabled) {
        const sites = getSites();
        sites[index].enabled = !!enabled;
        saveSites(sites);
    }

    function addRedirect(index) {
        const sites = getSites();
        sites[index].redirectSelectors.push('');
        saveSites(sites);
        renderSites();
    }

    function removeRedirect(index, rIndex) {
        const sites = getSites();
        sites[index].redirectSelectors.splice(rIndex, 1);
        saveSites(sites);
        renderSites();
    }

    function updateRedirect(index, rIndex, value) {
        const sites = getSites();
        sites[index].redirectSelectors[rIndex] = value;
        saveSites(sites);
    }

    function updateSelector(index, sIndex, value) {
        const sites = getSites();
        if (typeof value === 'boolean') {
            sites[index].selectors[sIndex].enabled = !!value;
        } else {
            sites[index].selectors[sIndex].selector = value;
        }
        saveSites(sites);
    }

    function addSelector(siteIndex) {
        const sites = getSites();
        sites[siteIndex].selectors.push({enabled: true, selector: ''});
        saveSites(sites);
        renderSites();
    }

    function removeSelector(siteIndex, selectorIndex) {
        const sites = getSites();
        if (sites[siteIndex].selectors.length === 1) { alert('Нельзя удалить последний селектор!'); return; }
        sites[siteIndex].selectors.splice(selectorIndex, 1);
        saveSites(sites);
        renderSites();
    }

    /* ----------------- Draggable window ----------------- */
    function makeDraggable(elem) {
        let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        const header = document.getElementById('scraper-header');
        if (!header) return;
        header.onmousedown = dragMouseDown;
        function dragMouseDown(e) { e.preventDefault(); pos3 = e.clientX; pos4 = e.clientY; document.onmouseup = closeDragElement; document.onmousemove = elementDrag; }
        function elementDrag(e) {
            e.preventDefault();
            pos1 = pos3 - e.clientX;
            pos2 = pos4 - e.clientY;
            pos3 = e.clientX;
            pos4 = e.clientY;
            let newLeft = elem.offsetLeft - pos1;
            let newTop = elem.offsetTop - pos2;
            const maxLeft = window.innerWidth - elem.offsetWidth;
            const maxTop = window.innerHeight - elem.offsetHeight;
            newLeft = Math.max(0, Math.min(maxLeft, newLeft));
            newTop = Math.max(0, Math.min(maxTop, newTop));
            elem.style.top = newTop + "px";
            elem.style.left = newLeft + "px";
            elem.style.bottom = "auto";
            elem.style.right = "auto";
        }
        function closeDragElement() { document.onmouseup = null; document.onmousemove = null; }
    }

    /* ----------------- Scraping logic & listeners ----------------- */
    function onLastScrapedSite(name, oldValue, newValue, remote) {
        try {
            const parsed = newValue ? JSON.parse(newValue) : null;
            if (!parsed || typeof parsed.index === 'undefined') return;
            const enabledStr = GM_getValue('enabledIndices');
            const enabledIndices = enabledStr ? JSON.parse(enabledStr) : [];
            const totalEnabled = enabledIndices.length;
            const currentStep = enabledIndices.indexOf(parsed.index);
            const status = document.getElementById && document.getElementById('scraper-status');
            if (status) {
                status.textContent = 'Скопировано: ' + parsed.siteName + ' (' + (currentStep + 1) + '/' + totalEnabled + ')';
                status.style.color = '#0d652d';
                updateStatusVisibility(status);
            }
            const nextStep = currentStep + 1;
            if (nextStep < totalEnabled) {
                const nextOrigIndex = enabledIndices[nextStep];
                setTimeout(() => openSiteByOriginalIndex(nextOrigIndex), 200);
            } else {
                GM_setValue('scrapeActive', false);
                GM_setValue('enabledIndices', '');
                if (status) {
                    status.textContent = 'Готово: обработано ' + totalEnabled + ' сайт(ов).';
                    status.style.color = '#0d652d';
                    updateStatusVisibility(status);
                }
            }
        } catch (e) { console.error('Ошибка в onLastScrapedSite:', e); }
    }

    function startScraping() {
        const sites = getSites();
        const enabledIndices = [];
        for (let i = 0; i < sites.length; i++) {
            if (sites[i].enabled) enabledIndices.push(i);
        }
        const totalEnabled = enabledIndices.length;
        const status = document.getElementById && document.getElementById('scraper-status');

        if (totalEnabled === 0) {
            if (status) {
                status.textContent = 'Нет включенных сайтов для обработки!';
                status.style.color = '#d32f2f';
                updateStatusVisibility(status);
            } else {
                alert('Нет включенных сайтов для обработки!');
            }
            return;
        }

        if (GM_getValue('scrapeActive')) {
            if (status) {
                status.textContent = 'Скрейпинг уже выполняется — дождитесь завершения.';
                status.style.color = '#d39e00';
                updateStatusVisibility(status);
            }
            return;
        }

        // Validation
        for (let origI of enabledIndices) {
            const site = sites[origI];
            if (!site.url || site.url === 'https://') { alert('Сайт "' + site.name + '" имеет невалидный URL!'); return; }
            const enabledSelectors = (site.selectors || []).filter(s => s && s.enabled);
            if (!enabledSelectors.length || enabledSelectors.every(s => !s.selector.trim())) { alert('Сайт "' + site.name + '" не имеет включённых селекторов!'); return; }
            if (site.redirectSelectors && site.redirectSelectors.some(s => !s.trim())) { alert('Пустая переадресация для сайта "' + site.name + '"!'); return; }
        }

        if (status) {
            status.textContent = 'Открываем 1/' + totalEnabled + ' сайт(ов)...';
            status.style.color = '#1a73e8';
            updateStatusVisibility(status);
        }

        GM_setValue('scrapeActive', true);
        GM_setValue('enabledIndices', JSON.stringify(enabledIndices));
        GM_setValue('scrapeQueue', JSON.stringify(sites));
        GM_setValue('scrapedClipboard', '');
        GM_setValue('scrapedDataArray', JSON.stringify([]));
        try { GM_setClipboard(''); } catch (e) {}

        const firstOrigIndex = enabledIndices[0];
        openSiteByOriginalIndex(firstOrigIndex);

        const guardKey = 'scrapeGuardTimer';
        try { clearTimeout(window[guardKey]); } catch (e) {}
        window[guardKey] = setTimeout(() => {
            if (GM_getValue('scrapeActive')) {
                GM_setValue('scrapeActive', false);
                GM_setValue('enabledIndices', '');
                if (status) {
                    status.textContent = 'Прервано: время ожидания истекло.';
                    status.style.color = '#d32f2f';
                    updateStatusVisibility(status);
                }
            }
        }, 2 * 60 * 1000);
    }

    function openSiteByOriginalIndex(originalIndex) {
        const sites = getSites();
        if (originalIndex < 0 || originalIndex >= sites.length) return;
        const site = sites[originalIndex];
        const siteNum = originalIndex + 1;

        const openUrl = urlReplacements.apply(site.url, siteNum);

        GM_setValue(`redirectLevel_${originalIndex}`, 0);
        GM_setValue('currentOpeningIndex', originalIndex);
        GM_openInTab(openUrl, { active: true, insert: true });
    }

    if (!isGoogleSheets) {
        if (GM_getValue('scrapeActive')) {
            scrapePage();
        }
    }

    function scrapePage() {
        async function doScrape() {
            const configStr = GM_getValue('scrapeQueue');
            if (!configStr) return;
            let sites;
            try { sites = JSON.parse(configStr); } catch (e) { sites = []; }
            const currentUrl = window.location.href;

            let siteIndex = -1;
            for (let i = 0; i < sites.length; i++) {
                try {
                    const candidateHostname = new URL(sites[i].url).hostname;
                    if (currentUrl.includes(candidateHostname) || currentUrl.startsWith(sites[i].url)) { siteIndex = i; break; }
                } catch (e) {}
            }

            if (siteIndex === -1) {
                try {
                    const fallback = GM_getValue('currentOpeningIndex');
                    if (typeof fallback === 'number' && fallback >= 0 && fallback < sites.length) siteIndex = Number(fallback);
                } catch (e) {}
            }

            if (siteIndex === -1) {
                console.log('Конфигурация для этого сайта не найдена (текущий URL =', currentUrl, ')');
                return;
            }

            const siteConfig = sites[siteIndex];
            const siteNum = siteIndex + 1;

            const redirectSelectors = Array.isArray(siteConfig.redirectSelectors) ? siteConfig.redirectSelectors : [];
            const levelKey = `redirectLevel_${siteIndex}`;
            let currentLevel = GM_getValue(levelKey, 0);

            if (currentLevel < redirectSelectors.length) {
                const selector = redirectSelectors[currentLevel];
                if (selector.trim()) {
                    try {
                        const linkElem = document.querySelector(selector);
                        if (linkElem && linkElem.tagName.toLowerCase() === 'a' && linkElem.href) {
                            GM_setValue(levelKey, currentLevel + 1);
                            GM_openInTab(linkElem.href, { active: true, insert: true });
                            setTimeout(() => window.close(), 100);
                            return; // don't scrape
                        }
                    } catch (e) {
                        console.error('Error in redirect:', e);
                    }
                }
            }

            // If no redirect or all done, scrape
            const rawSelectors = Array.isArray(siteConfig.selectors) ? siteConfig.selectors : [];
            const selectors = rawSelectors.filter(s => s && s.enabled).map(s => s.selector);
            const timeoutMs = 20000;
            const pollInterval = 250;
            let values = await waitForSelectorsAndCollect(selectors, timeoutMs, pollInterval);

            // Apply replacements to each value individually
            const cleanedValues = values.map(value => resultReplacements.apply(value || '', siteNum));

            const data = { siteName: siteConfig.name, url: currentUrl, values: cleanedValues, timestamp: new Date().toISOString(), index: siteIndex };

            try {
                let dataArray = [];
                const allData = GM_getValue('scrapedDataArray');
                if (allData) { dataArray = JSON.parse(allData); if (!Array.isArray(dataArray)) dataArray = []; }
                dataArray.push(data);
                GM_setValue('scrapedDataArray', JSON.stringify(dataArray));
            } catch (e) { console.warn('Ошибка записи данных:', e); }

            try {
                const clipboardTextForSite = cleanedValues.join('\t');
                const prev = GM_getValue('scrapedClipboard') || '';
                const newCombined = prev ? prev + '\t' + clipboardTextForSite : clipboardTextForSite;
                GM_setValue('scrapedClipboard', newCombined);
                try { GM_setClipboard(newCombined); } catch (e) {}
            } catch (e) {
                console.warn('Ошибка с клипбордом:', e);
            }

            try {
                GM_setValue('lastScrapedSite', JSON.stringify({ index: siteIndex, siteName: siteConfig.name, url: currentUrl, timestamp: new Date().toISOString(), values: cleanedValues }));
            } catch (e) { console.warn('GM_setValue lastScrapedSite failed:', e); }

            setTimeout(() => { try { window.close(); } catch (e) { console.log('window.close failed'); } }, 100);
        }

        if (document.readyState === 'loading') {
            window.addEventListener('load', doScrape);
        } else {
            doScrape();
        }
    }

    function waitForSelectorsAndCollect(selectors, timeoutMs = 15000, pollInterval = 300) {
        return new Promise((resolve) => {
            const start = Date.now();
            const results = selectors.map(() => 'Не найдено');
            function poll() {
                let anyFound = false;
                for (let i = 0; i < selectors.length; i++) {
                    const sel = selectors[i];
                    if (!sel || !sel.trim()) { results[i] = 'Пустой селектор'; continue; }
                    try {
                        const el = document.querySelector(sel);
                        if (el) {
                            const text = (el.textContent || el.innerText || '').trim();
                            results[i] = text || 'Пустой текст';
                            anyFound = true;
                        }
                    } catch (e) { results[i] = 'Ошибка селектора'; }
                }
                if (anyFound || (Date.now() - start) >= timeoutMs) { resolve(results); } else { setTimeout(poll, pollInterval); }
            }
            poll();
        });
    }

    /* ----------------- Init ----------------- */
    if (isGoogleSheets) {
        createFloatingWindow();
    }

    function updateStatusVisibility(status) {
        if (status.textContent.trim() === '') {
            status.style.display = 'none';
        } else {
            status.style.display = '';
        }
    }

})();
