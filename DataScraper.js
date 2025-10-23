// ==UserScript==
// @name         Google Sheets Data Scraper (regex rules editable inline)
// @namespace    http://tampermonkey.net/
// @version      5.1
// @description  Последовательное копирование данных — URL regex-замены редактируемые прямо в формах (инлайн), авто-сохранение и мгновенное применение + замена итоговых результатов
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

    /* ----------------- URL replacements storage + helpers ----------------- */
    function getUrlReplacements() {
        const raw = GM_getValue('urlReplacements');
        try { return raw ? JSON.parse(raw) : []; } catch (e) { GM_setValue('urlReplacements', JSON.stringify([])); return []; }
    }
    function saveUrlReplacements(rules) {
        GM_setValue('urlReplacements', JSON.stringify(rules));
    }
    function addUrlReplacement(findPattern = '', replaceValue = '') {
        const rules = getUrlReplacements();
        rules.push({ find: findPattern, replace: replaceValue, enabled: true });
        saveUrlReplacements(rules);
        renderUrlReplacementsUI();
        return rules.length - 1;
    }
    function updateUrlReplacement(index, findPattern, replaceValue, enabled) {
        const rules = getUrlReplacements();
        if (index < 0 || index >= rules.length) return;
        rules[index].find = findPattern;
        rules[index].replace = replaceValue;
        rules[index].enabled = !!enabled;
        saveUrlReplacements(rules);
    }
    function deleteUrlReplacement(index) {
        const rules = getUrlReplacements();
        if (index < 0 || index >= rules.length) return;
        rules.splice(index, 1);
        saveUrlReplacements(rules);
        renderUrlReplacementsUI();
    }
    function applyUrlReplacements(originalUrl) {
        let url = originalUrl;
        const rules = getUrlReplacements();
        for (let i = 0; i < rules.length; i++) {
            const r = rules[i];
            if (!r || !r.enabled || typeof r.find !== 'string' || r.find.trim() === '') continue;
            try {
                const re = new RegExp(r.find, 'g');
                url = url.replace(re, r.replace == null ? '' : String(r.replace));
            } catch (e) {
                console.warn(`URL replacement rule #${i} invalid RegExp "${r.find}" — skipped`, e);
            }
        }
        return url;
    }

    /* ----------------- Result replacements storage + helpers ----------------- */
    function getResultReplacements() {
        const raw = GM_getValue('resultReplacements');
        try { return raw ? JSON.parse(raw) : []; } catch (e) { GM_setValue('resultReplacements', JSON.stringify([])); return []; }
    }
    function saveResultReplacements(rules) {
        GM_setValue('resultReplacements', JSON.stringify(rules));
    }
    function addResultReplacement(findPattern = '', replaceValue = '') {
        const rules = getResultReplacements();
        rules.push({ find: findPattern, replace: replaceValue, enabled: true });
        saveResultReplacements(rules);
        renderResultReplacementsUI();
        return rules.length - 1;
    }
    function updateResultReplacement(index, findPattern, replaceValue, enabled) {
        const rules = getResultReplacements();
        if (index < 0 || index >= rules.length) return;
        rules[index].find = findPattern;
        rules[index].replace = replaceValue;
        rules[index].enabled = !!enabled;
        saveResultReplacements(rules);
    }
    function deleteResultReplacement(index) {
        const rules = getResultReplacements();
        if (index < 0 || index >= rules.length) return;
        rules.splice(index, 1);
        saveResultReplacements(rules);
        renderResultReplacementsUI();
    }
    function applyResultReplacements(originalText) {
        let text = originalText;
        const rules = getResultReplacements();
        for (let i = 0; i < rules.length; i++) {
            const r = rules[i];
            if (!r || !r.enabled || typeof r.find !== 'string' || r.find.trim() === '') continue;
            try {
                const re = new RegExp(r.find, 'g');
                text = text.replace(re, r.replace == null ? '' : String(r.replace));
            } catch (e) {
                console.warn(`Result replacement rule #${i} invalid RegExp "${r.find}" — skipped`, e);
            }
        }
        return text;
    }

    /* ----------------- Sites storage (unchanged) ----------------- */
    function getDefaultSites() {
        return [
            {
                name: '2ip.io',
                url: 'https://2ip.io',
                selectors: [
                    'body > div.page-wrapper > div.main-content > div.container > div > div.content > section.ip-block > div > div.ip-info_right > div > div > div:nth-child(2)',
                    'body > div.page-wrapper > div.main-content > div.container > div > div.content > section.ip-block > div > div.ip-info_right > div > div > div:nth-child(3)'
                ]
            }
        ];
    }
    function getSites() {
        const saved = GM_getValue('sites');
        try { return saved ? JSON.parse(saved) : getDefaultSites(); } catch (e) { GM_setValue('sites', JSON.stringify(getDefaultSites())); return getDefaultSites(); }
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
            return count > 0 ? `${hostname} (${count + 1})` : hostname;
        } catch (e) {
            return 'Untitled';
        }
    }

    /* ----------------- UI: main floating window with inline-editable regex rules ----------------- */
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

        // URL replacements block (top) — editable inline
        const urlReplaceBlock = document.createElement('div');
        urlReplaceBlock.id = 'url-replace-block';
        urlReplaceBlock.style.marginBottom = '10px';

        const replaceTitle = document.createElement('div');
        replaceTitle.style.fontWeight = '700';
        replaceTitle.style.marginBottom = '8px';
        replaceTitle.textContent = 'Глобальная замена в URL (Regex)';

        // Add-new inputs (for quick adding)
        const quickRow = document.createElement('div');
        quickRow.className = 'url-rule-row';
        quickRow.style.flexWrap = 'wrap';
        quickRow.style.marginBottom = '14px';

        const quickFind = document.createElement('input');
        quickFind.type = 'text';
        quickFind.placeholder = 'Найти (regex)';
        quickFind.id = 'urlreplace-find-quick';
        quickFind.style.flex = '1 1 50%';
        quickFind.style.padding = '4px 8px';
        quickFind.style.fontFamily = 'monospace';
        quickFind.style.fontSize = '13px';

        const quickReplace = document.createElement('input');
        quickReplace.type = 'text';
        quickReplace.placeholder = 'Заменить на';
        quickReplace.id = 'urlreplace-to-quick';
        quickReplace.style.flex = '1 1 50%';
        quickReplace.style.padding = '4px 8px';
        quickReplace.style.fontFamily = 'monospace';
        quickReplace.style.fontSize = '13px';

        const actions = document.createElement('div');
        actions.className = 'rule-actions';

        const quickAddBtn = document.createElement('button');
        quickAddBtn.textContent = 'Добавить';
        quickAddBtn.className = 'rule-add-btn';
        quickAddBtn.style.padding = '6px 8px';
        quickAddBtn.onclick = function() {
            const find = quickFind.value;
            const replace = quickReplace.value;
            if (!find.trim()) { alert('Паттерн regex обязателен'); quickFind.focus(); return; }
            try { new RegExp(find); } catch (e) { alert('Некорректный regex: ' + e.message); quickFind.focus(); return; }
            const newIndex = addUrlReplacement(find, replace);
            setTimeout(() => {
                const inlineFind = document.querySelector(`#url-rule-find-${newIndex}`);
                if (inlineFind) inlineFind.focus();
            }, 50);
            quickFind.value = '';
            quickReplace.value = '';
        };

        actions.appendChild(quickAddBtn);

        quickRow.appendChild(quickFind);
        quickRow.appendChild(quickReplace);
        quickRow.appendChild(actions);

        const rulesList = document.createElement('div');
        rulesList.id = 'url-replace-rules';
        rulesList.style.display = 'flex';
        rulesList.style.flexDirection = 'column';
        rulesList.style.gap = '8px';

        urlReplaceBlock.appendChild(replaceTitle);
        urlReplaceBlock.appendChild(quickRow);
        urlReplaceBlock.appendChild(rulesList);

        // Result replacements block — editable inline
        const resultReplaceBlock = document.createElement('div');
        resultReplaceBlock.id = 'result-replace-block';
        resultReplaceBlock.style.marginBottom = '10px';

        const resultReplaceTitle = document.createElement('div');
        resultReplaceTitle.style.fontWeight = '700';
        resultReplaceTitle.style.marginBottom = '8px';
        resultReplaceTitle.textContent = 'Замена в итоговом результате (Regex)';

        // Add-new inputs for result replacement
        const quickResultRow = document.createElement('div');
        quickResultRow.className = 'url-rule-row';
        quickResultRow.style.flexWrap = 'wrap';
        quickResultRow.style.marginBottom = '14px';

        const quickResultFind = document.createElement('input');
        quickResultFind.type = 'text';
        quickResultFind.placeholder = 'Найти (regex)';
        quickResultFind.id = 'resultreplace-find-quick';
        quickResultFind.style.flex = '1 1 50%';
        quickResultFind.style.padding = '4px 8px';
        quickResultFind.style.fontFamily = 'monospace';
        quickResultFind.style.fontSize = '13px';

        const quickResultReplace = document.createElement('input');
        quickResultReplace.type = 'text';
        quickResultReplace.placeholder = 'Заменить на';
        quickResultReplace.id = 'resultreplace-to-quick';
        quickResultReplace.style.flex = '1 1 50%';
        quickResultReplace.style.padding = '4px 8px';
        quickResultReplace.style.fontFamily = 'monospace';
        quickResultReplace.style.fontSize = '13px';

        const resultActions = document.createElement('div');
        resultActions.className = 'rule-actions';

        const quickResultAddBtn = document.createElement('button');
        quickResultAddBtn.textContent = 'Добавить';
        quickResultAddBtn.className = 'rule-add-btn';
        quickResultAddBtn.style.padding = '6px 8px';
        quickResultAddBtn.onclick = function() {
            const find = quickResultFind.value;
            const replace = quickResultReplace.value;
            if (!find.trim()) { alert('Паттерн regex обязателен'); quickResultFind.focus(); return; }
            try { new RegExp(find); } catch (e) { alert('Некорректный regex: ' + e.message); quickResultFind.focus(); return; }
            const newIndex = addResultReplacement(find, replace);
            setTimeout(() => {
                const inlineFind = document.querySelector(`#result-rule-find-${newIndex}`);
                if (inlineFind) inlineFind.focus();
            }, 50);
            quickResultFind.value = '';
            quickResultReplace.value = '';
        };

        resultActions.appendChild(quickResultAddBtn);

        quickResultRow.appendChild(quickResultFind);
        quickResultRow.appendChild(quickResultReplace);
        quickResultRow.appendChild(resultActions);

        const resultRulesList = document.createElement('div');
        resultRulesList.id = 'result-replace-rules';
        resultRulesList.style.display = 'flex';
        resultRulesList.style.flexDirection = 'column';
        resultRulesList.style.gap = '8px';

        resultReplaceBlock.appendChild(resultReplaceTitle);
        resultReplaceBlock.appendChild(quickResultRow);
        resultReplaceBlock.appendChild(resultRulesList);

        // rest of UI
        const sitesContainer = document.createElement('div');
        sitesContainer.id = 'sites-container';

        const addSiteBtn = document.createElement('button');
        addSiteBtn.id = 'add-site-btn';
        addSiteBtn.textContent = '+ Добавить сайт';
        addSiteBtn.className = 'secondary-btn';

        const startBtn = document.createElement('button');
        startBtn.id = 'scraper-start';
        startBtn.textContent = 'Запустить копирование';

        const status = document.createElement('div');
        status.id = 'scraper-status';
        status.style.display = 'none';

        content.appendChild(startBtn);
        content.appendChild(status);
        content.appendChild(urlReplaceBlock);
        content.appendChild(resultReplaceBlock);
        content.appendChild(sitesContainer);
        content.appendChild(addSiteBtn);

        container.appendChild(header);
        container.appendChild(content);

        const style = document.createElement('style');
        style.textContent = `
            #data-scraper-window { position: fixed; bottom: 20px; right: 20px; width: 480px; height: 500px; background: #ffffff; border: 1px solid #d0d0d0; border-radius: 8px; box-shadow: 0 6px 18px rgba(0,0,0,0.15); z-index: 999999; font-family: Inter, Arial, sans-serif; font-size: 13px; display:flex; flex-direction:column; resize: both; overflow: auto; min-width: 220px; min-height: 50px; }
            #scraper-header { background:#f5f5f5; padding:10px 12px; border-bottom:1px solid #e6e6e6; border-radius:8px 8px 0 0; display:flex; justify-content:space-between; align-items:center; cursor:move; font-weight:700; }
            #scraper-toggle { background:none;border:none;font-size:20px;cursor:pointer;padding:0;width:24px;height:24px;line-height:1; }
            #scraper-content { padding:12px; overflow-y:auto; flex:1; display:flex; flex-direction:column; gap:6px; }
            #scraper-content.collapsed { display: none; }
            .secondary-btn { padding:8px 10px; background:#fff; color:#1a73e8; border:1px solid #1a73e8; border-radius:6px; font-size:13px; font-weight:600; cursor:pointer; }
            .rule-add-btn { background:#1a73e8; color:white; border:none; border-radius:6px; font-size:13px; font-weight:600; cursor:pointer; }
            #scraper-start { padding:10px; background:#1a73e8; color:white; border:none; border-radius:6px; font-size:14px; font-weight:700; cursor:pointer; }
            #scraper-status { padding:8px; background:#f8f9fa; border-radius:6px; min-height:20px; font-size:12px; color:#5f6368; }
            .site-block { padding:10px 0px; display:flex; flex-direction:column; gap:8px; }
            .sites-divider { border-top: 1px solid #e6e6e6; margin: 5px 0; }
            .site-name-display { font-weight: bold; margin-bottom: 5px; }
            .url-rule-row { display:flex; gap:8px; align-items:center; }
            .url-rule-row input[type="text"] { padding: 4px 8px; font-family:monospace; font-size:13px; }
            .url-rule-row input[type="checkbox"] { width:18px; height:18px; }
            .rule-actions { display:flex; gap:6px; }
            .rule-delete-btn { background:#f44336; color:white; border:none; padding:6px 8px; border-radius:6px; cursor:pointer; }
            .rule-save-btn { background:#0d873e; color:white; border:none; padding:6px 8px; border-radius:6px; cursor:pointer; }
            .rule-error { color:#b00020; font-size:12px; margin-top:4px; }
            .selector-item { display:flex; gap:6px; }
            .remove-selector-btn { background:#f44336; color:white; border:none; padding:4px 8px; border-radius:6px; cursor:pointer; font-size:12px; white-space:nowrap; }
            .add-selector-btn { background:#0d873e; color:white; border:none; padding:8px 10px; border-radius:6px; cursor:pointer; font-size:13px; }
        `;
        document.head.appendChild(style);
        document.body.appendChild(container);

        makeDraggable(container);
        renderUrlReplacementsUI();
        renderResultReplacementsUI();
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
        addSiteBtn.addEventListener('click', addNewSite);
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

    /* ----------------- Render URL replacements as editable forms ----------------- */
    function renderUrlReplacementsUI() {
        const container = document.getElementById('url-replace-rules');
        if (!container) return;
        container.textContent = '';
        const rules = getUrlReplacements();

        if (rules.length === 0) {
            const hint = document.createElement('div');
            hint.style.color = '#666';
            hint.style.fontSize = '13px';
            hint.textContent = 'Правил пока нет — добавьте через форму выше.';
            container.appendChild(hint);
            return;
        }

        rules.forEach((r, idx) => {
            const rowWrap = document.createElement('div');
            rowWrap.className = 'url-rule-row';
            rowWrap.style.flexWrap = 'wrap';

            const enabledCheckbox = document.createElement('input');
            enabledCheckbox.type = 'checkbox';
            enabledCheckbox.checked = !!r.enabled;
            enabledCheckbox.title = 'Включить/выключить правило';
            enabledCheckbox.style.flex = '0 0 18px';
            enabledCheckbox.addEventListener('change', () => {
                updateUrlReplacement(idx, findInput.value, replaceInput.value, enabledCheckbox.checked);
                validateRegexInline(findInput, errorDiv);
            });

            const findInput = document.createElement('input');
            findInput.type = 'text';
            findInput.value = r.find || '';
            findInput.id = `url-rule-find-${idx}`;
            findInput.placeholder = 'Найти (regex) 1';
            findInput.style.flex = '1 1 50%';
            findInput.style.fontFamily = 'monospace';

            const replaceInput = document.createElement('input');
            replaceInput.type = 'text';
            replaceInput.value = r.replace || '';
            replaceInput.id = `url-rule-rep-${idx}`;
            replaceInput.placeholder = 'Будет заменён на';
            replaceInput.style.flex = '1 1 50%';
            replaceInput.style.fontFamily = 'monospace';

            const actions = document.createElement('div');
            actions.className = 'rule-actions';

            const deleteBtn = document.createElement('button');
            deleteBtn.textContent = 'Удалить';
            deleteBtn.className = 'rule-delete-btn';
            deleteBtn.onclick = () => {
                if (confirm('Удалить правило?')) deleteUrlReplacement(idx);
            };

            actions.appendChild(deleteBtn);

            rowWrap.appendChild(enabledCheckbox);
            rowWrap.appendChild(findInput);
            rowWrap.appendChild(replaceInput);
            rowWrap.appendChild(actions);

            const errorDiv = document.createElement('div');
            errorDiv.className = 'rule-error';
            errorDiv.style.display = 'none';

            findInput.addEventListener('blur', () => {
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
                updateUrlReplacement(idx, findInput.value, replaceInput.value, enabledCheckbox.checked);
            });
            findInput.addEventListener('change', () => { findInput.blur(); });

            replaceInput.addEventListener('blur', () => {
                updateUrlReplacement(idx, findInput.value, replaceInput.value, enabledCheckbox.checked);
            });
            replaceInput.addEventListener('change', () => { replaceInput.blur(); });

            validateRegexInline(findInput, errorDiv);

            container.appendChild(rowWrap);
            container.appendChild(errorDiv);
        });
    }

    /* ----------------- Render Result replacements as editable forms ----------------- */
    function renderResultReplacementsUI() {
        const container = document.getElementById('result-replace-rules');
        if (!container) return;
        container.textContent = '';
        const rules = getResultReplacements();

        if (rules.length === 0) {
            const hint = document.createElement('div');
            hint.style.color = '#666';
            hint.style.fontSize = '13px';
            hint.textContent = 'Правил пока нет — добавьте через форму выше.';
            container.appendChild(hint);
            return;
        }

        rules.forEach((r, idx) => {
            const rowWrap = document.createElement('div');
            rowWrap.className = 'url-rule-row';
            rowWrap.style.flexWrap = 'wrap';

            const enabledCheckbox = document.createElement('input');
            enabledCheckbox.type = 'checkbox';
            enabledCheckbox.checked = !!r.enabled;
            enabledCheckbox.title = 'Включить/выключить правило';
            enabledCheckbox.style.flex = '0 0 18px';
            enabledCheckbox.addEventListener('change', () => {
                updateResultReplacement(idx, findInput.value, replaceInput.value, enabledCheckbox.checked);
                validateRegexInline(findInput, errorDiv);
            });

            const findInput = document.createElement('input');
            findInput.type = 'text';
            findInput.value = r.find || '';
            findInput.id = `result-rule-find-${idx}`;
            findInput.placeholder = 'Найти (regex)';
            findInput.style.flex = '1 1 50%';
            findInput.style.fontFamily = 'monospace';

            const replaceInput = document.createElement('input');
            replaceInput.type = 'text';
            replaceInput.value = r.replace || '';
            replaceInput.id = `result-rule-rep-${idx}`;
            replaceInput.placeholder = 'Будет заменён на';
            replaceInput.style.flex = '1 1 50%';
            replaceInput.style.fontFamily = 'monospace';

            const actions = document.createElement('div');
            actions.className = 'rule-actions';

            const deleteBtn = document.createElement('button');
            deleteBtn.textContent = 'Удалить';
            deleteBtn.className = 'rule-delete-btn';
            deleteBtn.onclick = () => {
                if (confirm('Удалить правило?')) deleteResultReplacement(idx);
            };

            actions.appendChild(deleteBtn);

            rowWrap.appendChild(enabledCheckbox);
            rowWrap.appendChild(findInput);
            rowWrap.appendChild(replaceInput);
            rowWrap.appendChild(actions);

            const errorDiv = document.createElement('div');
            errorDiv.className = 'rule-error';
            errorDiv.style.display = 'none';

            findInput.addEventListener('blur', () => {
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
                updateResultReplacement(idx, findInput.value, replaceInput.value, enabledCheckbox.checked);
            });
            findInput.addEventListener('change', () => { findInput.blur(); });

            replaceInput.addEventListener('blur', () => {
                updateResultReplacement(idx, findInput.value, replaceInput.value, enabledCheckbox.checked);
            });
            replaceInput.addEventListener('change', () => { replaceInput.blur(); });

            validateRegexInline(findInput, errorDiv);

            container.appendChild(rowWrap);
            container.appendChild(errorDiv);
        });
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

    /* ----------------- Sites UI (minimal, unchanged functionality) ----------------- */
    function renderSites() {
        const container = document.getElementById('sites-container');
        if (!container) return;
        container.textContent = '';
        const sites = getSites();
        sites.forEach((site, index) => {
            container.appendChild(createSiteBlock(site, index));
            if (index < sites.length - 1) {
                const divider = document.createElement('div');
                divider.className = 'sites-divider';
                container.appendChild(divider);
            }
        });
    }
    function createSiteBlock(site, index) {
        const block = document.createElement('div');
        block.className = 'site-block';
        block.dataset.index = index;

        const nameDisplay = document.createElement('div');
        nameDisplay.className = 'site-name-display';
        nameDisplay.textContent = site.name || 'Untitled';

        const urlInput = document.createElement('input');
        urlInput.type = 'text';
        urlInput.value = site.url || '';
        urlInput.placeholder = 'URL';
        urlInput.addEventListener('change', (e) => { updateSiteUrl(index, e.target.value); });

        const selectorsLabel = document.createElement('div');
        selectorsLabel.textContent = 'Селекторы:';
        selectorsLabel.style.fontWeight = '600';

        const selectorsDiv = document.createElement('div');
        (site.selectors || []).forEach((sel, sIndex) => {
            const selRow = document.createElement('div');
            selRow.className = 'selector-item';
            const selInput = document.createElement('input');
            selInput.type = 'text';
            selInput.value = sel || '';
            selInput.placeholder = 'CSS селектор';
            selInput.style.flex = '1';
            selInput.addEventListener('change', (e) => updateSelector(index, sIndex, e.target.value));
            const rem = document.createElement('button');
            rem.className = 'remove-selector-btn';
            rem.textContent = '✕';
            rem.onclick = () => removeSelector(index, sIndex);
            selRow.appendChild(selInput);
            selRow.appendChild(rem);
            selectorsDiv.appendChild(selRow);
        });

        const addSelBtn = document.createElement('button');
        addSelBtn.className = 'add-selector-btn';
        addSelBtn.textContent = '+ Добавить селектор';
        addSelBtn.onclick = () => { addSelector(index); };

        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'secondary-btn';
        deleteBtn.textContent = 'Удалить сайт';
        deleteBtn.onclick = () => deleteSite(index);

        block.appendChild(nameDisplay);
        block.appendChild(urlInput);
        block.appendChild(selectorsLabel);
        block.appendChild(selectorsDiv);
        block.appendChild(addSelBtn);
        block.appendChild(deleteBtn);

        return block;
    }
    function addNewSite() {
        const sites = getSites();
        const newName = generateUniqueSiteName(sites, sites.length, 'https://example.com');
        sites.push({ name: newName, url: 'https://', selectors: [''] });
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
    function updateSelector(siteIndex, selectorIndex, value) { const sites = getSites(); sites[siteIndex].selectors[selectorIndex] = value; saveSites(sites); }
    function addSelector(siteIndex) { const sites = getSites(); sites[siteIndex].selectors.push(''); saveSites(sites); renderSites(); }
    function removeSelector(siteIndex, selectorIndex) { const sites = getSites(); if (sites[siteIndex].selectors.length === 1) { alert('Нельзя удалить последний селектор!'); return; } sites[siteIndex].selectors.splice(selectorIndex, 1); saveSites(sites); renderSites(); }

    /* ----------------- Draggable ----------------- */
    function makeDraggable(elem) {
        let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        const header = document.getElementById('scraper-header');
        if (!header) return;
        header.onmousedown = dragMouseDown;
        function dragMouseDown(e) { e.preventDefault(); pos3 = e.clientX; pos4 = e.clientY; document.onmouseup = closeDragElement; document.onmousemove = elementDrag; }
        function elementDrag(e) { e.preventDefault(); pos1 = pos3 - e.clientX; pos2 = pos4 - e.clientY; pos3 = e.clientX; pos4 = e.clientY; elem.style.top = (elem.offsetTop - pos2) + "px"; elem.style.left = (elem.offsetLeft - pos1) + "px"; elem.style.bottom = "auto"; elem.style.right = "auto"; }
        function closeDragElement() { document.onmouseup = null; document.onmousemove = null; }
    }

    /* ----------------- Scraping logic & listeners (with per-value replacement) ----------------- */
    function onLastScrapedSite(name, oldValue, newValue, remote) {
        try {
            const parsed = newValue ? JSON.parse(newValue) : null;
            if (!parsed || typeof parsed.index === 'undefined') return;
            const queueStr = GM_getValue('scrapeQueue');
            const queue = queueStr ? JSON.parse(queueStr) : [];
            const nextIndex = parsed.index + 1;
            const status = document.getElementById && document.getElementById('scraper-status');
            if (status) {
                status.textContent = `Скопировано: ${parsed.siteName} (${parsed.index + 1}/${queue.length})`;
                status.style.color = '#0d652d';
                updateStatusVisibility(status);
            }
            if (nextIndex < queue.length) {
                setTimeout(() => openSiteByIndex(nextIndex), 200);
            } else {
                GM_setValue('scrapeActive', false);
                if (status) {
                    status.textContent = `Готово: обработано ${queue.length} сайт(ов).`;
                    status.style.color = '#0d652d';
                    updateStatusVisibility(status);
                }
            }
        } catch (e) { console.error('Ошибка в onLastScrapedSite:', e); }
    }

    function startScraping() {
        const sites = getSites();
        const status = document.getElementById && document.getElementById('scraper-status');

        if (GM_getValue('scrapeActive')) {
            if (status) {
                status.textContent = 'Скрейпинг уже выполняется — дождитесь завершения.';
                status.style.color = '#d39e00';
                updateStatusVisibility(status);
            }
            return;
        }

        for (let i = 0; i < sites.length; i++) {
            if (!sites[i].url || sites[i].url === 'https://') { alert(`Сайт "${sites[i].name}" имеет невалидный URL!`); return; }
            if (!sites[i].selectors || !sites[i].selectors.length || sites[i].selectors.every(s => !s.trim())) { alert(`Сайт "${sites[i].name}" не имеет селекторов!`); return; }
        }

        if (status) {
            status.textContent = `Открываем 1/${sites.length} сайт(ов)...`;
            status.style.color = '#1a73e8';
            updateStatusVisibility(status);
        }

        GM_setValue('scrapeActive', true);
        GM_setValue('scrapeQueue', JSON.stringify(sites));
        GM_setValue('scrapedClipboard', '');
        GM_setValue('scrapedDataArray', JSON.stringify([]));
        try { GM_setClipboard(''); } catch (e) {}

        openSiteByIndex(0);

        const guardKey = 'scrapeGuardTimer';
        try { clearTimeout(window[guardKey]); } catch (e) {}
        window[guardKey] = setTimeout(() => {
            if (GM_getValue('scrapeActive')) {
                GM_setValue('scrapeActive', false);
                if (status) {
                    status.textContent = 'Прервано: время ожидания истекло.';
                    status.style.color = '#d32f2f';
                    updateStatusVisibility(status);
                }
            }
        }, 2 * 60 * 1000);
    }

    function openSiteByIndex(index) {
        const queueStr = GM_getValue('scrapeQueue');
        const queue = queueStr ? JSON.parse(queueStr) : [];
        if (index < 0 || index >= queue.length) return;
        const site = queue[index];

        const openUrl = applyUrlReplacements(site.url);

        try { GM_setValue('currentOpeningIndex', index); } catch (e) {}
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

            const timeoutMs = 20000;
            const pollInterval = 250;
            const selectors = Array.isArray(siteConfig.selectors) ? siteConfig.selectors : [];
            let values = await waitForSelectorsAndCollect(selectors, timeoutMs, pollInterval);

            // Apply replacements to each value individually
            const cleanedValues = values.map(value => applyResultReplacements(value || ''));

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

    /* ----------------- Init UI on Google Sheets ----------------- */
    if (isGoogleSheets) {
        createFloatingWindow();
        if (!window.__scraper_listener_registered) {
            window.__scraper_listener_registered = true;
            try { window.__scraper_listener_id = GM_addValueChangeListener('lastScrapedSite', onLastScrapedSite); } catch (e) {}
        }
    }

    function updateStatusVisibility(status) {
        if (status.textContent.trim() === '') {
            status.style.display = 'none';
        } else {
            status.style.display = '';
        }
    }

})();