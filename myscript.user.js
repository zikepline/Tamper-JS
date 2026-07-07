// ==UserScript==
// @name         CampaignsParser
// @namespace    https://github.com/zikepline
// @version      1.1.0
// @description  Mass parser for campaigns. Open https://admin.convertagain.com/?tm_ca_panel=2
// @author       zikepline
//
// @match        https://admin.convertagain.com/*
// @run-at       document-idle
//
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addValueChangeListener
// @grant        GM_openInTab
//
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
//
// @downloadURL  https://raw.githubusercontent.com/YOUR_GITHUB_USERNAME/CampaignsParser/main/CampaignsParser.user.js
// @updateURL    https://raw.githubusercontent.com/YOUR_GITHUB_USERNAME/CampaignsParser/main/CampaignsParser.user.js
// ==/UserScript==

(function () {
    'use strict';

    /************************************************************
     * CONFIG
     ************************************************************/

    const CONFIG = {
        panelParam: 'tm_ca_panel=2',
        workerTokenParam: 'tm_ca_worker_token',

        defaultBaseHost: 'admin.convertagain.com',

        baseHostOptions: [
            { label: 'ConvertAgain', host: 'admin.convertagain.com' }
        ],

        maxAttemptsPerCampaign: 3,

        // Увеличено: при 1000+ задачах браузер может подвисать, 15 сек было слишком жёстко.
        workerTimeoutMs: 45000,

        pageReadyTimeoutMs: 10000,
        campaignInputTimeoutMs: 10000,
        editorReadyTimeoutMs: 10000,

        closeWorkerTabAfterDone: true,
        closeWorkerTabAfterError: true,

        // false — вкладки открываются в фоне.
        // true — браузер будет переключаться на открываемую вкладку.
        openWorkerTabActive: false,

        defaultParallelWorkers: 10,
        maxParallelWorkersHardLimit: 10,

        // LIGHT MODE
        // Для больших очередей live-preview лучше держать выключенным.
        renderLivePreviewTable: true,
        livePreviewLimit: 30,

        queueRowsBeforeActive: 5,
        queueRowsAfterActive: 20,
        queueRenderIntervalMs: 2500,

        maxLogLines: 300,

        // Worker-вкладки не будут спамить INFO-логами в панель.
        silentWorkerInfoLogs: true
    };

    const KEYS = {
        queue: 'ca_parser_queue_ids_v6_light',
        nextIndex: 'ca_parser_next_index_v6_light',
        running: 'ca_parser_queue_running_v6_light',
        stopRequested: 'ca_parser_stop_requested_v6_light',

        activeJobs: 'ca_parser_active_jobs_v6_light',
        attempts: 'ca_parser_attempts_map_v6_light',
        statuses: 'ca_parser_statuses_map_v6_light',

        // Новый формат хранения результатов:
        // не один огромный объект, а отдельный ключ на каждую кампанию.
        resultKeyPrefix: 'ca_parser_result_item_v6_light_',
        resultsCount: 'ca_parser_results_count_v6_light',
        currentRunId: 'ca_parser_current_run_id_v6_light',

        workerStatusEvent: 'ca_parser_worker_status_event_v6_light',
        logEvent: 'ca_parser_log_event_v6_light',

        startedAt: 'ca_parser_started_at_v6_light',
        doneCount: 'ca_parser_done_count_v6_light',
        errorCount: 'ca_parser_error_count_v6_light',
        totalCount: 'ca_parser_total_count_v6_light',
        parallelWorkers: 'ca_parser_parallel_workers_v6_light',
        selectedBaseHost: 'ca_parser_selected_base_host_v6_light'
    };

    const FIELD_MAPPING = {
        campaign_id: 'campaign_id',
        campaignName: 'Naming',
        clickUrl: 'Click URL',
        topLevelDomain: 'Top Level Domain',
        CAMPAIGN_TYPE: 'Type',
        ctr: 'CTR, %',
        advertiserName: 'Advertiser Name',

        campaignPurpose: 'Campaign Purpose',
        deliveryType: 'Delivery Type',

        frequencyCapping: 'Freq Cap',
        interval: 'Interval',
        frequencyCapValue: 'Limit',

        frequencyCappingImp2: 'Freq Cap2',
        intervalImp2: 'Interval2',
        frequencyCapValueImp2: 'Limit2',

        fqCapClick: 'Freq CapPush',
        intervalClick: 'IntervalPush',
        fqCapClickValue: 'LimitPush',

        countryMode: 'Country Mode',
        countrySelected: 'Country Selected',

        cityMode: 'City Mode',
        citySelected: 'City Selected',

        languageMode: 'Language Mode',
        languageSelected: 'Language Selected',

        deviceTypeMode: 'Device Type Mode',
        deviceTypeSelected: 'Device Types Selected',

        audienceMatchingType: 'Audience Matching Type',
        audienceInclude: 'Audience Include',
        audienceExclude: 'Audience Exclude',
        whitelist: 'Whitelist',
        blacklist: 'Blacklist',

        isDayPartingEnabled: 'Campaign Scheduling',
        dayPartingSchedule: 'Расписание',
        schedulingStartDate: 'Start Date',
        schedulingEndDate: 'End Date',

        status: 'Статус',
        error: 'Ошибка',
        url: 'URL страницы'
    };

    const COLUMN_ORDER = [
        'campaign_id',
        'Naming',
        'Advertiser Name',
        'Type',
        'Click URL',
        'Top Level Domain',
        'CTR, %',
        'Campaign Purpose',
        'Delivery Type',

        'Freq Cap',
        'Interval',
        'Limit',

        'Freq Cap2',
        'Interval2',
        'Limit2',

        'Freq CapPush',
        'IntervalPush',
        'LimitPush',

        'Country Mode',
        'Country Selected',

        'City Mode',
        'City Selected',

        'Language Mode',
        'Language Selected',

        'Device Type Mode',
        'Device Types Selected',

        'Audience Matching Type',
        'Audience Include',
        'Audience Exclude',
        'Whitelist',
        'Blacklist',

        'Campaign Scheduling',
        'Расписание',
        'Start Date',
        'End Date',

        'Статус',
        'Ошибка',
        'URL страницы'
    ];

    const isPanel = location.href.includes(CONFIG.panelParam);

    const isCampaignEditPage = isCampaignEditPageUrl(location.href);

    /************************************************************
     * COMMON HELPERS
     ************************************************************/

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    function makeToken(index, id) {
        return `${Date.now()}_${index}_${id}_${Math.random().toString(16).slice(2)}`;
    }

    function getWorkerTokenFromUrl() {
        try {
            const url = new URL(location.href);
            return url.searchParams.get(CONFIG.workerTokenParam);
        } catch (e) {
            return null;
        }
    }

    function getCampaignIdFromUrl(url) {
        const match = String(url || '').match(/\/campaigns\/([^/]+)\/edit/);
        return match ? decodeURIComponent(match[1]) : null;
    }

    function normalizeBaseHost(value) {
        return String(value || '')
            .trim()
            .replace(/^https?:\/\//i, '')
            .replace(/\/.*$/g, '')
            .trim();
    }

    function getSelectedBaseHost() {
        const saved = GM_getValue(KEYS.selectedBaseHost, CONFIG.defaultBaseHost);
        return normalizeBaseHost(saved || CONFIG.defaultBaseHost);
    }

    function setSelectedBaseHost(host) {
        const normalized = normalizeBaseHost(host);

        if (normalized) {
            GM_setValue(KEYS.selectedBaseHost, normalized);
        }

        return normalized;
    }

    function getCampaignEditUrl(id, baseHost = getSelectedBaseHost()) {
        return `https://${normalizeBaseHost(baseHost)}/campaigns/${encodeURIComponent(id)}/edit`;
    }

    function buildCampaignWorkerUrl(id, token, baseHost = getSelectedBaseHost()) {
        return `${getCampaignEditUrl(id, baseHost)}?${CONFIG.workerTokenParam}=${encodeURIComponent(token)}`;
    }

    function isCampaignEditPageUrl(url) {
        try {
            const parsed = new URL(url);
            return /^\/campaigns\/[^/]+\/edit$/.test(parsed.pathname);
        } catch (e) {
            return false;
        }
    }

    function parseInputLines(text) {
        const seen = new Set();

        return String(text || '')
            .split('\n')
            .map(x => x.trim())
            .filter(Boolean)
            .map(line => {
                const fromUrl = getCampaignIdFromUrl(line);
                return fromUrl || line;
            })
            .map(x => x.trim())
            .filter(Boolean)
            .filter(id => {
                if (seen.has(id)) return false;
                seen.add(id);
                return true;
            });
    }

    function nowText() {
        return new Date().toLocaleTimeString();
    }

    function formatDateForFile() {
        const d = new Date();
        const pad = n => String(n).padStart(2, '0');

        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}-${pad(d.getMinutes())}-${pad(d.getSeconds())}`;
    }

    function formatDuration(ms) {
        const totalSeconds = Math.max(0, Math.round(ms / 1000));

        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;

        if (hours > 0) {
            return `${hours} ч ${minutes} мин ${seconds} сек`;
        }

        if (minutes > 0) {
            return `${minutes} мин ${seconds} сек`;
        }

        return `${seconds} сек`;
    }

    function clampWorkers(value) {
        const n = Number.parseInt(value, 10);

        if (!Number.isFinite(n) || n < 1) return 1;
        if (n > CONFIG.maxParallelWorkersHardLimit) return CONFIG.maxParallelWorkersHardLimit;

        return n;
    }

    function resetQueueState() {
        GM_setValue(KEYS.queue, []);
        GM_setValue(KEYS.nextIndex, 0);
        GM_setValue(KEYS.running, false);
        GM_setValue(KEYS.stopRequested, false);

        GM_setValue(KEYS.activeJobs, {});
        GM_setValue(KEYS.attempts, {});
        GM_setValue(KEYS.statuses, {});

        GM_setValue(KEYS.resultsCount, 0);
        GM_setValue(KEYS.currentRunId, null);

        GM_setValue(KEYS.workerStatusEvent, null);
        GM_setValue(KEYS.startedAt, null);
        GM_setValue(KEYS.doneCount, 0);
        GM_setValue(KEYS.errorCount, 0);
        GM_setValue(KEYS.totalCount, 0);
    }

    function emitLog(level, message, extra = {}) {
        if (CONFIG.silentWorkerInfoLogs && !isPanel && level === 'INFO') {
            return;
        }

        GM_setValue(KEYS.logEvent, {
            eventId: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
            time: Date.now(),
            level,
            message,
            url: location.href,
            ...extra
        });
    }

    function emitWorkerStatus(status, extra = {}) {
        GM_setValue(KEYS.workerStatusEvent, {
            eventId: `${Date.now()}_${Math.random().toString(16).slice(2)}`,
            time: Date.now(),
            status,
            url: location.href,
            ...extra
        });
    }

    function isStopRequested() {
        return GM_getValue(KEYS.stopRequested, false) === true ||
            GM_getValue(KEYS.running, false) !== true;
    }

    function throwIfStopped() {
        if (isStopRequested()) {
            throw new Error('STOP_REQUESTED');
        }
    }

    async function waitFor(fn, timeout = 5000, step = 300) {
        const start = Date.now();

        while (Date.now() - start < timeout) {
            throwIfStopped();

            const res = fn();
            if (res) return res;

            await sleep(step);
        }

        return null;
    }

    async function waitForPageReady(timeout = CONFIG.pageReadyTimeoutMs) {
        const ok = await waitFor(() => {
            return document.readyState === 'complete' && document.body;
        }, timeout, 500);

        if (!ok) {
            throw new Error(`PAGE_NOT_READY_AFTER_${timeout}_MS`);
        }
    }

    async function waitForDomStable(timeout = 10000, stableMs = 700) {
        return new Promise((resolve, reject) => {
            const start = Date.now();
            let lastChange = Date.now();

            if (!document.body) {
                resolve(false);
                return;
            }

            const observer = new MutationObserver(() => {
                lastChange = Date.now();
            });

            observer.observe(document.body, {
                childList: true,
                subtree: true,
                attributes: true,
                characterData: true
            });

            const timer = setInterval(() => {
                try {
                    throwIfStopped();

                    const now = Date.now();

                    if (now - lastChange >= stableMs) {
                        clearInterval(timer);
                        observer.disconnect();
                        resolve(true);
                        return;
                    }

                    if (now - start >= timeout) {
                        clearInterval(timer);
                        observer.disconnect();
                        resolve(false);
                    }
                } catch (e) {
                    clearInterval(timer);
                    observer.disconnect();
                    reject(e);
                }
            }, 300);
        });
    }

    function escapeHtml(str) {
        return String(str ?? '')
            .replaceAll('&', '&amp;')
            .replaceAll('<', '&lt;')
            .replaceAll('>', '&gt;')
            .replaceAll('"', '&quot;')
            .replaceAll("'", '&#039;');
    }

    function waitForBody(callback) {
        if (document.body) {
            callback();
            return;
        }

        const timer = setInterval(() => {
            if (document.body) {
                clearInterval(timer);
                callback();
            }
        }, 100);
    }

    function renameResultFields(raw) {
        const out = {};

        Object.entries(raw || {}).forEach(([key, value]) => {
            if (key === '__runId') return;

            const mapped = FIELD_MAPPING[key] || key;
            out[mapped] = value ?? '';
        });

        return out;
    }

    function getStoredResultById(id) {
        const currentRunId = GM_getValue(KEYS.currentRunId, null);
        const result = GM_getValue(KEYS.resultKeyPrefix + id, null);

        if (!result) return null;

        if (currentRunId && result.__runId !== currentRunId) {
            return null;
        }

        return result;
    }

    function getOrderedResultsArray() {
        const ids = GM_getValue(KEYS.queue, []);
        const rows = [];

        ids.forEach(id => {
            const result = getStoredResultById(id);

            if (result) {
                rows.push(renameResultFields(result));
            }
        });

        return rows;
    }

    function getPreviewResultsArray(limit = CONFIG.livePreviewLimit) {
        const ids = GM_getValue(KEYS.queue, []);
        const rows = [];

        for (const id of ids) {
            const result = getStoredResultById(id);

            if (result) {
                rows.push(renameResultFields(result));
            }

            if (rows.length >= limit) {
                break;
            }
        }

        return rows;
    }

    function downloadExcelLikeFile() {
        const rows = getOrderedResultsArray();

        if (!rows.length) {
            alert('Нет данных для скачивания. Сначала запусти парсинг.');
            return;
        }

        if (typeof XLSX === 'undefined') {
            alert('Библиотека XLSX не загрузилась. Проверь @require в шапке Tampermonkey-скрипта.');
            return;
        }

        const headers = COLUMN_ORDER.filter(header =>
            rows.some(row => Object.prototype.hasOwnProperty.call(row, header))
        );

        const aoa = [];

        aoa.push(headers);

        rows.forEach(row => {
            aoa.push(headers.map(header => row[header] ?? ''));
        });

        const worksheet = XLSX.utils.aoa_to_sheet(aoa);

        worksheet['!freeze'] = {
            xSplit: 2,
            ySplit: 1,
            topLeftCell: 'C2',
            activePane: 'bottomRight',
            state: 'frozen'
        };

        worksheet['!cols'] = headers.map(header => {
            const widthMap = {
                'campaign_id': 12,
                'Naming': 70,
                'Advertiser Name': 22,
                'Type': 14,
                'Click URL': 70,
                'Top Level Domain': 15,
                'CTR, %': 6,

                'Campaign Purpose': 18,
                'Delivery Type': 18,

                'Freq Cap': 20,
                'Interval': 12,
                'Limit': 10,

                'Freq Cap2': 22,
                'Interval2': 12,
                'Limit2': 10,

                'Freq CapPush': 18,
                'IntervalPush': 14,
                'LimitPush': 12,

                'Country Mode': 14,
                'Country Selected': 28,

                'City Mode': 12,
                'City Selected': 28,

                'Language Mode': 16,
                'Language Selected': 24,

                'Device Type Mode': 18,
                'Device Types Selected': 26,

                'Audience Matching Type': 24,
                'Audience Include': 28,
                'Audience Exclude': 28,
                'Whitelist': 28,
                'Blacklist': 28,

                'Campaign Scheduling': 20,
                'Расписание': 45,
                'Start Date': 14,
                'End Date': 14,

                'Статус': 12,
                'Ошибка': 40,
                'URL страницы': 45
            };

            return {
                wch: widthMap[header] || 15
            };
        });

        const range = XLSX.utils.decode_range(worksheet['!ref']);

        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });

            if (!worksheet[cellAddress]) continue;

            worksheet[cellAddress].s = {
                font: {
                    bold: true
                },
                alignment: {
                    vertical: 'center',
                    horizontal: 'center',
                    wrapText: true
                }
            };
        }

        for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });

                if (!worksheet[cellAddress]) continue;

                worksheet[cellAddress].s = {
                    ...(worksheet[cellAddress].s || {}),
                    alignment: {
                        vertical: 'top',
                        wrapText: true
                    }
                };
            }
        }

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Parsing Results');

        const fileName = `parsing_results_${formatDateForFile()}.xlsx`;

        XLSX.writeFile(workbook, fileName, {
            bookType: 'xlsx',
            type: 'binary',
            cellStyles: true
        });
    }

    /************************************************************
     * PANEL
     ************************************************************/

    function runPanel() {
        waitForBody(initPanel);
    }

    function initPanel() {
        let activeTabs = {};
        let watchdogTimer = null;
        let queueRenderTimer = null;
        let showFullQueue = false;
        let lastResultsRenderCount = -1;

        if (GM_getValue(KEYS.running, false) !== true) {
            resetQueueState();
        }

        document.title = 'ConvertAgain Campaign Parser';

        const savedBaseHost = getSelectedBaseHost();

        const isSavedHostPreset = CONFIG.baseHostOptions
        .some(option => option.host === savedBaseHost);

        const baseHostOptionsHtml = CONFIG.baseHostOptions.map((option, index) => {
            const checked = option.host === savedBaseHost ? 'checked' : '';

            return `
        <label class="ca-radio-line">
            <input
                type="radio"
                name="ca-base-host"
                value="${escapeHtml(option.host)}"
                ${checked}
            >
            <span>${escapeHtml(option.label)}</span>
            <small>${escapeHtml(option.host)}</small>
        </label>
    `;
        }).join('');

        const customHostChecked = !isSavedHostPreset ? 'checked' : '';
        const customHostValue = !isSavedHostPreset ? savedBaseHost : '';
        document.body.innerHTML = `
            <div id="ca-panel">
                <div class="ca-card">
                    <h1>ConvertAgain Campaign Parser</h1>

                    <p class="ca-muted">
                        Вставь ID кампаний или ссылки на edit-страницы, каждый с новой строки.
                    </p>

                    <div class="ca-grid">
                        <div>
                            <label class="ca-label">Campaign IDs / URLs</label>
                            <textarea id="ca-input" placeholder="Например:
57516
5565
https://admin.convertagain.com/campaigns/1234/edit"></textarea>
                        </div>

                        <div class="ca-side">
                            <label class="ca-label">Threads / потоки</label>
                            <input id="ca-workers" type="number" min="1" max="${CONFIG.maxParallelWorkersHardLimit}" value="${GM_getValue(KEYS.parallelWorkers, CONFIG.defaultParallelWorkers)}">
                            <p class="ca-hint">
                                Максимум: ${CONFIG.maxParallelWorkersHardLimit}.<br>
                                Для 1000+ кампаний лучше 6–10 потоков, если браузер держит нагрузку.
                            </p>

                            <div class="ca-site-box">
                                <label class="ca-label">Сайт для парсинга</label>

                                ${baseHostOptionsHtml}

                                <label class="ca-radio-line ca-radio-custom">
                                    <input
                                        type="radio"
                                        name="ca-base-host"
                                        value="__custom__"
                                        ${customHostChecked}
                                    >
                                    <span>Свой вариант</span>
                                </label>

                                <input
                                    id="ca-custom-host"
                                    type="text"
                                    placeholder="admin.convertagain.com"
                                    value="${escapeHtml(customHostValue)}"
                                >

                                <p class="ca-hint">
                                    Вводи без https://, например: admin.convertagain.com
                                </p>
                            </div>
                        </div>
                    </div>

                    <div class="ca-actions">
                        <button id="ca-start" class="ca-btn ca-start">Start parsing</button>
                        <button id="ca-stop" class="ca-btn ca-stop">Stop / Esc</button>
                        <button id="ca-reset" class="ca-btn ca-reset">Reset queue</button>
                        <button id="ca-show-all" class="ca-btn ca-show-all">Show all queue</button>
                        <button id="ca-download" class="ca-btn ca-download">Download Excel</button>
                        <button id="ca-clear" class="ca-btn ca-clear">Clear logs</button>
                    </div>

                    <div class="ca-stats">
                        <div><b>Статус:</b> <span id="ca-status">idle</span></div>
                        <div><b>Прогресс:</b> <span id="ca-progress">0 / 0</span></div>
                        <div><b>Активных потоков:</b> <span id="ca-active-workers">0</span></div>
                        <div><b>Успешно:</b> <span id="ca-done">0</span></div>
                        <div><b>Ошибок:</b> <span id="ca-errors">0</span></div>
                        <div><b>Время:</b> <span id="ca-time">—</span></div>
                    </div>

                    <h2>Queue</h2>
                    <div id="ca-queue"></div>

                    <h2>Parsed results</h2>
                    <div id="ca-results"></div>

                    <h2>Logs</h2>
                    <pre id="ca-logs"></pre>
                </div>
            </div>
        `;

        const style = document.createElement('style');
        style.textContent = `
            html, body {
                margin: 0 !important;
                padding: 0 !important;
                background: #f4f4f4 !important;
                font-family: Arial, sans-serif !important;
            }

            #ca-panel {
                padding: 24px;
            }

            .ca-card {
                max-width: 1200px;
                margin: 0 auto;
                background: #fff;
                border-radius: 14px;
                padding: 24px;
                box-shadow: 0 4px 20px rgba(0,0,0,.12);
            }

            h1 {
                margin: 0 0 8px;
                font-size: 24px;
            }

            h2 {
                margin: 22px 0 10px;
                font-size: 18px;
            }

            .ca-muted, .ca-hint {
                color: #666;
                line-height: 1.45;
            }

            .ca-hint {
                font-size: 13px;
                margin-top: 8px;
            }

            .ca-grid {
                display: grid;
                grid-template-columns: minmax(0, 1fr) 220px;
                gap: 16px;
                align-items: start;
            }

            .ca-label {
                display: block;
                margin-bottom: 8px;
                font-weight: 700;
                font-size: 14px;
            }

            #ca-input {
                width: 100%;
                height: 380px;
                box-sizing: border-box;
                padding: 14px;
                font-size: 14px;
                border: 1px solid #ccc;
                border-radius: 10px;
                resize: vertical;
                outline: none;
            }

            #ca-workers {
                width: 100%;
                box-sizing: border-box;
                padding: 12px;
                font-size: 18px;
                border: 1px solid #ccc;
                border-radius: 10px;
            }

            .ca-actions {
                display: flex;
                flex-wrap: wrap;
                gap: 10px;
                margin: 14px 0;
            }

            .ca-btn {
                border: none;
                border-radius: 10px;
                padding: 12px 18px;
                color: #fff;
                font-size: 14px;
                font-weight: 700;
                cursor: pointer;
            }

            .ca-start { background: #0a8f2d; }
            .ca-stop { background: #d90000; }
            .ca-reset { background: #7952b3; }
            .ca-clear { background: #333; }
            .ca-show-all { background: #006d9c; }
            .ca-download { background: #0b6f44; }

.ca-site-box {
    margin-top: 18px;
    padding-top: 14px;
    border-top: 1px solid #ddd;
}

.ca-radio-line {
    display: grid;
    grid-template-columns: 18px 1fr;
    gap: 8px;
    align-items: center;
    margin: 8px 0;
    font-size: 14px;
    cursor: pointer;
}

.ca-radio-line input {
    margin: 0;
}

.ca-radio-line small {
    grid-column: 2;
    color: #777;
    font-size: 12px;
    margin-top: -4px;
}

#ca-custom-host {
    width: 100%;
    box-sizing: border-box;
    margin-top: 8px;
    padding: 10px 12px;
    border: 1px solid #ccc;
    border-radius: 10px;
    font-size: 14px;
}

            #ca-finished-modal {
                position: fixed;
                inset: 0;
                z-index: 9999999;
                display: flex;
                align-items: center;
                justify-content: center;
                background: rgba(0, 0, 0, 0.45);
            }

            .ca-finished-box {
                min-width: 360px;
                max-width: 520px;
                padding: 28px;
                border-radius: 18px;
                background: #fff;
                box-shadow: 0 10px 40px rgba(0, 0, 0, 0.35);
                text-align: center;
            }

            .ca-finished-main {
                display: inline-block;
                padding: 18px 34px;
                border-radius: 14px;
                background: #0a8f2d;
                color: #fff;
                font-size: 26px;
                font-weight: 800;
                line-height: 1.2;
            }

            .ca-finished-info {
                margin-top: 18px;
                color: #333;
                font-size: 15px;
                line-height: 1.6;
            }

            .ca-finished-ok {
                margin-top: 22px;
                min-width: 120px;
                border: none;
                border-radius: 10px;
                padding: 12px 24px;
                background: #222;
                color: #fff;
                font-size: 15px;
                font-weight: 700;
                cursor: pointer;
            }

            .ca-finished-ok:hover {
                background: #000;
            }

            .ca-stats {
                display: grid;
                grid-template-columns: repeat(6, 1fr);
                gap: 10px;
                margin-top: 16px;
                padding: 12px;
                background: #f7f7f7;
                border-radius: 10px;
            }

            #ca-queue {
                border: 1px solid #ddd;
                border-radius: 10px;
                overflow: auto;
                max-height: 360px;
            }

            .ca-row {
                display: grid;
                grid-template-columns: 70px 1fr 110px 140px 160px;
                gap: 10px;
                padding: 8px 12px;
                border-bottom: 1px solid #eee;
                font-size: 13px;
            }

            .ca-row:last-child {
                border-bottom: none;
            }

            .ca-row.running {
                background: #e8f1ff;
                font-weight: 700;
            }

            .ca-row.done {
                background: #e8f7eb;
            }

            .ca-row.error {
                background: #ffe8e8;
            }

            #ca-results {
                border: 1px solid #ddd;
                border-radius: 10px;
                overflow: auto;
                max-height: 160px;
                background: #fff;
            }

            .ca-results-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 12px;
            }

            .ca-results-table th,
            .ca-results-table td {
                border-bottom: 1px solid #eee;
                border-right: 1px solid #eee;
                padding: 6px 8px;
                text-align: left;
                vertical-align: top;
                max-width: 320px;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }

            .ca-results-table th {
                position: sticky;
                top: 0;
                background: #f2f2f2;
                z-index: 1;
                font-weight: 700;
            }

            #ca-logs {
                height: 300px;
                overflow: auto;
                background: #111;
                color: #eee;
                padding: 14px;
                border-radius: 10px;
                font-size: 13px;
                white-space: pre-wrap;
            }

            @media (max-width: 900px) {
                .ca-grid {
                    grid-template-columns: 1fr;
                }

                .ca-stats {
                    grid-template-columns: repeat(2, 1fr);
                }
            }
        `;
        document.head.appendChild(style);

        const input = document.getElementById('ca-input');
        const workersInput = document.getElementById('ca-workers');

        const customHostInput = document.getElementById('ca-custom-host');
        const baseHostRadios = [...document.querySelectorAll('input[name="ca-base-host"]')];

        function readSelectedBaseHostFromUi() {
            const checked = baseHostRadios.find(radio => radio.checked);

            if (!checked) {
                return CONFIG.defaultBaseHost;
            }

            if (checked.value === '__custom__') {
                return normalizeBaseHost(customHostInput.value);
            }

            return normalizeBaseHost(checked.value);
        }

        function saveSelectedBaseHostFromUi() {
            const selectedHost = readSelectedBaseHostFromUi();

            if (!selectedHost) {
                return '';
            }

            return setSelectedBaseHost(selectedHost);
        }
        const startBtn = document.getElementById('ca-start');
        const stopBtn = document.getElementById('ca-stop');
        const resetBtn = document.getElementById('ca-reset');
        const showAllBtn = document.getElementById('ca-show-all');
        const downloadBtn = document.getElementById('ca-download');
        const clearBtn = document.getElementById('ca-clear');

        const statusEl = document.getElementById('ca-status');
        const progressEl = document.getElementById('ca-progress');
        const activeWorkersEl = document.getElementById('ca-active-workers');
        const doneEl = document.getElementById('ca-done');
        const errorsEl = document.getElementById('ca-errors');
        const timeEl = document.getElementById('ca-time');

        const queueEl = document.getElementById('ca-queue');
        const resultsEl = document.getElementById('ca-results');
        const logsEl = document.getElementById('ca-logs');

        function log(level, message) {
            const line = `[${nowText()}] [${level}] ${message}`;

            const oldLines = logsEl.textContent
                ? logsEl.textContent.split('\n').filter(Boolean)
                : [];

            oldLines.push(line);

            const limitedLines = oldLines.slice(-CONFIG.maxLogLines);

            logsEl.textContent = limitedLines.join('\n') + '\n';
            logsEl.scrollTop = logsEl.scrollHeight;
        }

        baseHostRadios.forEach(radio => {
            radio.addEventListener('change', () => {
                const selectedHost = saveSelectedBaseHostFromUi();

                if (selectedHost) {
                    log('INFO', `Выбран сайт для парсинга: ${selectedHost}`);
                }
            });
        });

        customHostInput.addEventListener('input', () => {
            const customRadio = baseHostRadios.find(radio => radio.value === '__custom__');

            if (customRadio) {
                customRadio.checked = true;
            }

            const selectedHost = saveSelectedBaseHostFromUi();

            if (selectedHost) {
                setPanelStatus(`site: ${selectedHost}`);
            }
        });

        function setPanelStatus(text) {
            statusEl.textContent = text;
        }

        function refreshStats() {
            const ids = GM_getValue(KEYS.queue, []);
            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            const doneCount = GM_getValue(KEYS.doneCount, 0);
            const errorCount = GM_getValue(KEYS.errorCount, 0);
            const totalCount = GM_getValue(KEYS.totalCount, ids.length);
            const startedAt = GM_getValue(KEYS.startedAt, null);
            const processed = doneCount + errorCount;

            progressEl.textContent = totalCount
                ? `${processed} / ${totalCount}`
                : '0 / 0';

            activeWorkersEl.textContent = Object.keys(activeJobs).length;
            doneEl.textContent = doneCount;
            errorsEl.textContent = errorCount;

            timeEl.textContent = startedAt
                ? formatDuration(Date.now() - startedAt)
                : '—';

            const resultsCount = GM_getValue(KEYS.resultsCount, 0);

            if (resultsCount !== lastResultsRenderCount) {
                lastResultsRenderCount = resultsCount;
                renderResults();
            }
        }

        function renderResults() {
            const totalParsedRows = GM_getValue(KEYS.resultsCount, 0);

            if (!CONFIG.renderLivePreviewTable) {
                resultsEl.innerHTML = `
                    <div style="padding: 12px; color: #666; line-height: 1.5;">
                        Live preview отключён для ускорения работы.<br>
                        Спарсено строк: <b>${totalParsedRows}</b>.<br>
                        Полный результат доступен через кнопку <b>Download Excel</b>.
                    </div>
                `;
                return;
            }

            const rows = getPreviewResultsArray(CONFIG.livePreviewLimit);

            if (!rows.length) {
                resultsEl.innerHTML = `
                    <div style="padding: 12px; color: #666;">
                        Пока нет спарсенных данных.
                    </div>
                `;
                return;
            }

            const headers = COLUMN_ORDER.filter(header =>
                rows.some(row => Object.prototype.hasOwnProperty.call(row, header))
            );

            const previewNotice = totalParsedRows > CONFIG.livePreviewLimit
                ? `
                    <div style="padding: 10px 12px; color: #555; background: #fff8dc; border-bottom: 1px solid #eee;">
                        Показаны первые ${CONFIG.livePreviewLimit} строк из ${totalParsedRows}.
                        Для полного просмотра скачай Excel-файл.
                    </div>
                `
                : '';

            resultsEl.innerHTML = `
                ${previewNotice}
                <table class="ca-results-table">
                    <thead>
                        <tr>
                            ${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        ${rows.map(row => `
                            <tr>
                                ${headers.map(h => `<td title="${escapeHtml(row[h] ?? '')}">${escapeHtml(row[h] ?? '')}</td>`).join('')}
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            `;
        }

        function getActiveJobByCampaignId(id) {
            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            return Object.values(activeJobs).find(job => job.id === id) || null;
        }

        function getProcessingTimeText(id, status) {
            const job = getActiveJobByCampaignId(id);

            if (job && job.openedAt) {
                return formatDuration(Date.now() - job.openedAt);
            }

            if (status === 'done' || status === 'error') {
                const result = getStoredResultById(id);

                if (result && result.processingTime) {
                    return result.processingTime;
                }
            }

            return '—';
        }

        function getWorkerSlotText(id) {
            const job = getActiveJobByCampaignId(id);

            if (job && job.workerSlot) {
                const maxWorkers = GM_getValue(KEYS.parallelWorkers, CONFIG.defaultParallelWorkers);
                return `${job.workerSlot}/${maxWorkers}`;
            }

            const result = getStoredResultById(id);

            if (result && result.workerSlot) {
                return result.workerSlot;
            }

            return '—';
        }

        function getVisualStatus(id, index, nextIndex, running) {
            const statuses = GM_getValue(KEYS.statuses, {});
            const activeJob = getActiveJobByCampaignId(id);

            if (activeJob) {
                return activeJob.lastStatus || 'running';
            }

            if (statuses[id]) {
                return statuses[id];
            }

            if (running && index >= nextIndex) {
                return 'waiting';
            }

            return 'waiting';
        }

        function renderQueue() {
            refreshStats();

            const ids = GM_getValue(KEYS.queue, []);
            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            const nextIndex = GM_getValue(KEYS.nextIndex, 0);
            const running = GM_getValue(KEYS.running, false);

            if (!ids.length) {
                queueEl.innerHTML = `
                    <div class="ca-row">
                        <div>—</div>
                        <div>Очередь пустая</div>
                        <div>—</div>
                        <div>—</div>
                        <div>idle</div>
                    </div>
                `;
                return;
            }

            const activeIndices = Object.values(activeJobs)
                .map(job => job.index)
                .filter(index => Number.isInteger(index));

            const focusIndex = activeIndices.length
                ? Math.min(...activeIndices)
                : Math.min(nextIndex, ids.length - 1);

            const before = CONFIG.queueRowsBeforeActive;
            const after = CONFIG.queueRowsAfterActive;

            const start = showFullQueue ? 0 : Math.max(0, focusIndex - before);
            const end = showFullQueue ? ids.length : Math.min(ids.length, focusIndex + after + 1);

            let html = '';

            if (!showFullQueue && start > 0) {
                html += `
                    <div class="ca-row">
                        <div>...</div>
                        <div>Скрыто предыдущих: ${start}</div>
                        <div>old</div>
                        <div>—</div>
                        <div>—</div>
                    </div>
                `;
            }

            html += ids.slice(start, end).map((id, offset) => {
                const i = start + offset;
                const st = getVisualStatus(id, i, nextIndex, running);

                let cls = '';

                if (st === 'opening' || st === 'loading' || st === 'running' || st === 'parsing') {
                    cls = 'running';
                }

                if (st === 'done') {
                    cls = 'done';
                }

                if (st === 'error') {
                    cls = 'error';
                }

                const workerSlotText = getWorkerSlotText(id);
                const processingTimeText = getProcessingTimeText(id, st);

                return `
                    <div class="ca-row ${cls}">
                        <div>#${i + 1}</div>
                        <div>${escapeHtml(id)}</div>
                        <div>${escapeHtml(workerSlotText)}</div>
                        <div>${escapeHtml(processingTimeText)}</div>
                        <div>${escapeHtml(st)}</div>
                    </div>
                `;
            }).join('');

            if (!showFullQueue && end < ids.length) {
                html += `
                    <div class="ca-row">
                        <div>...</div>
                        <div>Скрыто следующих: ${ids.length - end}</div>
                        <div>—</div>
                        <div>—</div>
                        <div>waiting</div>
                    </div>
                `;
            }

            queueEl.innerHTML = html;
        }

        function saveStatus(id, status) {
            const statuses = GM_getValue(KEYS.statuses, {});
            statuses[id] = status;
            GM_setValue(KEYS.statuses, statuses);
        }

        function saveResult(id, data) {
            const currentRunId = GM_getValue(KEYS.currentRunId, null);
            const oldResult = getStoredResultById(id);

            GM_setValue(KEYS.resultKeyPrefix + id, {
                ...data,
                __runId: currentRunId
            });

            if (!oldResult) {
                GM_setValue(KEYS.resultsCount, GM_getValue(KEYS.resultsCount, 0) + 1);
            }
        }

        function removeActiveJob(token) {
            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            delete activeJobs[token];
            GM_setValue(KEYS.activeJobs, activeJobs);

            try {
                if (activeTabs[token] && typeof activeTabs[token].close === 'function') {
                    activeTabs[token].close();
                }
            } catch (e) {}

            delete activeTabs[token];
        }

        function stopQueue(reason = 'manual stop') {
            GM_setValue(KEYS.running, false);
            GM_setValue(KEYS.stopRequested, true);

            setPanelStatus('stopped');
            log('STOP', reason);

            Object.keys(activeTabs).forEach(token => {
                try {
                    if (activeTabs[token] && typeof activeTabs[token].close === 'function') {
                        activeTabs[token].close();
                    }
                } catch (e) {}
            });

            activeTabs = {};
            GM_setValue(KEYS.activeJobs, {});
            renderQueue();
        }

        function hardResetPanelState(reason = 'reset') {
            Object.keys(activeTabs).forEach(token => {
                try {
                    if (activeTabs[token] && typeof activeTabs[token].close === 'function') {
                        activeTabs[token].close();
                    }
                } catch (e) {}
            });

            activeTabs = {};
            resetQueueState();

            setPanelStatus('idle');
            log('RESET', reason);
            renderQueue();
        }

        function allocateWorkerSlot() {
            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            const maxWorkers = GM_getValue(KEYS.parallelWorkers, CONFIG.defaultParallelWorkers);

            const usedSlots = new Set(
                Object.values(activeJobs)
                    .map(job => job.workerSlot)
                    .filter(slot => Number.isInteger(slot))
            );

            for (let slot = 1; slot <= maxWorkers; slot++) {
                if (!usedSlots.has(slot)) {
                    return slot;
                }
            }

            return null;
        }

        function openJob(id, index, attempt, workerSlot = null) {
            if (GM_getValue(KEYS.running, false) !== true) return;

            const maxWorkers = GM_getValue(KEYS.parallelWorkers, CONFIG.defaultParallelWorkers);

            if (!workerSlot) {
                workerSlot = allocateWorkerSlot();
            }

            if (!workerSlot) {
                log('WARN', `Нет свободного слота для кампании ${id}`);
                return;
            }

            const token = makeToken(index, id);
            const url = buildCampaignWorkerUrl(id, token);

            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            activeJobs[token] = {
                token,
                id,
                index,
                attempt,
                workerSlot,
                openedAt: Date.now(),
                lastStatusAt: Date.now(),
                lastStatus: 'opening',
                url
            };
            GM_setValue(KEYS.activeJobs, activeJobs);

            saveStatus(id, 'opening');

            log(
                'OPEN',
                `Открываю кампанию ${id} | поток ${workerSlot}/${maxWorkers} | попытка ${attempt}/${CONFIG.maxAttemptsPerCampaign}`
            );

            renderQueue();

            try {
                activeTabs[token] = GM_openInTab(url, {
                    active: CONFIG.openWorkerTabActive,
                    insert: true,
                    setParent: true
                });
            } catch (e) {
                log('ERROR', `Не удалось открыть вкладку для ${id}: ${e.message}`);
                handleFinalError(token, id, 'OPEN_TAB_FAILED');
            }
        }

        function fillWorkers() {
            if (GM_getValue(KEYS.running, false) !== true) return;

            const ids = GM_getValue(KEYS.queue, []);
            let nextIndex = GM_getValue(KEYS.nextIndex, 0);
            const maxWorkers = GM_getValue(KEYS.parallelWorkers, CONFIG.defaultParallelWorkers);

            let activeJobs = GM_getValue(KEYS.activeJobs, {});
            let statuses = GM_getValue(KEYS.statuses, {});
            let attempts = GM_getValue(KEYS.attempts, {});

            while (Object.keys(activeJobs).length < maxWorkers && nextIndex < ids.length) {
                const id = ids[nextIndex];

                if (
                    statuses[id] === 'done' ||
                    statuses[id] === 'error' ||
                    statuses[id] === 'running' ||
                    statuses[id] === 'opening' ||
                    statuses[id] === 'loading' ||
                    statuses[id] === 'parsing'
                ) {
                    nextIndex++;
                    GM_setValue(KEYS.nextIndex, nextIndex);
                    continue;
                }

                const attempt = (attempts[id] || 0) + 1;
                attempts[id] = attempt;
                GM_setValue(KEYS.attempts, attempts);

                nextIndex++;
                GM_setValue(KEYS.nextIndex, nextIndex);

                const workerSlot = allocateWorkerSlot();

                if (!workerSlot) {
                    break;
                }

                openJob(id, nextIndex - 1, attempt, workerSlot);

                activeJobs = GM_getValue(KEYS.activeJobs, {});
                statuses = GM_getValue(KEYS.statuses, {});
                attempts = GM_getValue(KEYS.attempts, {});
            }

            checkFinish();
        }

        function checkFinish() {
            const running = GM_getValue(KEYS.running, false);
            if (!running) return;

            const ids = GM_getValue(KEYS.queue, []);
            const nextIndex = GM_getValue(KEYS.nextIndex, 0);
            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            const doneCount = GM_getValue(KEYS.doneCount, 0);
            const errorCount = GM_getValue(KEYS.errorCount, 0);

            if (nextIndex >= ids.length && Object.keys(activeJobs).length === 0 && doneCount + errorCount >= ids.length) {
                finishQueue();
            }
        }

        function finishQueue() {
            const startedAt = GM_getValue(KEYS.startedAt, null);
            const finishedAt = Date.now();

            const doneCount = GM_getValue(KEYS.doneCount, 0);
            const errorCount = GM_getValue(KEYS.errorCount, 0);
            const totalCount = GM_getValue(KEYS.totalCount, doneCount + errorCount);

            const durationText = startedAt
                ? formatDuration(finishedAt - startedAt)
                : 'неизвестно';

            GM_setValue(KEYS.running, false);
            GM_setValue(KEYS.stopRequested, false);
            GM_setValue(KEYS.activeJobs, {});
            GM_setValue(KEYS.attempts, {});

            setPanelStatus('finished');

            log('DONE', `Парсинг завершен. Успешно: ${doneCount}; ошибок: ${errorCount}; всего: ${totalCount}; время: ${durationText}. Можно скачать Excel.`);

            renderQueue();

            showParsingFinishedModal(doneCount, errorCount, totalCount, durationText);
        }

        function showParsingFinishedModal(doneCount, errorCount, totalCount, durationText) {
            const oldModal = document.getElementById('ca-finished-modal');
            if (oldModal) oldModal.remove();

            const modal = document.createElement('div');
            modal.id = 'ca-finished-modal';

            modal.innerHTML = `
                <div class="ca-finished-box">
                    <div class="ca-finished-main">
                        Парсинг завершен
                    </div>

                    <div class="ca-finished-info">
                        Успешно: <b>${doneCount}</b> &nbsp;|&nbsp;
                        Ошибок: <b>${errorCount}</b> &nbsp;|&nbsp;
                        Всего: <b>${totalCount}</b><br>
                        Время: <b>${escapeHtml(durationText)}</b>
                    </div>

                    <button id="ca-finished-ok" class="ca-finished-ok">
                        OK
                    </button>
                </div>
            `;

            document.body.appendChild(modal);

            document.getElementById('ca-finished-ok').onclick = () => {
                modal.remove();
            };
        }

        function handleFinalError(token, id, reason) {
            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            const job = activeJobs[token];

            const processingTime = job?.openedAt
                ? formatDuration(Date.now() - job.openedAt)
                : '—';

            const maxWorkers = GM_getValue(KEYS.parallelWorkers, CONFIG.defaultParallelWorkers);
            const workerSlot = job?.workerSlot
                ? `${job.workerSlot}/${maxWorkers}`
                : '—';

            removeActiveJob(token);
            saveStatus(id, 'error');

            saveResult(id, {
                campaign_id: id,
                url: getCampaignEditUrl(id),
                status: 'ERROR',
                error: reason,
                workerSlot,
                processingTime
            });

            const errorCount = GM_getValue(KEYS.errorCount, 0) + 1;
            GM_setValue(KEYS.errorCount, errorCount);

            log('ERROR', `${id}: ${reason}`);
            renderQueue();

            setTimeout(() => {
                fillWorkers();
            }, 300);
        }

        function retryJob(token, id, reason) {
            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            const job = activeJobs[token];

            const attempts = GM_getValue(KEYS.attempts, {});
            const currentAttempt = attempts[id] || job?.attempt || 1;

            if (currentAttempt >= CONFIG.maxAttemptsPerCampaign) {
                handleFinalError(token, id, `${reason}; попытки закончились`);
                return;
            }

            removeActiveJob(token);

            saveStatus(id, 'waiting');
            log('RETRY', `${id}: ${reason}; следующая попытка ${currentAttempt + 1}/${CONFIG.maxAttemptsPerCampaign}`);

            setTimeout(() => {
                const activeJobs = GM_getValue(KEYS.activeJobs, {});
                const maxWorkers = GM_getValue(KEYS.parallelWorkers, CONFIG.defaultParallelWorkers);

                if (Object.keys(activeJobs).length >= maxWorkers) {
                    fillWorkers();
                    return;
                }

                const attemptsMap = GM_getValue(KEYS.attempts, {});
                const nextAttempt = (attemptsMap[id] || 0) + 1;
                attemptsMap[id] = nextAttempt;
                GM_setValue(KEYS.attempts, attemptsMap);

                const workerSlot = allocateWorkerSlot();

                if (!workerSlot) {
                    fillWorkers();
                    return;
                }

                openJob(id, -1, nextAttempt, workerSlot);
            }, 300);
        }

        function startWatchdog() {
            if (watchdogTimer) {
                clearInterval(watchdogTimer);
            }

            watchdogTimer = setInterval(() => {
                const running = GM_getValue(KEYS.running, false);
                if (!running) return;

                const activeJobs = GM_getValue(KEYS.activeJobs, {});
                const now = Date.now();

                Object.values(activeJobs).forEach(job => {
                    const age = now - (job.lastStatusAt || job.openedAt || now);

                    if (age > CONFIG.workerTimeoutMs) {
                        log('WARN', `${job.id}: поток завис или не отвечает ${Math.round(age / 1000)} сек.`);
                        retryJob(job.token, job.id, `Вкладка зависла или не ответила ${Math.round(age / 1000)} сек.`);
                    }
                });

                refreshStats();
            }, 1000);
        }

        function startQueueRenderTimer() {
            if (queueRenderTimer) {
                clearInterval(queueRenderTimer);
            }

            queueRenderTimer = setInterval(() => {
                if (GM_getValue(KEYS.running, false) === true) {
                    renderQueue();
                }
            }, CONFIG.queueRenderIntervalMs);
        }

        startBtn.onclick = () => {
            const ids = parseInputLines(input.value);
            const workers = clampWorkers(workersInput.value);

            if (!ids.length) {
                log('WARN', 'Список ID пустой');
                return;
            }

            if (GM_getValue(KEYS.running, false) === true) {
                log('WARN', 'Очередь уже запущена. Сначала нажми Stop или Reset queue.');
                return;
            }

            const selectedBaseHost = saveSelectedBaseHostFromUi();

            if (!selectedBaseHost) {
                log('WARN', 'Не выбран сайт для парсинга. Укажи домен без https://');
                return;
            }

            activeTabs = {};

            GM_setValue(KEYS.queue, ids);
            GM_setValue(KEYS.nextIndex, 0);
            GM_setValue(KEYS.running, true);
            GM_setValue(KEYS.stopRequested, false);

            GM_setValue(KEYS.activeJobs, {});
            GM_setValue(KEYS.attempts, {});
            GM_setValue(KEYS.statuses, {});

            GM_setValue(KEYS.resultsCount, 0);
            GM_setValue(KEYS.currentRunId, `${Date.now()}_${Math.random().toString(16).slice(2)}`);

            GM_setValue(KEYS.workerStatusEvent, null);
            GM_setValue(KEYS.startedAt, Date.now());
            GM_setValue(KEYS.doneCount, 0);
            GM_setValue(KEYS.errorCount, 0);
            GM_setValue(KEYS.totalCount, ids.length);
            GM_setValue(KEYS.parallelWorkers, workers);

            setPanelStatus('running');
            log('START', `Запущен парсинг: ${ids.length} кампаний; потоков: ${workers}; сайт: ${selectedBaseHost}`);

            renderQueue();
            startWatchdog();
            startQueueRenderTimer();
            fillWorkers();
        };

        stopBtn.onclick = () => {
            stopQueue('Остановлено кнопкой Stop');
        };

        resetBtn.onclick = () => {
            hardResetPanelState('Очередь, статусы и результаты сброшены вручную');
        };

        showAllBtn.onclick = () => {
            showFullQueue = !showFullQueue;

            showAllBtn.textContent = showFullQueue
                ? 'Hide full queue'
                : 'Show all queue';

            log('INFO', showFullQueue ? 'Показан полный список очереди' : 'Показан компактный список очереди');
            renderQueue();
        };

        downloadBtn.onclick = () => {
            downloadExcelLikeFile();
            log('INFO', 'Excel-файл сформирован и скачан');
        };

        clearBtn.onclick = () => {
            logsEl.textContent = '';
            log('INFO', 'Логи очищены');
        };

        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                stopQueue('Остановлено клавишей Escape');
            }
        });

        GM_addValueChangeListener(KEYS.logEvent, (name, oldValue, event) => {
            if (!event) return;
            log(event.level || 'INFO', event.message || '');
        });

        GM_addValueChangeListener(KEYS.workerStatusEvent, (name, oldValue, event) => {
            if (!event || !event.token || !event.id) return;

            const activeJobs = GM_getValue(KEYS.activeJobs, {});
            const job = activeJobs[event.token];

            if (!job) {
                return;
            }

            job.lastStatusAt = Date.now();
            job.lastStatus = event.status;
            activeJobs[event.token] = job;
            GM_setValue(KEYS.activeJobs, activeJobs);

            if (
                event.status === 'opening' ||
                event.status === 'loading' ||
                event.status === 'running' ||
                event.status === 'parsing'
            ) {
                setPanelStatus('running');
                return;
            }

            if (event.status === 'done') {
                const activeJobs = GM_getValue(KEYS.activeJobs, {});
                const job = activeJobs[event.token];

                const processingTime = job?.openedAt
                    ? formatDuration(Date.now() - job.openedAt)
                    : '—';

                const maxWorkers = GM_getValue(KEYS.parallelWorkers, CONFIG.defaultParallelWorkers);
                const workerSlot = job?.workerSlot
                    ? `${job.workerSlot}/${maxWorkers}`
                    : '—';

                removeActiveJob(event.token);
                saveStatus(event.id, 'done');

                if (event.data) {
                    saveResult(event.id, {
                        ...event.data,
                        workerSlot,
                        processingTime
                    });
                }

                const doneCount = GM_getValue(KEYS.doneCount, 0) + 1;
                GM_setValue(KEYS.doneCount, doneCount);

                log('SUCCESS', `${event.id}: данные спарсены успешно`);
                renderQueue();

                setTimeout(() => {
                    fillWorkers();
                }, 200);

                return;
            }

            if (event.status === 'error') {
                log('ERROR', `${event.id}: ${event.error || 'unknown error'}`);
                renderQueue();

                retryJob(event.token, event.id, event.error || 'worker error');
                return;
            }

            if (event.status === 'stopped') {
                removeActiveJob(event.token);
                saveStatus(event.id, 'stopped');
                renderQueue();
            }
        });

        window.addEventListener('beforeunload', () => {
            if (watchdogTimer) {
                clearInterval(watchdogTimer);
            }

            if (queueRenderTimer) {
                clearInterval(queueRenderTimer);
            }
        });

        setPanelStatus(GM_getValue(KEYS.running, false) ? 'running' : 'idle');
        renderQueue();
        renderResults();
        log('INFO', 'Панель парсера LIGHT готова');
    }

    /************************************************************
     * WORKER
     ************************************************************/

    async function runWorker() {
        const running = GM_getValue(KEYS.running, false);
        const token = getWorkerTokenFromUrl();

        if (!running || !token) {
            createManualButton();
            return;
        }

        const activeJobs = GM_getValue(KEYS.activeJobs, {});
        const job = activeJobs[token];

        if (!job) {
            emitLog('WARN', `Рабочая вкладка открыта без активной задачи: ${getCampaignIdFromUrl(location.href)}`, { token });
            return;
        }

        const currentId = getCampaignIdFromUrl(location.href);

        try {
            emitWorkerStatus('loading', { token, id: currentId });
            emitLog('INFO', `Рабочая вкладка открыта: ${currentId}`, { token, id: currentId });

            await waitForPageReady();
            await waitForCampaignParserReady(token, currentId);

            emitWorkerStatus('parsing', { token, id: currentId });
            emitLog('INFO', `Начинаю парсинг кампании ${currentId}`, { token, id: currentId });

            const data = await parseCampaignData(token, currentId);

            emitWorkerStatus('done', {
                token,
                id: currentId,
                data
            });

            emitLog('SUCCESS', `Кампания ${currentId} спарсена`, { token, id: currentId });

            if (CONFIG.closeWorkerTabAfterDone) {
                setTimeout(() => {
                    window.close();
                }, 150);
            }

        } catch (e) {
            const msg = e && e.message ? e.message : String(e);

            if (msg === 'STOP_REQUESTED') {
                emitWorkerStatus('stopped', { token, id: currentId });
                emitLog('STOP', `Кампания ${currentId}: остановлено пользователем`, { token, id: currentId });
            } else {
                emitWorkerStatus('error', {
                    token,
                    id: currentId,
                    error: msg
                });
                emitLog('ERROR', `Кампания ${currentId}: ${msg}`, { token, id: currentId });
            }

            if (CONFIG.closeWorkerTabAfterError) {
                setTimeout(() => {
                    window.close();
                }, 150);
            }
        }
    }

    /************************************************************
     * READY CHECKS FOR PARSER
     ************************************************************/

    async function waitForCampaignParserReady(token, id) {
        emitLog('INFO', 'Жду полной загрузки страницы кампании для парсинга', { token, id });

        await waitForPageReady();

        const campaignInput = await waitFor(
            () =>
                document.querySelector('input#campaignName') ||
                document.querySelector('input[name="campaignName"]'),
            CONFIG.campaignInputTimeoutMs,
            300
        );

        if (!campaignInput) {
            throw new Error('CAMPAIGN_NAME_INPUT_NOT_FOUND');
        }

        emitLog('INFO', 'Campaign Name найден, жду основные поля формы', { token, id });

        const mainFieldsReady = await waitFor(() => {
            const hasCampaignName =
                !!document.querySelector('input#campaignName') ||
                !!document.querySelector('input[name="campaignName"]');

            const hasTrackingUrl =
                !!document.querySelector('input#trackingUrl') ||
                !!document.querySelector('input[name="trackingUrl"]') ||
                !!document.querySelector('input#clickUrl') ||
                !!document.querySelector('input[name="clickUrl"]') ||
                !!document.querySelector('textarea[name*="url" i]') ||
                !!document.querySelector('textarea[id*="url" i]') ||
                !!document.querySelector('input[name*="url" i]') ||
                !!document.querySelector('input[id*="url" i]') ||
                !!document.querySelector('input[name*="link" i]') ||
                !!document.querySelector('input[id*="link" i]') ||
                document.body.innerText.includes('Click URL');

            const hasCtr =
                !!document.querySelector('input[name="ctr"]');

            const hasPurpose =
                !!document.querySelector('input[name="campaignPurpose"]') ||
                document.body.innerText.includes('Campaign Purpose');

            const hasDelivery =
                !!document.querySelector('input[name="deliveryType"]') ||
                document.body.innerText.includes('Delivery Type');

            const hasFrequency =
                !!document.querySelector('input[name="frequencyCapping"]') ||
                [...document.querySelectorAll('.form__text-field__name')]
                    .some(el => el.textContent.includes('Frequency capping per Impressions')) ||
                document.body.innerText.includes('Frequency capping per Impressions');

            return hasCampaignName && (hasPurpose || hasDelivery || hasFrequency || hasTrackingUrl || hasCtr);
        }, CONFIG.editorReadyTimeoutMs, 700);

        if (!mainFieldsReady) {
            throw new Error('MAIN_FORM_FIELDS_NOT_READY');
        }

        emitLog('INFO', 'Основные поля формы найдены, жду React-select и мультиселекты', { token, id });

        await waitFor(() => {
            const hasReactSelect =
                !!document.querySelector('[class*="singleValue"]') ||
                !!document.querySelector('[class*="multiValue"]') ||
                !!document.querySelector('.custom-multiselect') ||
                !!document.querySelector('.css-b62m3t-container');

            const hasLabels =
                document.querySelectorAll('.form__text-field__name').length >= 5;

            return hasReactSelect || hasLabels;
        }, 300, 100);

        emitLog('INFO', 'Жду стабилизации DOM перед чтением данных', { token, id });

        await waitForDomStable(900, 250);
        await sleep(150);

        emitLog('INFO', 'Страница стабилизировалась, можно парсить', { token, id });
    }

    /************************************************************
     * PARSER LOGIC
     ************************************************************/

    async function parseCampaignData(token, id) {
        throwIfStopped();

        emitLog('INFO', 'Контрольная проверка перед парсингом: жду заполненные значения', { token, id });

        function findCampaignNameValue() {
            const input =
                document.querySelector('input#campaignName') ||
                document.querySelector('input[name="campaignName"]');

            return input?.value?.trim() || '—';
        }

        function findClickUrlValue() {
            const directInput =
                document.querySelector('input#trackingUrl') ||
                document.querySelector('input[name="trackingUrl"]') ||
                document.querySelector('input#clickUrl') ||
                document.querySelector('input[name="clickUrl"]') ||
                document.querySelector('textarea[name*="url" i]') ||
                document.querySelector('textarea[id*="url" i]') ||
                document.querySelector('input[name*="url" i]') ||
                document.querySelector('input[id*="url" i]') ||
                document.querySelector('input[name*="link" i]') ||
                document.querySelector('input[id*="link" i]');

            if (directInput?.value) {
                return directInput.value.trim();
            }

            const labels = [...document.querySelectorAll('.form__text-field__name, label')];

            const clickUrlLabel = labels.find(el =>
                el.textContent.trim().toLowerCase().includes('click url')
            );

            const group =
                clickUrlLabel?.closest('.form-group') ||
                clickUrlLabel?.parentElement;

            if (group) {
                const input =
                    group.querySelector('input') ||
                    group.querySelector('textarea');

                if (input?.value) {
                    return input.value.trim();
                }

                const link = group.querySelector('a[href]');
                if (link?.href) {
                    return link.href.trim();
                }

                const text = group.textContent.trim();
                const urlMatch = text.match(/https?:\/\/[^\s]+/i);

                if (urlMatch) {
                    return urlMatch[0];
                }
            }

            return '';
        }

        const valuesReady = await waitFor(() => {
            const campaignName = findCampaignNameValue();
            const trackingUrl = findClickUrlValue();
            const ctr = document.querySelector('input[name="ctr"]')?.value;

            const hasAnyRadioChecked =
                !!document.querySelector('input[name="campaignPurpose"]:checked') ||
                !!document.querySelector('input[name="deliveryType"]:checked') ||
                !!document.querySelector('input[name="frequencyCapping"]:checked');

            const hasAnySelectValue =
                !!document.querySelector('[class*="singleValue"]') ||
                !!document.querySelector('[class*="multiValue"]') ||
                !!document.querySelector('.custom-multiselect-label span');

            return Boolean(campaignName && campaignName !== '—') &&
                (Boolean(trackingUrl) || Boolean(ctr) || hasAnyRadioChecked || hasAnySelectValue);
        }, 10000, 500);

        if (!valuesReady) {
            emitLog('WARN', 'Не все значения появились за 10 секунд. Парсю доступные данные.', { token, id });
        }

        await waitForDomStable(900, 250);
        await sleep(150);

        const result = {};

        function getSelectedValuesByLabelText(labelText) {
            const labelElements = document.querySelectorAll('.form__text-field__name');
            let formGroup = null;

            for (const el of labelElements) {
                if (el.textContent.trim() === labelText) {
                    formGroup = el.closest('.form-group');
                    break;
                }
            }

            if (!formGroup) {
                return '—';
            }

            const selectedSpans = formGroup.querySelectorAll('.custom-multiselect .custom-multiselect-label > span');
            const values = Array.from(selectedSpans)
                .map(span => span.textContent.trim())
                .filter(text => text.length > 0);

            return values.length > 0 ? values.join('; ') : '—';
        }

        function getCampaignType(campaignName) {
            if (!campaignName || campaignName === '—') return '—';

            const campaignTypes = [
                'TJ Interstitial',
                'TJ Banner',
                'InPush',
                'Banner',
                'Video',
                'Native',
                'Pops',
                'Push'
            ];

            const lowerCampaignName = campaignName.toLowerCase();

            for (const type of campaignTypes) {
                if (lowerCampaignName.includes(type.toLowerCase())) {
                    return type;
                }
            }

            return '—';
        }

        function getSingleValueNearCheckedRadio(name) {
            const radio = document.querySelector(`input[name="${CSS.escape(name)}"]:checked`);
            const formGroup = radio?.closest('.form-group');

            if (!formGroup) {
                return '';
            }

            const singleValue = formGroup.querySelector(
                '.css-1dimb5e-singleValue, .css-1uccc91-singleValue, [class*="singleValue"]'
            );

            return singleValue?.textContent?.trim() || '';
        }

        function getMultiValuesNearRadio(name) {
            const anyRadio = document.querySelector(`input[name="${CSS.escape(name)}"]`);
            const formGroup = anyRadio?.closest('.form-group');

            if (!formGroup) {
                return '—';
            }

            const tags = formGroup.querySelectorAll(
                '.css-unpojn-multiValue .css-9jq23d, .css-1p3m7a8-multiValue .css-9jq23d, [class*="multiValue"] [class*="label"]'
            );

            const values = Array.from(tags)
                .map(el => el.textContent.trim())
                .filter(Boolean);

            return values.length ? values.join('; ') : '—';
        }

        function getRadioLabelText(radio) {
            if (!radio) return '';

            const label =
                document.querySelector(`label[for="${radio.id}"]`) ||
                radio.closest('label') ||
                radio.nextElementSibling;

            return label?.textContent?.trim() || radio.value || '';
        }

        emitLog('INFO', 'Считываю основные поля', { token, id });

        result.clickUrl = findClickUrlValue();

        result.campaignName = findCampaignNameValue();

        const advertiserNameElement = document.querySelector('[class*="-singleValue"]');
        result.advertiserName = advertiserNameElement?.textContent?.trim() || '—';

        result.CAMPAIGN_TYPE = getCampaignType(result.campaignName);

        const ctrInput = document.querySelector('input[name="ctr"]');
        result.ctr = ctrInput?.value || '';

        const topLevelDomainInput =
            document.querySelector('input#topLevelDomain') ||
            document.querySelector('input[name="topLevelDomain"]');

        result.topLevelDomain = topLevelDomainInput?.value || '';

        emitLog('INFO', 'Считываю Campaign Purpose и Delivery Type', { token, id });

        const campaignPurposeRadio = document.querySelector('input[name="campaignPurpose"]:checked');
        result.campaignPurpose = getRadioLabelText(campaignPurposeRadio);

        const deliveryRadio = document.querySelector('input[name="deliveryType"]:checked');
        result.deliveryType = getRadioLabelText(deliveryRadio);

        emitLog('INFO', 'Считываю Frequency Capping', { token, id });

        const freqCapRadio = document.querySelector('input[name="frequencyCapping"]:checked');
        result.frequencyCapping = getRadioLabelText(freqCapRadio);

        result.interval = '';
        result.frequencyCapValue = '';

        if (freqCapRadio?.value !== 'NO_CAP_ON_IMPRESSION') {
            result.interval = getSingleValueNearCheckedRadio('frequencyCapping');

            const capValueInput = document.querySelector('input[name="frequencyCapValue"]');
            result.frequencyCapValue = capValueInput?.value || '';
        }

        const freqCapImp2Radio = document.querySelector('input[name="frequencyCappingImp2"]:checked');
        result.frequencyCappingImp2 = getRadioLabelText(freqCapImp2Radio);

        result.intervalImp2 = '';
        result.frequencyCapValueImp2 = '';

        if (freqCapImp2Radio?.value !== 'NO_CAP_ON_IMPRESSION') {
            result.intervalImp2 = getSingleValueNearCheckedRadio('frequencyCappingImp2');

            const capValueInput = document.querySelector('input[name="frequencyCapValueImp2"]');
            result.frequencyCapValueImp2 = capValueInput?.value || '';
        }

        const fqCapClickRadio = document.querySelector('input[name="fqCapClick"]:checked');

        if (fqCapClickRadio) {
            result.fqCapClick = getRadioLabelText(fqCapClickRadio);

            result.intervalClick = '';
            result.fqCapClickValue = '';

            if (fqCapClickRadio.value !== 'NO_CAP_ON_IMPRESSION') {
                result.intervalClick = getSingleValueNearCheckedRadio('fqCapClick');

                const capValueInput = document.querySelector('input[name="fqCapClickValue"]');
                result.fqCapClickValue = capValueInput?.value || '';
            }
        } else {
            result.fqCapClick = '';
            result.intervalClick = '';
            result.fqCapClickValue = '';
        }

        emitLog('INFO', 'Считываю гео, языки и устройства', { token, id });

        const countryRadio = document.querySelector('input[name="isIncludeCountry"]:checked');
        result.countryMode = countryRadio ? (countryRadio.value === 'true' ? 'Include' : 'Exclude') : '';
        result.countrySelected = getMultiValuesNearRadio('isIncludeCountry');

        const cityRadio = document.querySelector('input[name="isIncludeCity"]:checked');
        result.cityMode = cityRadio ? (cityRadio.value === 'true' ? 'Include' : 'Exclude') : '';
        result.citySelected = getMultiValuesNearRadio('isIncludeCity');

        const langRadio = document.querySelector('input[name="isIncludeLanguage"]:checked');
        result.languageMode = langRadio ? (langRadio.value === 'true' ? 'Include' : 'Exclude') : '';
        result.languageSelected = getMultiValuesNearRadio('isIncludeLanguage');

        const deviceRadio = document.querySelector('input[name="isIncludeDeviceType"]:checked');
        result.deviceTypeMode = deviceRadio ? (deviceRadio.value === 'true' ? 'Include' : 'Exclude') : '';
        result.deviceTypeSelected = getMultiValuesNearRadio('isIncludeDeviceType');

        emitLog('INFO', 'Считываю аудитории и списки', { token, id });

        const audienceRadio = document.querySelector('input[name="audienceMatchingType"]:checked');
        result.audienceMatchingType = audienceRadio ? audienceRadio.value : '—';

        result.audienceInclude = getSelectedValuesByLabelText('Audience Include');
        result.audienceExclude = getSelectedValuesByLabelText('Audience Exclude');
        result.whitelist = getSelectedValuesByLabelText('White List');
        result.blacklist = getSelectedValuesByLabelText('Black List');

        emitLog('INFO', 'Считываю расписание', { token, id });

        const dayPartingRadio = document.querySelector('input[name="isDayPartingEnable"]:checked');
        result.isDayPartingEnabled = dayPartingRadio ? (dayPartingRadio.value === 'true' ? 'Enable' : 'Disable') : '—';

        result.dayPartingSchedule = '—';

        if (dayPartingRadio?.value === 'true') {
            const schedule = {};
            const rows = document.querySelectorAll('.rt-tr-group');

            rows.forEach(row => {
                const dayCell = row.querySelector('.rt-td:first-child');
                const dayName = dayCell?.textContent?.trim();

                if (dayName && ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].includes(dayName)) {
                    const checkboxes = row.querySelectorAll('.checkbox-control input[type="checkbox"]');
                    const hours = [];

                    checkboxes.forEach((cb, index) => {
                        if (cb.hasAttribute('checked') || cb.checked) {
                            hours.push(String(index).padStart(2, '0'));
                        }
                    });

                    if (hours.length > 0) {
                        schedule[dayName] = hours.join(' ');
                    }
                }
            });

            if (Object.keys(schedule).length > 0) {
                result.dayPartingSchedule = Object.entries(schedule)
                    .map(([day, hours]) => `${day} ${hours}`)
                    .join(' | ');
            } else {
                result.dayPartingSchedule = 'No hours selected';
            }
        }

        const startDateInput = document.querySelector('input[name="scheduling[0].startDate"]');
        const endDateInput = document.querySelector('input[name="scheduling[0].endDate"]');

        result.schedulingStartDate = startDateInput?.value || '—';
        result.schedulingEndDate = endDateInput?.value || '—';

        result.campaign_id = id;
        result.url = `${location.origin}/campaigns/${encodeURIComponent(id)}/edit`;
        result.status = 'OK';

        emitLog('INFO', 'Данные кампании собраны', { token, id });

        return result;
    }

    /************************************************************
     * MANUAL BUTTON MODE
     ************************************************************

    function createManualButton() {
        waitForBody(() => {
            if (document.getElementById('tm-parse-btn')) return;

            const btn = document.createElement('button');

            btn.id = 'tm-parse-btn';
            btn.innerText = 'PARSE CAMPAIGN';

            Object.assign(btn.style, {
                position: 'fixed',
                top: '20px',
                right: '20px',
                zIndex: 999999,
                background: '#0b6f44',
                color: '#fff',
                border: 'none',
                borderRadius: '10px',
                padding: '14px 18px',
                fontSize: '14px',
                fontWeight: '700',
                cursor: 'pointer',
                boxShadow: '0 0 15px rgba(0,0,0,0.4)'
            });

            btn.onclick = async () => {
                if (btn.dataset.running) return;

                btn.dataset.running = '1';
                btn.innerText = 'PARSING...';

                try {
                    GM_setValue(KEYS.stopRequested, false);
                    GM_setValue(KEYS.running, true);

                    const id = getCampaignIdFromUrl(location.href);

                    await waitForCampaignParserReady(null, id);
                    const data = await parseCampaignData(null, id);

                    console.log('PARSED CAMPAIGN DATA:', data);

                    GM_setValue(KEYS.running, false);

                    btn.innerText = 'DONE ✓';
                    btn.style.background = '#0a8f2d';

                    alert('Данные спарсены. Результат выведен в console.log.');

                } catch (e) {
                    console.error(e);

                    GM_setValue(KEYS.running, false);

                    btn.innerText = 'ERROR';
                    btn.style.background = '#000';
                }

                setTimeout(() => {
                    btn.dataset.running = '';
                    btn.innerText = 'PARSE CAMPAIGN';
                    btn.style.background = '#0b6f44';
                }, 3000);
            };

            document.body.appendChild(btn);
        });
    }

    /************************************************************
     * ROUTER
     ************************************************************/

    if (isPanel) {
        runPanel();
        return;
    }

    if (isCampaignEditPage) {
        runWorker();
        return;
    }

})();