(function () {
    'use strict';

    const XLSX_URLS = [
        'vendor/xlsx.full.min.js',
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
    ];

    const STORAGE_KEY = 'outlook_draft_prep_preset_v1';
    const XLSX_MEMORY_RISK_MB = 75;
    const supportsWorker = typeof Worker !== 'undefined';
    const supportsDirectoryPicker = typeof window.showDirectoryPicker === 'function';

    let entrySeq = 1;

    const state = {
        excelEntries: [],
        excelSourceFolderHandle: null,
        presetEntries: [],
        presetFolderHandle: null,
        commonEntries: [],
        commonFolderHandle: null,
        firstHeaders: [],
        headerTokenDefs: [],
        jobs: [],
        showWarningsOnly: false,
        fileMetaCache: new Map()
    };

    const els = {
        excelFolderBtn: document.getElementById('excelFolderBtn'),
        excelFilesBtn: document.getElementById('excelFilesBtn'),
        excelFilesInput: document.getElementById('excelFilesInput'),
        excelStatus: document.getElementById('excelStatus'),
        excelList: document.getElementById('excelList'),
        excelCountBadge: document.getElementById('excelCountBadge'),
        recipientColumn: document.getElementById('recipientColumn'),
        replyToInput: document.getElementById('replyToInput'),

        subjectInput: document.getElementById('subjectInput'),
        toCcInput: document.getElementById('toCcInput'),
        bccInput: document.getElementById('bccInput'),
        bodyInput: document.getElementById('bodyInput'),
        saveTemplateBtn: document.getElementById('saveTemplateBtn'),
        loadTemplateBtn: document.getElementById('loadTemplateBtn'),
        tokenChips: document.getElementById('tokenChips'),
        tokenBadge: document.getElementById('tokenBadge'),

        attachMatchedPreset: document.getElementById('attachMatchedPreset'),
        attachCommon: document.getElementById('attachCommon'),
        attachSourceWorkbook: document.getElementById('attachSourceWorkbook'),
        skipOnMissingMatch: document.getElementById('skipOnMissingMatch'),
        attachmentModeBadge: document.getElementById('attachmentModeBadge'),
        attachmentsHelpText: document.getElementById('attachmentsHelpText'),

        presetFolderBtn: document.getElementById('presetFolderBtn'),
        presetFilesBtn: document.getElementById('presetFilesBtn'),
        presetFilesInput: document.getElementById('presetFilesInput'),
        presetStatus: document.getElementById('presetStatus'),
        presetList: document.getElementById('presetList'),

        commonFolderBtn: document.getElementById('commonFolderBtn'),
        commonFilesBtn: document.getElementById('commonFilesBtn'),
        commonFilesInput: document.getElementById('commonFilesInput'),
        commonStatus: document.getElementById('commonStatus'),
        commonList: document.getElementById('commonList'),

        buildReviewBtn: document.getElementById('buildReviewBtn'),
        generateDraftsBtn: document.getElementById('generateDraftsBtn'),
        warningsOnlyBtn: document.getElementById('warningsOnlyBtn'),
        buildProgressWrap: document.getElementById('buildProgressWrap'),
        buildProgressFill: document.getElementById('buildProgressFill'),
        buildStatus: document.getElementById('buildStatus'),
        reviewRows: document.getElementById('reviewRows'),

        sumTotal: document.getElementById('sumTotal'),
        sumReady: document.getElementById('sumReady'),
        sumSkipped: document.getElementById('sumSkipped'),
        sumWarnings: document.getElementById('sumWarnings')
    };

    const safeStorage = {
        getItem(key) {
            try {
                return window.localStorage.getItem(key);
            } catch (error) {
                return null;
            }
        },
        setItem(key, value) {
            try {
                window.localStorage.setItem(key, value);
            } catch (error) {
                // Ignore restricted storage
            }
        }
    };

    let xlsxReady = null;

    init();

    function init() {
        bindEvents();
        loadPresetIntoForm(false);
        updateAttachmentModeBadge();
        updateAttachmentUiState();
        updateRunCommandBlock();
        renderTokenChips();
        renderEntryList(els.excelList, []);
        renderEntryList(els.presetList, []);
        renderEntryList(els.commonList, []);
        updateReviewSummary();
        updateExportState();
    }

    function bindEvents() {
        els.excelFolderBtn.addEventListener('click', selectExcelFolder);
        els.excelFilesBtn.addEventListener('click', () => els.excelFilesInput.click());
        els.excelFilesInput.addEventListener('change', (e) => selectExcelFilesFromInput(e.target.files));
        els.recipientColumn.addEventListener('change', clearJobsAfterConfigChange);

        els.presetFolderBtn.addEventListener('click', selectPresetFolder);
        els.presetFilesBtn.addEventListener('click', () => els.presetFilesInput.click());
        els.presetFilesInput.addEventListener('change', (e) => selectPresetFilesFromInput(e.target.files));

        els.commonFolderBtn.addEventListener('click', selectCommonFolder);
        els.commonFilesBtn.addEventListener('click', () => els.commonFilesInput.click());
        els.commonFilesInput.addEventListener('change', (e) => selectCommonFilesFromInput(e.target.files));

        [
            els.subjectInput,
            els.toCcInput,
            els.bccInput,
            els.bodyInput,
            els.replyToInput
        ].forEach((el) => {
            el.addEventListener('input', () => {
                persistPresetFromForm();
                clearJobsAfterConfigChange();
            });
        });

        [
            els.attachMatchedPreset,
            els.attachCommon,
            els.attachSourceWorkbook,
            els.skipOnMissingMatch
        ].forEach((el) => {
            el.addEventListener('change', () => {
                updateAttachmentModeBadge();
                updateAttachmentUiState();
                updateRunCommandBlock();
                clearJobsAfterConfigChange();
            });
        });

        els.saveTemplateBtn.addEventListener('click', () => {
            persistPresetFromForm(true);
            setStatus(els.buildStatus, 'Template preset saved in local browser storage.', 'success');
        });

        els.loadTemplateBtn.addEventListener('click', () => {
            const loaded = loadPresetIntoForm(true);
            if (loaded) {
                setStatus(els.buildStatus, 'Template preset loaded.', 'success');
                clearJobsAfterConfigChange();
            } else {
                setStatus(els.buildStatus, 'No saved preset found in this browser.', 'info');
            }
        });

        els.buildReviewBtn.addEventListener('click', buildReviewJobs);
        if (els.generateDraftsBtn) els.generateDraftsBtn.addEventListener('click', exportDraftKitToFolder);
        els.warningsOnlyBtn.addEventListener('click', toggleWarningsOnly);
    }

    async function selectExcelFolder() {
        if (!supportsDirectoryPicker) {
            setStatus(els.excelStatus, 'Folder picker is not supported here. Use "Choose Files".', 'info');
            els.excelFilesInput.click();
            return;
        }
        try {
            const handle = await window.showDirectoryPicker();
            const files = await collectFilesFromDirectory(handle, isXlsxFile);
            state.excelSourceFolderHandle = handle;
            applyExcelEntries(files, `Loaded ${files.length} Excel file(s) from folder "${handle.name}".`);
        } catch (error) {
            if (error && error.name === 'AbortError') return;
            setStatus(els.excelStatus, 'Folder picker failed: ' + (error.message || String(error)), 'error');
        }
    }

    function selectExcelFilesFromInput(fileList) {
        const files = Array.from(fileList || []).filter((f) => isXlsxFile(f.name));
        if (files.length === 0) {
            setStatus(els.excelStatus, 'No .xlsx files were selected.', 'error');
            return;
        }
        state.excelSourceFolderHandle = null;
        const entries = files.map((file) => wrapFile(file));
        applyExcelEntries(entries, `Loaded ${entries.length} Excel file(s) from file picker.`);
    }

    async function applyExcelEntries(entries, statusMessage) {
        state.excelEntries = sortEntries(entries);
        state.jobs = [];
        renderReviewRows();
        updateReviewSummary();
        updateExportState();
        renderEntryList(els.excelList, state.excelEntries);
        els.excelCountBadge.textContent = `${state.excelEntries.length} file${state.excelEntries.length === 1 ? '' : 's'}`;

        if (state.excelEntries.length === 0) {
            els.recipientColumn.disabled = true;
            els.recipientColumn.innerHTML = '<option value="">Load Excel files first</option>';
            state.firstHeaders = [];
            state.headerTokenDefs = [];
            renderTokenChips();
            setStatus(els.excelStatus, 'No Excel files loaded.', 'warning');
            return;
        }

        setStatus(els.excelStatus, statusMessage + ' Reading headers...', 'info');
        try {
            const firstFile = await state.excelEntries[0].getFile();
            const info = await getSheetInfoFromFile(firstFile);
            state.firstHeaders = Array.isArray(info.headers) ? info.headers.slice() : [];
            state.headerTokenDefs = buildHeaderTokenDefs(state.firstHeaders);
            populateRecipientColumnSelect(state.firstHeaders);
            renderTokenChips();
            setStatus(els.excelStatus, `Loaded ${state.excelEntries.length} Excel file(s). ${state.firstHeaders.length} columns detected from first file.`, 'success');
        } catch (error) {
            state.firstHeaders = [];
            state.headerTokenDefs = [];
            populateRecipientColumnSelect([]);
            renderTokenChips();
            setStatus(els.excelStatus, 'Failed to read Excel headers: ' + (error.message || String(error)), 'error');
        }
    }

    async function selectPresetFolder() {
        await selectAttachmentFolder('preset');
    }

    async function selectCommonFolder() {
        await selectAttachmentFolder('common');
    }

    async function selectAttachmentFolder(kind) {
        if (!supportsDirectoryPicker) {
            const msg = 'Folder picker is not supported here. Use "Choose Files" and keep files in a single folder for PowerShell path resolution.';
            setStatus(kind === 'preset' ? els.presetStatus : els.commonStatus, msg, 'info');
            (kind === 'preset' ? els.presetFilesInput : els.commonFilesInput).click();
            return;
        }
        try {
            const handle = await window.showDirectoryPicker();
            const files = await collectFilesFromDirectory(handle, () => true);
            const message = `Loaded ${files.length} attachment file(s) from folder "${handle.name}".`;
            applyAttachmentEntries(kind, files, message, handle);
        } catch (error) {
            if (error && error.name === 'AbortError') return;
            setStatus(kind === 'preset' ? els.presetStatus : els.commonStatus, 'Folder picker failed: ' + (error.message || String(error)), 'error');
        }
    }

    function selectPresetFilesFromInput(fileList) {
        const files = Array.from(fileList || []);
        if (files.length === 0) {
            setStatus(els.presetStatus, 'No attachment files were selected.', 'error');
            return;
        }
        state.presetFolderHandle = null;
        applyAttachmentEntries('preset', files.map((f) => wrapFile(f)), `Loaded ${files.length} preset attachment file(s) from file picker.`);
    }

    function selectCommonFilesFromInput(fileList) {
        const files = Array.from(fileList || []);
        if (files.length === 0) {
            setStatus(els.commonStatus, 'No common attachment files were selected.', 'error');
            return;
        }
        state.commonFolderHandle = null;
        applyAttachmentEntries('common', files.map((f) => wrapFile(f)), `Loaded ${files.length} common attachment file(s) from file picker.`);
    }

    function applyAttachmentEntries(kind, entries, statusMessage, folderHandle) {
        const sorted = sortEntries(entries);
        if (kind === 'preset') {
            state.presetFolderHandle = folderHandle || null;
            state.presetEntries = sorted;
            renderEntryList(els.presetList, sorted);
            setStatus(els.presetStatus, statusMessage, 'success');
        } else {
            state.commonFolderHandle = folderHandle || null;
            state.commonEntries = sorted;
            renderEntryList(els.commonList, sorted);
            setStatus(els.commonStatus, statusMessage, 'success');
        }
        clearJobsAfterConfigChange();
    }

    function clearJobsAfterConfigChange() {
        if (state.jobs.length > 0) {
            state.jobs = [];
            renderReviewRows();
            updateReviewSummary();
            updateExportState();
            setStatus(els.buildStatus, 'Configuration changed. Rebuild the review table before export.', 'info');
        }
    }

    function toggleWarningsOnly() {
        state.showWarningsOnly = !state.showWarningsOnly;
        els.warningsOnlyBtn.textContent = state.showWarningsOnly ? 'Show All Jobs' : 'Show Warnings Only';
        renderReviewRows();
    }

    async function buildReviewJobs() {
        const validation = validateBeforeBuild();
        if (!validation.ok) {
            setStatus(els.buildStatus, validation.message, 'error');
            return;
        }

        const recipientColumnIndex = Number.parseInt(els.recipientColumn.value, 10);
        const ccParse = parseAddressField(els.toCcInput.value);
        const bccParse = parseAddressField(els.bccInput.value);
        const replyToParse = parseAddressField(els.replyToInput.value);
        const invalidParts = [
            ...ccParse.invalid.map((v) => `CC: ${v}`),
            ...bccParse.invalid.map((v) => `BCC: ${v}`),
            ...replyToParse.invalid.map((v) => `Reply-To: ${v}`)
        ];
        if (invalidParts.length > 0) {
            setStatus(els.buildStatus, `Invalid email value(s): ${invalidParts.join(', ')}`, 'error');
            return;
        }

        const subjectTemplate = els.subjectInput.value || '';
        const bodyTemplate = els.bodyInput.value || '';
        const unknownTemplateTokens = uniqueStrings([
            ...findUnknownTokens(subjectTemplate, state.headerTokenDefs),
            ...findUnknownTokens(bodyTemplate, state.headerTokenDefs)
        ]);
        if (unknownTemplateTokens.length > 0) {
            setStatus(els.buildStatus, `Template contains unknown token(s): ${unknownTemplateTokens.map((t) => `{{${t}}}`).join(', ')}`, 'error');
            return;
        }

        els.buildReviewBtn.disabled = true;
        setStatus(els.buildStatus, 'Building review table...', 'info');
        showBuildProgress(0);

        try {
            const presetPool = await buildAttachmentPool(state.presetEntries);
            const commonPool = await buildAttachmentPool(state.commonEntries);
            const jobs = [];

            for (let i = 0; i < state.excelEntries.length; i++) {
                const entry = state.excelEntries[i];
                const file = await entry.getFile();
                const progressBase = i / state.excelEntries.length;
                const progressSpan = 1 / state.excelEntries.length;

                let extracted;
                try {
                    extracted = await extractDraftDataFromExcel(file, recipientColumnIndex, state.firstHeaders.length, (p) => {
                        const overall = Math.round((progressBase + (p * progressSpan)) * 100);
                        showBuildProgress(overall);
                    });
                } catch (error) {
                    jobs.push(makeErrorJob(entry, file, error.message || String(error)));
                    continue;
                }

                const job = buildJobFromExtracted({
                    entry,
                    file,
                    extracted,
                    recipientColumnIndex,
                    ccList: ccParse.valid,
                    bccList: bccParse.valid,
                    replyToList: replyToParse.valid,
                    subjectTemplate,
                    bodyTemplate,
                    presetPool,
                    commonPool
                });

                jobs.push(job);
                showBuildProgress(Math.round(((i + 1) / state.excelEntries.length) * 100));
                await yieldToUI();
            }

            state.jobs = jobs;
            renderReviewRows();
            updateReviewSummary();
            updateExportState();

            const ready = jobs.filter((j) => j.status === 'ready').length;
            const skipped = jobs.filter((j) => j.status !== 'ready').length;
            setStatus(els.buildStatus, `Review built. ${ready} ready draft(s), ${skipped} skipped/error.`, ready > 0 ? 'success' : 'warning');
        } catch (error) {
            setStatus(els.buildStatus, 'Failed to build review: ' + (error.message || String(error)), 'error');
        } finally {
            hideBuildProgress();
            els.buildReviewBtn.disabled = false;
        }
    }

    function validateBeforeBuild() {
        if (state.excelEntries.length === 0) {
            return { ok: false, message: 'Load Excel files first.' };
        }
        if (!els.recipientColumn.value) {
            return { ok: false, message: 'Select the recipient email column.' };
        }
        if (!els.attachMatchedPreset.checked && !els.attachCommon.checked && !els.attachSourceWorkbook.checked) {
            return { ok: false, message: 'Choose at least one attachment strategy.' };
        }
        if (els.attachMatchedPreset.checked && state.presetEntries.length === 0) {
            return { ok: false, message: 'Matched preset attachments are enabled but no preset attachment files were loaded.' };
        }
        if (els.attachCommon.checked && state.commonEntries.length === 0) {
            return { ok: false, message: 'Common attachments are enabled but no common attachment files were loaded.' };
        }
        return { ok: true };
    }

    function makeErrorJob(entry, file, message) {
        return {
            id: `job-${Math.random().toString(36).slice(2, 10)}`,
            sourceExcelFileName: file.name,
            sourceExcelRelativePath: entry.relativePath || file.name,
            sourceExcelBaseName: stripExtension(file.name),
            status: 'error',
            to: [],
            cc: [],
            bcc: [],
            replyTo: [],
            subject: '',
            bodyText: '',
            attachments: [],
            warnings: [],
            errors: [message],
            metrics: {
                recipientCount: 0,
                rowCount: 0
            }
        };
    }

    function buildJobFromExtracted(args) {
        const {
            entry,
            file,
            extracted,
            ccList,
            bccList,
            replyToList,
            subjectTemplate,
            bodyTemplate,
            presetPool,
            commonPool
        } = args;

        const warnings = [];
        const errors = [];
        const baseName = stripExtension(file.name);
        const headerCheck = compareHeaders(state.firstHeaders, extracted.headers || []);
        if (!headerCheck.match) {
            warnings.push('Header schema differs from first file. Column indexes may not align.');
        }

        if (args.recipientColumnIndex >= extracted.totalCols) {
            errors.push('Selected recipient column is missing in this file.');
        }

        const tokenValues = buildTokenValues(file, extracted);
        const subject = renderTemplate(subjectTemplate, tokenValues);
        const bodyText = renderTemplate(bodyTemplate, tokenValues);

        const toList = extracted.emails || [];
        if (toList.length === 0) {
            errors.push('No valid recipient emails found in selected column.');
        } else if (toList.length > 1) {
            warnings.push(`Multiple unique recipients found (${toList.length}).`);
        }

        const attachments = [];

        if (els.attachMatchedPreset.checked) {
            const matched = presetPool.byBaseName.get(baseName.toLowerCase()) || [];
            if (matched.length === 0) {
                const msg = 'No preset attachment matched this Excel file name.';
                if (els.skipOnMissingMatch.checked) {
                    errors.push(msg);
                } else {
                    warnings.push(msg);
                }
            } else {
                matched.forEach((meta) => {
                    attachments.push({
                        kind: 'matchedPreset',
                        fileName: meta.fileName,
                        relativePath: meta.relativePath,
                        expectedBytes: meta.sizeBytes
                    });
                });
            }
        }

        if (els.attachCommon.checked) {
            commonPool.list.forEach((meta) => {
                attachments.push({
                    kind: 'common',
                    fileName: meta.fileName,
                    relativePath: meta.relativePath,
                    expectedBytes: meta.sizeBytes
                });
            });
        }

        if (els.attachSourceWorkbook.checked) {
            attachments.push({
                kind: 'sourceWorkbook',
                fileName: file.name,
                relativePath: entry.relativePath || file.name,
                expectedBytes: file.size
            });
        }

        if (attachments.length === 0) {
            errors.push('No attachments resolved for this draft.');
        }

        return {
            id: `job-${Math.random().toString(36).slice(2, 10)}`,
            sourceExcelFileName: file.name,
            sourceExcelRelativePath: entry.relativePath || file.name,
            sourceExcelBaseName: baseName,
            status: errors.length > 0 ? 'skipped' : 'ready',
            to: toList,
            cc: ccList.slice(),
            bcc: bccList.slice(),
            replyTo: replyToList.slice(),
            subject,
            bodyText,
            attachments,
            warnings,
            errors,
            metrics: {
                recipientCount: toList.length,
                rowCount: Math.max((extracted.totalRows || 0) - 1, 0),
                totalCols: extracted.totalCols || 0
            },
            tokenPreview: {
                filename: tokenValues.filename,
                filename_base: tokenValues.filename_base
            }
        };
    }

    function buildTokenValues(file, extracted) {
        const today = new Date();
        const tokenValues = {
            filename: file.name,
            filename_base: stripExtension(file.name),
            today: formatDateDisplay(today),
            date_iso: today.toISOString().slice(0, 10)
        };

        const firstValues = Array.isArray(extracted.firstValues) ? extracted.firstValues : [];
        state.headerTokenDefs.forEach((def) => {
            const raw = firstValues[def.columnIndex];
            tokenValues[def.token] = raw === undefined || raw === null ? '' : String(raw);
        });

        return tokenValues;
    }

    function renderReviewRows() {
        const rows = state.showWarningsOnly
            ? state.jobs.filter((j) => (j.warnings && j.warnings.length > 0) || (j.errors && j.errors.length > 0))
            : state.jobs.slice();

        if (rows.length === 0) {
            const message = state.jobs.length === 0
                ? 'Build the review table to see draft jobs.'
                : 'No warning rows match the current filter.';
            els.reviewRows.innerHTML = `<tr><td colspan="6" class="empty">${escapeHtml(message)}</td></tr>`;
            return;
        }

        const html = rows.map((job) => {
            const statusClass = job.status === 'ready' ? 'success' : (job.status === 'error' ? 'danger' : 'warn');
            const statusLabel = job.status === 'ready' ? 'Ready' : (job.status === 'error' ? 'Error' : 'Skipped');
            const warningItems = []
                .concat(job.errors || [])
                .concat(job.warnings || [])
                .map((w) => `<li>${escapeHtml(w)}</li>`)
                .join('');

            const toPreview = (job.to || []).slice(0, 3).map((e) => `<span class="pill" title="${escapeHtml(e)}">${escapeHtml(e)}</span>`).join('');
            const moreRecipients = (job.to || []).length > 3 ? `<span class="cell-sub">+${(job.to || []).length - 3} more recipient(s)</span>` : '';

            const attachmentPreview = (job.attachments || []).slice(0, 4)
                .map((a) => `<span class="pill" title="${escapeHtml(a.relativePath || a.fileName)}">${escapeHtml(a.fileName)}</span>`)
                .join('');
            const moreAttachments = (job.attachments || []).length > 4 ? `<span class="cell-sub">+${(job.attachments || []).length - 4} more attachment(s)</span>` : '';

            return `
                <tr>
                    <td><span class="badge ${statusClass}">${statusLabel}</span></td>
                    <td>
                        <span class="cell-title">${escapeHtml(job.sourceExcelFileName || '')}</span>
                        <span class="cell-sub">${escapeHtml(job.sourceExcelRelativePath || '')}</span>
                    </td>
                    <td>
                        <div class="pill-list">${toPreview || '<span class="cell-sub">No recipients</span>'}</div>
                        ${moreRecipients}
                    </td>
                    <td>
                        <span class="cell-title">${escapeHtml(job.subject || '')}</span>
                        <span class="cell-sub">${escapeHtml((job.metrics && job.metrics.rowCount ? `${job.metrics.rowCount} data row(s)` : '0 data rows'))}</span>
                    </td>
                    <td>
                        <div class="pill-list">${attachmentPreview || '<span class="cell-sub">No attachments</span>'}</div>
                        ${moreAttachments}
                    </td>
                    <td>${warningItems ? `<ul class="warning-list">${warningItems}</ul>` : '<span class="cell-sub">None</span>'}</td>
                </tr>
            `;
        }).join('');

        els.reviewRows.innerHTML = html;
    }

    function updateReviewSummary() {
        const total = state.jobs.length;
        const ready = state.jobs.filter((j) => j.status === 'ready').length;
        const skipped = state.jobs.filter((j) => j.status !== 'ready').length;
        const warnings = state.jobs.filter((j) => (j.warnings && j.warnings.length > 0) || (j.errors && j.errors.length > 0)).length;
        els.sumTotal.textContent = String(total);
        els.sumReady.textContent = String(ready);
        els.sumSkipped.textContent = String(skipped);
        els.sumWarnings.textContent = String(warnings);
    }

    function updateExportState() {
        const hasReady = state.jobs.some((j) => j.status === 'ready');
        if (els.generateDraftsBtn) {
            els.generateDraftsBtn.disabled = !hasReady;
        }
    }

    function updateAttachmentModeBadge() {
        const modes = [];
        if (els.attachMatchedPreset.checked) modes.push('matched preset');
        if (els.attachCommon.checked) modes.push('common');
        if (els.attachSourceWorkbook.checked) modes.push('source workbook');
        if (modes.length === 0) {
            setBadge(els.attachmentModeBadge, 'No attachments selected', 'warn');
        } else {
            setBadge(els.attachmentModeBadge, modes.join(' + '), 'neutral');
        }
    }

    function updateAttachmentUiState() {
        const useMatched = !!els.attachMatchedPreset.checked;
        const useCommon = !!els.attachCommon.checked;
        const useSource = !!els.attachSourceWorkbook.checked;

        els.presetFolderBtn.disabled = !useMatched;
        els.presetFilesBtn.disabled = !useMatched;
        els.commonFolderBtn.disabled = !useCommon;
        els.commonFilesBtn.disabled = !useCommon;
        els.skipOnMissingMatch.disabled = !useMatched;

        if (!useMatched && els.skipOnMissingMatch.checked) {
            // Keep the preference value but visually communicate it is inactive.
            // No value change needed because logic already ignores it unless matched mode is on.
        }

        if (useSource && !useMatched && !useCommon) {
            els.attachmentsHelpText.innerHTML = 'You are using <strong>the Excel files from Part 1</strong> as the attachments. No extra attachment folders are needed.';
            return;
        }

        if (!useSource && useMatched && !useCommon) {
            els.attachmentsHelpText.innerHTML = 'You are <strong>not</strong> attaching the Part 1 Excel files. The app will attach files from the <strong>Matched Attachment Pool</strong> instead.';
            return;
        }

        if (!useSource && !useMatched && useCommon) {
            els.attachmentsHelpText.innerHTML = 'Each draft will only include the <strong>Common Attachments</strong> you load here. The Part 1 Excel files will not be attached.';
            return;
        }

        if (!useSource && !useMatched && !useCommon) {
            els.attachmentsHelpText.innerHTML = 'Choose at least one attachment source. Most users should enable <strong>Attach the source Excel workbook itself</strong>.';
            return;
        }

        els.attachmentsHelpText.innerHTML = 'You can combine sources: Part 1 Excel files, matched files, and/or common files. The <strong>skip if missing</strong> rule only affects matched files.';
    }

    function updateRunCommandBlock() {
        if (!els.runCommandBlock) return;
        const parts = [
            'powershell -ExecutionPolicy Bypass -File .\\Create-OutlookDrafts.ps1',
            '-JobsPath .\\outlook-draft-jobs.json'
        ];
        if (els.attachMatchedPreset.checked) {
            parts.push('-PresetAttachmentFolder "C:\\Path\\To\\PresetAttachments"');
        }
        if (els.attachCommon.checked) {
            parts.push('-CommonAttachmentFolder "C:\\Path\\To\\CommonAttachments"');
        }
        if (els.attachSourceWorkbook.checked) {
            parts.push('-SourceWorkbookFolder "C:\\Path\\To\\ExcelFiles"');
        }
        els.runCommandBlock.textContent = parts.join(' ');
    }

    async function buildAttachmentPool(entries) {
        const list = [];
        const byBaseName = new Map();
        for (const entry of entries) {
            const meta = await getEntryMeta(entry);
            list.push(meta);
            const key = stripExtension(meta.fileName).toLowerCase();
            if (!byBaseName.has(key)) byBaseName.set(key, []);
            byBaseName.get(key).push(meta);
        }
        return { list, byBaseName };
    }

    async function getEntryMeta(entry) {
        if (state.fileMetaCache.has(entry.id)) {
            return state.fileMetaCache.get(entry.id);
        }
        const file = await entry.getFile();
        const meta = {
            id: entry.id,
            fileName: entry.name,
            relativePath: entry.relativePath || entry.name,
            sizeBytes: file.size,
            lastModified: file.lastModified || 0
        };
        state.fileMetaCache.set(entry.id, meta);
        return meta;
    }

    async function extractDraftDataFromExcel(file, emailColumnIndex, expectedColCount, onProgress) {
        const memoryRisk = checkXlsxMemoryRisk(file);
        if (!memoryRisk.ok) {
            throw new Error(memoryRisk.message);
        }

        const emails = new Set();
        const chunkSize = 5000;
        const firstValues = new Array(Math.max(expectedColCount, 0)).fill('');

        if (supportsWorker) {
            const worker = new XlsxWorkerClient();
            try {
                await worker.init(file);
                const info = await worker.getInfo();
                const totalRows = info.rowCount || 0;
                const totalCols = info.colCount || 0;

                if (emailColumnIndex >= totalCols) {
                    worker.terminate();
                    return { emails: [], totalRows, totalCols, headers: info.headers || [], firstValues };
                }

                let rowsProcessed = 0;
                for (let startRow = 1; startRow < totalRows; startRow += chunkSize) {
                    const endRow = Math.min(totalRows - 1, startRow + chunkSize - 1);
                    const rows = await worker.getRows(startRow, endRow, totalCols);
                    for (const row of rows) {
                        captureFirstValues(row, firstValues);
                        const parsed = parseEmailList(row[emailColumnIndex]);
                        parsed.forEach((addr) => emails.add(addr));
                    }
                    rowsProcessed += rows.length;
                    if (typeof onProgress === 'function' && totalRows > 1) {
                        onProgress(rowsProcessed / Math.max(totalRows - 1, 1));
                    }
                    if (rowsProcessed % (chunkSize * 2) === 0) {
                        await yieldToUI();
                    }
                }

                worker.terminate();
                return {
                    emails: Array.from(emails),
                    totalRows,
                    totalCols,
                    headers: info.headers || [],
                    firstValues
                };
            } catch (error) {
                worker.terminate();
                // Fall through to main-thread parser
            }
        }

        await ensureXlsx();
        const arrayBuffer = await file.arrayBuffer();
        const workbook = window.XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const info = getSheetInfoFromWorksheet(firstSheet);
        const totalRows = info.rowCount || 0;
        const totalCols = info.colCount || 0;

        if (emailColumnIndex >= totalCols) {
            return { emails: [], totalRows, totalCols, headers: info.headers || [], firstValues };
        }

        let rowsProcessed = 0;
        for (let startRow = 1; startRow < totalRows; startRow += chunkSize) {
            const endRow = Math.min(totalRows - 1, startRow + chunkSize - 1);
            const range = {
                s: { r: startRow, c: 0 },
                e: { r: endRow, c: Math.max(totalCols - 1, 0) }
            };
            const rows = window.XLSX.utils.sheet_to_json(firstSheet, {
                header: 1,
                range,
                defval: '',
                blankrows: true
            });
            for (const row of rows) {
                captureFirstValues(row, firstValues);
                const parsed = parseEmailList(row[emailColumnIndex]);
                parsed.forEach((addr) => emails.add(addr));
            }
            rowsProcessed += rows.length;
            if (typeof onProgress === 'function' && totalRows > 1) {
                onProgress(rowsProcessed / Math.max(totalRows - 1, 1));
            }
            if (rowsProcessed % (chunkSize * 2) === 0) {
                await yieldToUI();
            }
        }

        return {
            emails: Array.from(emails),
            totalRows,
            totalCols,
            headers: info.headers || [],
            firstValues
        };
    }

    function captureFirstValues(row, firstValues) {
        const safeRow = Array.isArray(row) ? row : [];
        const max = Math.min(safeRow.length, firstValues.length);
        for (let i = 0; i < max; i++) {
            if (firstValues[i]) continue;
            const value = safeRow[i];
            if (value === undefined || value === null) continue;
            const text = String(value).trim();
            if (text) firstValues[i] = text;
        }
    }

    async function getSheetInfoFromFile(file) {
        const memoryRisk = checkXlsxMemoryRisk(file);
        if (!memoryRisk.ok) {
            throw new Error(memoryRisk.message);
        }

        if (supportsWorker) {
            const worker = new XlsxWorkerClient();
            try {
                await worker.init(file);
                const info = await worker.getInfo();
                worker.terminate();
                return info;
            } catch (error) {
                worker.terminate();
            }
        }

        await ensureXlsx();
        const arrayBuffer = await file.arrayBuffer();
        const workbook = window.XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        return getSheetInfoFromWorksheet(firstSheet);
    }

    function getSheetInfoFromWorksheet(sheet) {
        const range = window.XLSX.utils.decode_range(sheet['!ref'] || 'A1');
        const rowCount = Math.max(0, range.e.r - range.s.r + 1);
        const colCount = Math.max(0, range.e.c - range.s.c + 1);
        const headerRows = window.XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            range: { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(colCount - 1, 0) } },
            defval: '',
            blankrows: false
        });
        const headerValues = headerRows[0] || [];
        const headers = [];
        for (let i = 0; i < colCount; i++) {
            const raw = headerValues[i];
            headers.push(raw !== undefined && raw !== '' ? String(raw) : `Column ${indexToColumnLetter(i)}`);
        }
        return { rowCount, colCount, headers };
    }

    function populateRecipientColumnSelect(headers) {
        if (!headers || headers.length === 0) {
            els.recipientColumn.disabled = true;
            els.recipientColumn.innerHTML = '<option value="">Load Excel files first</option>';
            return;
        }
        els.recipientColumn.disabled = false;
        const options = ['<option value="">-- Select recipient email column --</option>'];
        headers.forEach((header, index) => {
            options.push(`<option value="${index}">${escapeHtml(header || `Column ${index + 1}`)}</option>`);
        });
        els.recipientColumn.innerHTML = options.join('');
    }

    function buildHeaderTokenDefs(headers) {
        const used = new Set();
        return (headers || []).map((header, index) => {
            const base = slugifyToken(header || `column_${indexToColumnLetter(index)}`);
            let token = base || `column_${index + 1}`;
            let suffix = 2;
            while (used.has(token)) {
                token = `${base || `column_${index + 1}`}_${suffix++}`;
            }
            used.add(token);
            return {
                columnIndex: index,
                header: header || `Column ${index + 1}`,
                token
            };
        });
    }

    function renderTokenChips() {
        const chips = ['filename', 'filename_base', 'today', 'date_iso']
            .concat(state.headerTokenDefs.map((d) => d.token));

        els.tokenChips.innerHTML = chips.map((token) => `<code>{{${escapeHtml(token)}}}</code>`).join('');
        if (state.headerTokenDefs.length > 0) {
            setBadge(els.tokenBadge, `${state.headerTokenDefs.length} header token${state.headerTokenDefs.length === 1 ? '' : 's'}`, 'success');
        } else {
            setBadge(els.tokenBadge, 'Load Excel files', 'neutral');
        }
    }

    function parseAddressField(raw) {
        if (!raw || !String(raw).trim()) {
            return { valid: [], invalid: [] };
        }
        const parts = String(raw)
            .split(/[;,\n\r\t]+/)
            .map((v) => v.trim())
            .filter(Boolean);

        const valid = [];
        const invalid = [];
        for (const part of parts) {
            if (isValidEmail(part)) valid.push(part);
            else invalid.push(part);
        }
        return {
            valid: uniqueStrings(valid),
            invalid: uniqueStrings(invalid)
        };
    }

    function parseEmailList(rawValue) {
        if (rawValue === undefined || rawValue === null) return [];
        const parts = String(rawValue)
            .split(/[;,\n\r\t ]+/)
            .map((v) => v.trim())
            .filter(Boolean);
        const valid = [];
        for (const part of parts) {
            const normalized = part.replace(/^<|>$/g, '').trim();
            if (isValidEmail(normalized)) valid.push(normalized);
        }
        return uniqueStrings(valid);
    }

    function isValidEmail(value) {
        return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
    }

    function findUnknownTokens(template, headerTokenDefs) {
        const text = String(template || '');
        const found = text.match(/\{\{\s*([a-zA-Z0-9_:-]+)\s*\}\}/g) || [];
        if (found.length === 0) return [];
        const known = new Set(['filename', 'filename_base', 'today', 'date_iso']);
        headerTokenDefs.forEach((d) => known.add(d.token));
        const unknown = [];
        for (const match of found) {
            const token = match.replace(/\{\{|\}\}/g, '').trim();
            if (!known.has(token)) unknown.push(token);
        }
        return uniqueStrings(unknown);
    }

    function renderTemplate(template, tokenValues) {
        return String(template || '').replace(/\{\{\s*([a-zA-Z0-9_:-]+)\s*\}\}/g, (full, token) => {
            if (!Object.prototype.hasOwnProperty.call(tokenValues, token)) {
                return full;
            }
            const value = tokenValues[token];
            return value === undefined || value === null ? '' : String(value);
        });
    }

    function compareHeaders(expected, actual) {
        if (!Array.isArray(expected) || !Array.isArray(actual)) return { match: false };
        if (expected.length !== actual.length) return { match: false };
        for (let i = 0; i < expected.length; i++) {
            if (String(expected[i] || '').trim() !== String(actual[i] || '').trim()) {
                return { match: false };
            }
        }
        return { match: true };
    }

    function showBuildProgress(percent) {
        els.buildProgressWrap.hidden = false;
        els.buildProgressFill.style.width = `${percent}%`;
        els.buildProgressFill.textContent = `${percent}%`;
    }

    function hideBuildProgress() {
        els.buildProgressWrap.hidden = true;
        els.buildProgressFill.style.width = '0%';
        els.buildProgressFill.textContent = '0%';
    }

    async function exportDraftKitToFolder() {
        const hasReady = state.jobs.some((j) => j.status === 'ready');
        if (!hasReady) {
            setStatus(els.buildStatus, 'Build the preview first and make sure at least one row is Ready.', 'error');
            return;
        }

        if (!supportsDirectoryPicker) {
            setStatus(els.buildStatus, 'This browser cannot write files to a folder. Use Edge/Chrome for the streamlined Generate flow.', 'warning');
            return;
        }

        const readyCount = state.jobs.filter((j) => j.status === 'ready').length;
        const preferredParent = getPreferredAttachmentSourceFolderHandle();
        const targetMessage = preferredParent
            ? `A new draft-kit folder will be created automatically inside "${preferredParent.name}".`
            : 'You will be asked where to create the draft-kit folder (this happens when files were chosen without a folder handle).';
        const confirmed = window.confirm(
            `Generate ${readyCount} email draft job${readyCount === 1 ? '' : 's'} now?\n\n` +
            targetMessage
        );
        if (!confirmed) {
            setStatus(els.buildStatus, 'Generation cancelled.', 'info');
            return;
        }

        if (els.generateDraftsBtn) els.generateDraftsBtn.disabled = true;
        try {
            const parentFolder = await getDefaultKitParentFolderHandle();
            const kitFolderName = makeDraftKitFolderName();
            const dirHandle = await parentFolder.getDirectoryHandle(kitFolderName, { create: true });
            const jsonText = getJobsJsonText();
            const csvText = getJobsCsvText();

            const [ps1Text, cmdText] = await Promise.all([
                fetchLocalTextAsset('Create-OutlookDrafts.ps1'),
                fetchLocalTextAsset('Create Outlook Drafts.cmd')
            ]);

            await writeTextFileToDirectory(dirHandle, 'outlook-draft-jobs.json', jsonText);
            await writeTextFileToDirectory(dirHandle, 'outlook-draft-jobs-summary.csv', csvText);
            await writeTextFileToDirectory(dirHandle, 'Create-OutlookDrafts.ps1', ps1Text);
            await writeTextFileToDirectory(dirHandle, 'Create Outlook Drafts.cmd', cmdText);

            const parentFolderName = parentFolder && parentFolder.name ? parentFolder.name : 'the folder you selected';
            const generatedFolderName = dirHandle && dirHandle.name ? dirHandle.name : 'the generated draft-kit folder';
            setStatusHtml(
                els.buildStatus,
                `<strong>DONE. NEXT STEPS:</strong><br>` +
                `1. Open the folder you selected in Step 1: <strong>${escapeHtml(parentFolderName)}</strong><br>` +
                `2. Open the new folder: <strong>${escapeHtml(generatedFolderName)}</strong><br>` +
                `3. Double-click <strong>Create Outlook Drafts.cmd</strong><br>` +
                `4. In the popup, click <strong>Test</strong> first, then <strong>Generate</strong>`,
                'success'
            );
        } catch (error) {
            if (error && error.name === 'AbortError') {
                setStatus(els.buildStatus, 'Folder selection cancelled.', 'info');
            } else {
                setStatus(els.buildStatus, 'Could not generate files: ' + (error.message || String(error)), 'error');
            }
        } finally {
            updateExportState();
        }
    }

    async function getDefaultKitParentFolderHandle() {
        const preferred = getPreferredAttachmentSourceFolderHandle();
        if (preferred) {
            return preferred;
        }

        // Fallback only when no reusable folder handle is available (e.g. file picker mode).
        setStatus(els.buildStatus, 'Choose where to create the generated draft kit folder (this appears because no source folder handle is available).', 'info');
        return window.showDirectoryPicker();
    }

    function getPreferredAttachmentSourceFolderHandle() {
        if (els.attachSourceWorkbook.checked && state.excelSourceFolderHandle) {
            return state.excelSourceFolderHandle;
        }
        if (els.attachMatchedPreset.checked && state.presetFolderHandle) {
            return state.presetFolderHandle;
        }
        if (els.attachCommon.checked && state.commonFolderHandle) {
            return state.commonFolderHandle;
        }
        return null;
    }

    function makeDraftKitFolderName() {
        const now = new Date();
        const yyyy = now.getFullYear();
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const dd = String(now.getDate()).padStart(2, '0');
        const hh = String(now.getHours()).padStart(2, '0');
        const mi = String(now.getMinutes()).padStart(2, '0');
        const ss = String(now.getSeconds()).padStart(2, '0');
        return `outlook-draft-kit-${yyyy}${mm}${dd}-${hh}${mi}${ss}`;
    }

    function downloadJobsJson() {
        downloadTextFile('outlook-draft-jobs.json', getJobsJsonText(), 'application/json');
    }

    function downloadJobsCsv() {
        downloadTextFile('outlook-draft-jobs-summary.csv', getJobsCsvText(), 'text/csv;charset=utf-8');
    }

    function getJobsJsonText() {
        const exportObj = buildExportObject();
        return JSON.stringify(exportObj, null, 2);
    }

    function getJobsCsvText() {
        const rows = [[
            'status',
            'excel_file',
            'recipients_count',
            'to',
            'cc',
            'bcc',
            'subject',
            'attachments_count',
            'attachments',
            'warnings',
            'errors'
        ]];

        state.jobs.forEach((job) => {
            rows.push([
                job.status || '',
                job.sourceExcelFileName || '',
                String((job.to || []).length),
                (job.to || []).join('; '),
                (job.cc || []).join('; '),
                (job.bcc || []).join('; '),
                job.subject || '',
                String((job.attachments || []).length),
                (job.attachments || []).map((a) => a.fileName).join('; '),
                (job.warnings || []).join(' | '),
                (job.errors || []).join(' | ')
            ]);
        });

        return rows.map((row) => row.map(csvCell).join(',')).join('\r\n');
    }

    function buildExportObject() {
        const readyJobs = state.jobs.filter((j) => j.status === 'ready');
        const allJobs = state.jobs.map((job) => ({
            id: job.id,
            status: job.status,
            sourceExcelFileName: job.sourceExcelFileName,
            sourceExcelRelativePath: job.sourceExcelRelativePath,
            sourceExcelBaseName: job.sourceExcelBaseName,
            to: job.to,
            cc: job.cc,
            bcc: job.bcc,
            replyTo: job.replyTo,
            subject: job.subject,
            bodyText: job.bodyText,
            attachments: job.attachments,
            warnings: job.warnings,
            errors: job.errors,
            metrics: job.metrics
        }));

        return {
            schemaVersion: '2.0-outlook-draft-jobs',
            generator: {
                app: 'Outlook Draft Prep',
                generatedAt: new Date().toISOString(),
                browserUserAgent: navigator.userAgent
            },
            instructions: {
                note: 'Create drafts with Create-OutlookDrafts.ps1. Provide folder roots because browsers do not expose local file paths.',
                folderRoots: {
                    presetAttachmentFolderRequired: !!els.attachMatchedPreset.checked,
                    commonAttachmentFolderRequired: !!els.attachCommon.checked,
                    sourceWorkbookFolderRequired: !!els.attachSourceWorkbook.checked
                }
            },
            configSnapshot: {
                recipientColumnIndex: Number.parseInt(els.recipientColumn.value || '-1', 10),
                recipientColumnHeader: state.firstHeaders[Number.parseInt(els.recipientColumn.value || '-1', 10)] || null,
                subjectTemplate: els.subjectInput.value || '',
                bodyTemplate: els.bodyInput.value || '',
                cc: parseAddressField(els.toCcInput.value).valid,
                bcc: parseAddressField(els.bccInput.value).valid,
                replyTo: parseAddressField(els.replyToInput.value).valid,
                attachmentStrategy: {
                    matchedPreset: !!els.attachMatchedPreset.checked,
                    common: !!els.attachCommon.checked,
                    sourceWorkbook: !!els.attachSourceWorkbook.checked,
                    skipOnMissingMatch: !!els.skipOnMissingMatch.checked
                },
                headerTokens: state.headerTokenDefs
            },
            batchSummary: {
                totalJobs: allJobs.length,
                readyJobs: readyJobs.length,
                skippedOrErrorJobs: allJobs.length - readyJobs.length
            },
            jobs: allJobs
        };
    }

    function persistPresetFromForm() {
        const payload = {
            subject: els.subjectInput.value || '',
            body: els.bodyInput.value || '',
            cc: els.toCcInput.value || '',
            bcc: els.bccInput.value || '',
            replyTo: els.replyToInput.value || '',
            attachmentFlags: {
                matchedPreset: !!els.attachMatchedPreset.checked,
                common: !!els.attachCommon.checked,
                sourceWorkbook: !!els.attachSourceWorkbook.checked,
                skipOnMissingMatch: !!els.skipOnMissingMatch.checked
            }
        };
        safeStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    }

    function loadPresetIntoForm(applyStatus) {
        const raw = safeStorage.getItem(STORAGE_KEY);
        if (!raw) return false;
        try {
            const data = JSON.parse(raw);
            if (typeof data.subject === 'string') els.subjectInput.value = data.subject;
            if (typeof data.body === 'string') els.bodyInput.value = data.body;
            if (typeof data.cc === 'string') els.toCcInput.value = data.cc;
            if (typeof data.bcc === 'string') els.bccInput.value = data.bcc;
            if (typeof data.replyTo === 'string') els.replyToInput.value = data.replyTo;

            if (data.attachmentFlags && typeof data.attachmentFlags === 'object') {
                if (typeof data.attachmentFlags.matchedPreset === 'boolean') els.attachMatchedPreset.checked = data.attachmentFlags.matchedPreset;
                if (typeof data.attachmentFlags.common === 'boolean') els.attachCommon.checked = data.attachmentFlags.common;
                if (typeof data.attachmentFlags.sourceWorkbook === 'boolean') els.attachSourceWorkbook.checked = data.attachmentFlags.sourceWorkbook;
                if (typeof data.attachmentFlags.skipOnMissingMatch === 'boolean') els.skipOnMissingMatch.checked = data.attachmentFlags.skipOnMissingMatch;
            }
            updateAttachmentModeBadge();
            updateAttachmentUiState();
            updateRunCommandBlock();
            return true;
        } catch (error) {
            if (applyStatus) {
                setStatus(els.buildStatus, 'Saved preset could not be parsed.', 'error');
            }
            return false;
        }
    }

    function renderEntryList(container, entries) {
        if (!entries || entries.length === 0) {
            container.innerHTML = '<div class="list-item meta">No files loaded.</div>';
            return;
        }
        const limit = 200;
        const shown = entries.slice(0, limit);
        const html = shown.map((entry) => {
            const rel = entry.relativePath && entry.relativePath !== entry.name
                ? ` <span class="muted">(${escapeHtml(entry.relativePath)})</span>`
                : '';
            return `<div class="list-item">${escapeHtml(entry.name)}${rel}</div>`;
        });
        if (entries.length > limit) {
            html.push(`<div class="list-item meta">...and ${entries.length - limit} more</div>`);
        }
        container.innerHTML = html.join('');
    }

    async function collectFilesFromDirectory(dirHandle, filterFn) {
        const results = [];
        await walkDirectory(dirHandle, '', results, filterFn);
        return results;
    }

    async function walkDirectory(dirHandle, prefix, results, filterFn) {
        for await (const handle of dirHandle.values()) {
            if (handle.kind === 'file') {
                if (!filterFn(handle.name)) continue;
                const rel = prefix ? `${prefix}/${handle.name}` : handle.name;
                results.push(wrapFileHandle(handle, rel));
            } else if (handle.kind === 'directory') {
                const nextPrefix = prefix ? `${prefix}/${handle.name}` : handle.name;
                await walkDirectory(handle, nextPrefix, results, filterFn);
            }
        }
    }

    function wrapFile(file) {
        const relativePath = normalizeRelativePath(file.webkitRelativePath || file.name);
        return {
            id: `entry-${entrySeq++}`,
            name: file.name,
            relativePath,
            getFile: async () => file
        };
    }

    function wrapFileHandle(fileHandle, relativePath) {
        return {
            id: `entry-${entrySeq++}`,
            name: fileHandle.name,
            relativePath: normalizeRelativePath(relativePath || fileHandle.name),
            getFile: async () => fileHandle.getFile()
        };
    }

    function normalizeRelativePath(path) {
        return String(path || '').replace(/\\/g, '/').replace(/^\/+/, '');
    }

    function sortEntries(entries) {
        return entries.slice().sort((a, b) => {
            const left = (a.relativePath || a.name || '').toLowerCase();
            const right = (b.relativePath || b.name || '').toLowerCase();
            return left.localeCompare(right);
        });
    }

    function isXlsxFile(name) {
        return /\.xlsx$/i.test(String(name || ''));
    }

    function stripExtension(name) {
        return String(name || '').replace(/\.[^.]+$/, '');
    }

    function slugifyToken(value) {
        return String(value || '')
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, '_')
            .replace(/^_+|_+$/g, '')
            .slice(0, 60);
    }

    function setStatus(el, message, type) {
        el.textContent = message;
        el.className = `status show ${type}`;
    }

    function setStatusHtml(el, html, type) {
        el.innerHTML = html;
        el.className = `status show ${type}`;
    }

    function setBadge(el, text, kind) {
        el.textContent = text;
        el.className = `badge ${kind}`;
    }

    function escapeHtml(value) {
        return String(value === undefined || value === null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function csvCell(value) {
        const text = String(value === undefined || value === null ? '' : value);
        return `"${text.replace(/"/g, '""')}"`;
    }

    function uniqueStrings(values) {
        return Array.from(new Set((values || []).filter(Boolean)));
    }

    function indexToColumnLetter(index) {
        let letter = '';
        let i = index;
        while (i >= 0) {
            letter = String.fromCharCode((i % 26) + 65) + letter;
            i = Math.floor(i / 26) - 1;
        }
        return letter;
    }

    function formatDateDisplay(date) {
        const d = date instanceof Date ? date : new Date(date);
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    }

    function checkXlsxMemoryRisk(file) {
        if (!file || !isXlsxFile(file.name)) return { ok: true };
        const sizeMB = file.size / (1024 * 1024);
        if (sizeMB <= XLSX_MEMORY_RISK_MB) return { ok: true };
        return {
            ok: false,
            message: `"${file.name}" is ${sizeMB.toFixed(1)} MB. XLSX files above ${XLSX_MEMORY_RISK_MB} MB have a high risk of browser memory failures.`
        };
    }

    function yieldToUI() {
        return new Promise((resolve) => setTimeout(resolve, 0));
    }

    function downloadTextFile(fileName, text, mimeType) {
        const blob = new Blob([text], { type: mimeType || 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 10000);
    }

    async function fetchLocalTextAsset(path) {
        const response = await fetch(encodeURI(path), { cache: 'no-store' });
        if (!response.ok) {
            throw new Error(`Failed to load ${path} (${response.status})`);
        }
        return response.text();
    }

    async function writeTextFileToDirectory(dirHandle, fileName, text) {
        const fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(text);
        await writable.close();
    }

    function loadScript(url) {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = url;
            script.async = true;
            script.onload = () => resolve(url);
            script.onerror = () => reject(new Error(`Failed to load ${url}`));
            document.head.appendChild(script);
        });
    }

    async function loadDependency(urls, globalName) {
        if (window[globalName]) return true;
        for (const url of urls) {
            try {
                await loadScript(url);
                if (window[globalName]) return true;
            } catch (error) {
                // Try next URL
            }
        }
        throw new Error(`${globalName} failed to load. This network may block CDNs. Add local vendor files under /vendor.`);
    }

    function ensureXlsx() {
        if (!xlsxReady) {
            xlsxReady = loadDependency(XLSX_URLS, 'XLSX');
        }
        return xlsxReady;
    }

    class XlsxWorkerClient {
        constructor() {
            this.worker = new Worker('xlsx-worker.js');
            this.nextId = 1;
            this.pending = new Map();

            this.worker.onmessage = (event) => {
                const { id, ok, data, error } = event.data || {};
                const entry = this.pending.get(id);
                if (!entry) return;
                this.pending.delete(id);
                if (ok) entry.resolve(data);
                else entry.reject(new Error(error || 'Worker error'));
            };

            this.worker.onerror = (event) => {
                this.rejectAll(event.message || 'Worker crashed');
            };

            this.worker.onmessageerror = () => {
                this.rejectAll('Worker message error');
            };
        }

        rejectAll(message) {
            this.pending.forEach((entry) => entry.reject(new Error(message)));
            this.pending.clear();
        }

        request(type, payload, transfer) {
            return new Promise((resolve, reject) => {
                const id = this.nextId++;
                this.pending.set(id, { resolve, reject });
                try {
                    this.worker.postMessage({ id, type, ...payload }, transfer || []);
                } catch (error) {
                    this.pending.delete(id);
                    reject(error);
                }
            });
        }

        async init(file, sheetName) {
            const buffer = await file.arrayBuffer();
            return this.request('init', { buffer, sheetName }, [buffer]);
        }

        getInfo() {
            return this.request('getInfo', {});
        }

        getRows(startRow, endRow, maxCols) {
            return this.request('getRows', { startRow, endRow, maxCols });
        }

        terminate() {
            try {
                this.worker.terminate();
            } catch (error) {
                // Ignore
            }
            this.pending.clear();
        }
    }
})();
