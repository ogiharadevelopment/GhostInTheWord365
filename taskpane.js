/* global Office */

// グローバル変数
let currentFormat = null;
let savedFormats = {};
let currentLanguage = 'ja';
let currentFontSize = 12;
let currentLineSpacing = 1.0;
let isWideMode = true;
let selectedArea = null;

// 多言語対応テキスト
const texts = {
    ja: {
        appTitle: '書式管理',
        currentFormatTitle: '現在の書式',
        noSelectionText: 'テキストを選択してください',
        saveLabel: 'SAVE',
        saveInstruction: 'キーを押して保存',
        loadLabel: 'LOAD',
        loadInstruction: 'キーを押して適用',
        savedFormatsTitle: '保存された書式',
        noSavedFormatsText: '保存された書式はありません',
        keyGuideTitle: 'キーガイド',
        keyGuideText: '保存された書式にマウスオーバーしてキーを押すと書式を適用します',
        fontLabel: 'フォント',
        lineSpacingLabel: '行間',
        formatSaved: '書式を保存しました',
        formatApplied: '書式を適用しました',
        formatNotFound: '保存された書式が見つかりません',
        noTextSelected: 'テキストが選択されていません',
        widthToggle: '幅: 300px',
        widthToggleNarrow: '幅: 100px',
        deleteConfirm: (key) => `書式 "${key}" を削除しますか？`,
        japanese: '日本語',
        english: 'English'
    },
    en: {
        appTitle: 'Format Manager',
        currentFormatTitle: 'Current Format',
        noSelectionText: 'Please select text',
        saveLabel: 'SAVE',
        saveInstruction: 'Press key to save',
        loadLabel: 'LOAD',
        loadInstruction: 'Press key to apply',
        savedFormatsTitle: 'Saved Formats',
        noSavedFormatsText: 'No saved formats',
        keyGuideTitle: 'Key Guide',
        keyGuideText: 'Mouse over a saved format and press a key to apply it',
        fontLabel: 'Font',
        lineSpacingLabel: 'Line Spacing',
        formatSaved: 'Format saved',
        formatApplied: 'Format applied',
        formatNotFound: 'Saved format not found',
        noTextSelected: 'No text selected',
        widthToggle: 'Width: 300px',
        widthToggleNarrow: 'Width: 100px',
        deleteConfirm: (key) => `Delete format "${key}"?`,
        japanese: '日本語',
        english: 'English'
    }
};

// Office.jsの初期化
Office.onReady((info) => {
    console.log('=== Office.onReady called ===');
    console.log('Info object:', JSON.stringify(info, null, 2));
    console.log('Host type:', info.host);
    console.log('Platform:', info.platform);
    
    if (info.host === Office.HostType.Word) {
        console.log('✅ Word host detected - proceeding with initialization');
        console.log('DOM ready state:', document.readyState);
        
        // DOMContentLoadedイベントに依存せず、直接初期化を試行
        if (document.readyState === 'loading') {
            console.log('DOM still loading, waiting for DOMContentLoaded');
            document.addEventListener("DOMContentLoaded", initializeApp);
        } else {
            console.log('DOM already ready, initializing immediately');
            // 少し遅延してから初期化（DOM要素が確実に存在するように）
            setTimeout(initializeApp, 100);
        }
        
        // フォールバック: 3秒後に強制初期化
        setTimeout(() => {
            console.log('Fallback initialization after 3 seconds');
            initializeApp();
        }, 3000);
    } else {
        console.log('❌ Non-Word host detected:', info.host);
        console.log('Expected:', Office.HostType.Word);
    }
}).catch(error => {
    console.error('❌ Office.onReady error:', error);
});

// アプリケーションの初期化
function initializeApp() {
    console.log('=== initializeApp called ===');
    console.log('DOM ready state:', document.readyState);
    console.log('Current time:', new Date().toISOString());
    console.log('Document body exists:', !!document.body);
    console.log('Document head exists:', !!document.head);
    
    // 重複初期化を防ぐ
    if (window.appInitialized) {
        console.log('App already initialized, skipping');
        return;
    }
    window.appInitialized = true;
    
    try {
        console.log('Step 1: Word API availability check');
        // Word APIの可用性チェック
        checkWordAPIAvailability();
        
        console.log('Step 2: Language loading');
        // 言語設定の読み込み
        loadLanguage();
        
        console.log('Step 3: UI update');
        // UIの初期化
        updateUI();
        
        console.log('Step 4: Event listeners setup');
        // イベントリスナーの設定
        setupEventListeners();
        
        // 要素の存在確認
        console.log('=== Element existence check ===');
        const saveArea = document.getElementById('save-area');
        const fontControl = document.getElementById('font-control');
        const lineSpacingControl = document.getElementById('line-spacing-control');
        const langJa = document.getElementById('lang-ja');
        const langEn = document.getElementById('lang-en');
        
        console.log('Save area found:', !!saveArea);
        console.log('Font control found:', !!fontControl);
        console.log('Line spacing control found:', !!lineSpacingControl);
        console.log('Japanese button found:', !!langJa);
        console.log('English button found:', !!langEn);
        
        if (!saveArea || !fontControl || !lineSpacingControl) {
            console.error('❌ Critical elements missing - retrying in 500ms');
            window.appInitialized = false; // リトライのためにフラグをリセット
            setTimeout(initializeApp, 500);
            return;
        }
        
        console.log('Step 5: Saved formats loading');
        // 保存された書式の読み込み
        loadSavedFormats();
        
        console.log('Step 6: Selection change handler');
        // 選択変更の監視
        try {
            Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
            console.log('✅ Selection change handler added');
        } catch (error) {
            console.error('❌ Failed to add selection change handler:', error);
        }
        
        console.log('Step 7: Initial format update');
        // 初期表示
        updateCurrentFormat();
        
        console.log('Step 8: Synthetic click setup');
        // 疑似クリックイベントの設定
        setupSyntheticClick();
        
        console.log('Step 9: Final UI update');
        // 最終的なUI更新
        updateSavedFormatsList();
        
        console.log('Step 10: Initialize display values');
        // 初期表示値を設定
        updateFontSizeDisplay();
        updateLineSpacingDisplay();
        
        console.log('✅ App initialization completed successfully');
        console.log('=== Initialization Summary ===');
        console.log('All steps completed without errors');
        console.log('Ready for user interaction');
    } catch (error) {
        console.error('❌ App initialization error:', error);
        console.error('Error stack:', error.stack);
        window.appInitialized = false; // エラー時はフラグをリセット
    }
}

// イベントリスナーの設定
function setupEventListeners() {
    console.log('=== setupEventListeners called ===');
    
    try {
        // 言語切り替え
        const langJa = document.getElementById('lang-ja');
        const langEn = document.getElementById('lang-en');
        
        if (langJa) {
            langJa.addEventListener('click', () => setLanguage('ja'));
            console.log('✅ Japanese language button event added');
        } else {
            console.error('❌ Japanese language button not found');
        }
        
        if (langEn) {
            langEn.addEventListener('click', () => setLanguage('en'));
            console.log('✅ English language button event added');
        } else {
            console.error('❌ English language button not found');
        }
    
        // コントロール領域のイベント
        const saveArea = document.getElementById('save-area');
        const fontControl = document.getElementById('font-control');
        const lineSpacingControl = document.getElementById('line-spacing-control');
        const widthToggle = document.getElementById('width-toggle');
        
        if (saveArea) {
            console.log('✅ Save area found');
            // マウスイベント
            saveArea.addEventListener('mouseenter', (e) => {
                e.preventDefault();
                selectArea('save');
                // フォーカスを確実に取得
                setTimeout(() => {
                    saveArea.focus();
                    saveArea.click();
                }, 10);
            });
            console.log('✅ Save area mouseenter event added');
        } else {
            console.error('❌ Save area not found');
        }
        
        if (fontControl) {
            console.log('✅ Font control found');
            fontControl.addEventListener('mouseenter', (e) => {
                e.preventDefault();
                selectArea('font');
                setTimeout(() => {
                    fontControl.focus();
                    fontControl.click();
                }, 10);
            });
            fontControl.addEventListener('wheel', handleFontWheel);
            console.log('✅ Font control events added');
        } else {
            console.error('❌ Font control not found');
        }
        
        if (lineSpacingControl) {
            console.log('✅ Line spacing control found');
            lineSpacingControl.addEventListener('mouseenter', (e) => {
                e.preventDefault();
                selectArea('lineSpacing');
                setTimeout(() => {
                    lineSpacingControl.focus();
                    lineSpacingControl.click();
                }, 10);
            });
            lineSpacingControl.addEventListener('wheel', handleLineSpacingWheel);
            console.log('✅ Line spacing control events added');
        } else {
            console.error('❌ Line spacing control not found');
        }
        
        if (widthToggle) {
            console.log('✅ Width toggle found');
            widthToggle.addEventListener('click', toggleWidth);
            console.log('✅ Width toggle event added');
        } else {
            console.error('❌ Width toggle not found');
        }
    
        // フォーカスイベント
        if (saveArea) {
            saveArea.addEventListener('focus', () => selectArea('save'));
            console.log('✅ Save area focus event added');
        }
        if (fontControl) {
            fontControl.addEventListener('focus', () => selectArea('font'));
            console.log('✅ Font control focus event added');
        }
        if (lineSpacingControl) {
            lineSpacingControl.addEventListener('focus', () => selectArea('lineSpacing'));
            console.log('✅ Line spacing control focus event added');
        }
        
        // キーボードイベント
        if (saveArea) {
            saveArea.addEventListener('keydown', handleKeyPress);
            console.log('✅ Save area keydown event added');
        }
        if (fontControl) {
            fontControl.addEventListener('keydown', handleKeyPress);
            console.log('✅ Font control keydown event added');
        }
        if (lineSpacingControl) {
            lineSpacingControl.addEventListener('keydown', handleKeyPress);
            console.log('✅ Line spacing control keydown event added');
        }
        
        // クリックイベント（フォーカス用）
        if (saveArea) {
            saveArea.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                saveArea.focus();
            });
            console.log('✅ Save area click event added');
        }
        if (fontControl) {
            fontControl.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                fontControl.focus();
            });
            console.log('✅ Font control click event added');
        }
        if (lineSpacingControl) {
            lineSpacingControl.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                lineSpacingControl.focus();
            });
            console.log('✅ Line spacing control click event added');
        }
        
        // マウスリーブイベント（フォーカスを維持）
        if (saveArea) {
            saveArea.addEventListener('mouseleave', () => {
                // フォーカスを維持
            });
            console.log('✅ Save area mouseleave event added');
        }
        if (fontControl) {
            fontControl.addEventListener('mouseleave', () => {
                // フォーカスを維持
            });
            console.log('✅ Font control mouseleave event added');
        }
        if (lineSpacingControl) {
            lineSpacingControl.addEventListener('mouseleave', () => {
                // フォーカスを維持
            });
            console.log('✅ Line spacing control mouseleave event added');
        }
        
        console.log('✅ setupEventListeners completed successfully');
    } catch (error) {
        console.error('❌ setupEventListeners error:', error);
        console.error('Error stack:', error.stack);
    }
}

// 言語設定
function setLanguage(lang) {
    currentLanguage = lang;
    localStorage.setItem('formatManagerLanguage', lang);
    updateUI();
    
    // 言語ボタンの状態更新
    document.querySelectorAll('.lang-btn').forEach(btn => btn.classList.remove('active'));
    document.getElementById(`lang-${lang}`).classList.add('active');
}

// 言語設定の読み込み
function loadLanguage() {
    const savedLang = localStorage.getItem('formatManagerLanguage');
    if (savedLang && texts[savedLang]) {
        currentLanguage = savedLang;
    }
}

// UIの更新
function updateUI() {
    const t = texts[currentLanguage];
    
    // 要素の存在確認をしてから更新
    const elements = {
        'app-title': t.appTitle,
        'current-format-title': t.currentFormatTitle,
        'no-selection-text': t.noSelectionText,
        'save-label': t.saveLabel,
        'save-instruction': t.saveInstruction,
        'saved-formats-title': t.savedFormatsTitle,
        'no-saved-formats-text': t.noSavedFormatsText,
        'key-guide-title': t.keyGuideTitle,
        'key-guide-text': t.keyGuideText,
        'font-label': t.fontLabel,
        'line-spacing-label': t.lineSpacingLabel,
        'width-toggle': t.widthToggle,
        'lang-ja': t.japanese,
        'lang-en': t.english
    };
    
    for (const [id, text] of Object.entries(elements)) {
        const element = document.getElementById(id);
        if (element) {
            element.textContent = text;
        } else {
            console.warn(`Element with id '${id}' not found`);
        }
    }
}

// 領域の選択
function selectArea(area) {
    selectedArea = area;
    
    // 視覚的フィードバック
    document.querySelectorAll('.action-area, .control-area').forEach(el => el.classList.remove('selected'));
    
    // 対応する要素にクラスを追加
    if (area === 'save') {
        const saveArea = document.getElementById('save-area');
        if (saveArea) saveArea.classList.add('selected');
    } else if (area === 'font') {
        const fontControl = document.getElementById('font-control');
        if (fontControl) fontControl.classList.add('selected');
    } else if (area === 'lineSpacing') {
        const lineSpacingControl = document.getElementById('line-spacing-control');
        if (lineSpacingControl) lineSpacingControl.classList.add('selected');
    }
}

// キー押下の処理
function handleKeyPress(event) {
    // 特殊キーは無視
    if (event.key === 'Tab' || event.key === 'Shift' || event.key === 'Control' ||
        event.key === 'Alt' || event.key === 'Meta' || event.key === 'CapsLock' ||
        event.key === 'Enter' || event.key === 'Escape' || event.key === 'ArrowUp' ||
        event.key === 'ArrowDown' || event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
        return;
    }
    
    event.preventDefault();
    event.stopPropagation();
    
    const key = event.key.toLowerCase();
    const targetId = event.currentTarget.id;
    
    console.log(`Key pressed: ${key} in ${targetId}`);
    
    if (targetId === 'save-area') {
        saveFormat(key);
    } else if (targetId === 'font-control') {
        adjustFontSize(key);
    } else if (targetId === 'line-spacing-control') {
        adjustLineSpacing(key);
    }
    
    // 視覚的フィードバック
    event.currentTarget.classList.add('pulse');
    setTimeout(() => {
        event.currentTarget.classList.remove('pulse');
    }, 300);
}

// 書式の保存
function saveFormat(key) {
    if (!currentFormat) {
        showMessage(texts[currentLanguage].noTextSelected, 'error');
        return;
    }
    
    try {
        savedFormats[key] = {
            ...currentFormat,
            timestamp: new Date().toISOString()
        };
        
        localStorage.setItem('savedFormats', JSON.stringify(savedFormats));
        updateSavedFormatsList();
        
        // 視覚的フィードバック
        const saveArea = document.getElementById('save-area');
        saveArea.classList.add('saved');
        setTimeout(() => saveArea.classList.remove('saved'), 1000);
        
        showMessage(`${key}: ${texts[currentLanguage].formatSaved}`, 'success');
        
    } catch (error) {
        console.error('書式保存エラー:', error);
        showMessage('書式の保存に失敗しました', 'error');
    }
}

// 書式の適用（保存された書式から）
function loadFormat(key) {
    if (!savedFormats[key]) {
        showMessage(texts[currentLanguage].formatNotFound, 'error');
        return;
    }
    
    Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            
            // 選択範囲を確認
            selection.load('text');
            await context.sync();
            
            // テキストが選択されていない場合でも適用可能
            const format = savedFormats[key];
            const font = selection.font;
            const paragraph = selection.paragraphs.getFirst();
            
            font.name = format.font.name;
            font.size = format.font.size;
            font.bold = format.font.bold;
            font.italic = format.font.italic;
            font.color = format.font.color;
            font.underline = format.font.underline;
            font.highlightColor = format.font.highlightColor;
            
            paragraph.alignment = format.paragraph.alignment;
            paragraph.leftIndent = format.paragraph.leftIndent;
            paragraph.rightIndent = format.paragraph.rightIndent;
            paragraph.lineSpacing = format.paragraph.lineSpacing;
            paragraph.spaceAfter = format.paragraph.spaceAfter;
            paragraph.spaceBefore = format.paragraph.spaceBefore;
            
            await context.sync();
            
            const message = selection.text && selection.text.trim() !== '' 
                ? `${key}: ${texts[currentLanguage].formatApplied}`
                : `${key}: ${texts[currentLanguage].formatApplied} (カーソル位置)`;
            showMessage(message, 'success');
            
        } catch (error) {
            console.error('書式適用エラー:', error);
            showMessage('書式の適用に失敗しました', 'error');
        }
    }).catch(error => {
        console.error('Word.run エラー:', error);
        showMessage('書式の適用に失敗しました', 'error');
    });
}

// 選択変更時の処理
function onSelectionChanged() {
    console.log('Selection changed');
    try {
        updateCurrentFormat();
    } catch (error) {
        console.error('Selection change error:', error);
    }
}

// 現在の書式を更新
function updateCurrentFormat() {
    console.log('updateCurrentFormat called');
    
    if (typeof Word === 'undefined') {
        console.error('Word API not available');
        return;
    }
    
    Word.run(async (context) => {
        try {
            console.log('Word.run started');
            const selection = context.document.getSelection();
            
            // 選択範囲を確認
            selection.load('text');
            await context.sync();
            
            console.log('Selected text:', selection.text);
            
            // テキストが選択されているかチェック
            if (!selection.text || selection.text.trim() === '') {
                console.log('No text selected');
                currentFormat = null;
                displayCurrentFormat(null);
                return;
            }
            
            const font = selection.font;
            const paragraph = selection.paragraphs.getFirst();
            
            // 書式情報を読み込み
            font.load('name, size, bold, italic, color, underline, highlightColor');
            paragraph.load('alignment, leftIndent, rightIndent, lineSpacing, spaceAfter, spaceBefore');
            
            await context.sync();
            
            console.log('Font info:', {
                name: font.name,
                size: font.size,
                bold: font.bold,
                italic: font.italic,
                color: font.color
            });
            
            // 書式情報を取得
            currentFormat = {
                font: {
                    name: font.name,
                    size: font.size,
                    bold: font.bold,
                    italic: font.italic,
                    color: font.color,
                    underline: font.underline,
                    highlightColor: font.highlightColor
                },
                paragraph: {
                    alignment: paragraph.alignment,
                    leftIndent: paragraph.leftIndent,
                    rightIndent: paragraph.rightIndent,
                    lineSpacing: paragraph.lineSpacing,
                    spaceAfter: paragraph.spaceAfter,
                    spaceBefore: paragraph.spaceBefore
                }
            };
            
            // 現在のフォントサイズと行間を更新
            currentFontSize = font.size;
            currentLineSpacing = paragraph.lineSpacing;
            
            // 表示を更新
            updateFontSizeDisplay();
            updateLineSpacingDisplay();
            
            // 現在の書式を表示
            displayCurrentFormat(currentFormat);
            console.log('Format updated successfully');
            
        } catch (error) {
            console.error('書式取得エラー:', error);
            currentFormat = null;
            displayCurrentFormat(null);
        }
    }).catch(error => {
        console.error('Word.run エラー:', error);
        currentFormat = null;
        displayCurrentFormat(null);
    });
}

// 現在の書式を表示
function displayCurrentFormat(format) {
    const formatDisplay = document.getElementById('current-format-display');
    
    if (!format) {
        formatDisplay.innerHTML = `<p>${texts[currentLanguage].noSelectionText}</p>`;
        return;
    }
    
    const font = format.font;
    const paragraph = format.paragraph;
    
    // 配置の日本語表示
    const alignmentText = getAlignmentText(paragraph.alignment);
    
    const formatText = `
        <div class="format-info">
            <strong>${font.name}</strong> ${font.size}px<br>
            ${font.bold ? '太字' : ''} ${font.italic ? '斜体' : ''}<br>
            ${alignmentText} | 色: ${font.color}
        </div>
    `;
    
    formatDisplay.innerHTML = formatText;
}

// 配置の日本語表示を取得
function getAlignmentText(alignment) {
    const alignments = {
        'Left': currentLanguage === 'ja' ? '左揃え' : 'Left',
        'Center': currentLanguage === 'ja' ? '中央揃え' : 'Center',
        'Right': currentLanguage === 'ja' ? '右揃え' : 'Right',
        'Justified': currentLanguage === 'ja' ? '両端揃え' : 'Justified'
    };
    return alignments[alignment] || alignment;
}

// 保存された書式を読み込み
function loadSavedFormats() {
    try {
        const saved = localStorage.getItem('savedFormats');
        if (saved) {
            savedFormats = JSON.parse(saved);
            updateSavedFormatsList();
        }
    } catch (error) {
        console.error('保存された書式の読み込みエラー:', error);
    }
}

// 保存された書式一覧を更新
function updateSavedFormatsList() {
    const savedFormatsList = document.getElementById('saved-formats-list');
    
    if (Object.keys(savedFormats).length === 0) {
        savedFormatsList.innerHTML = `<p>${texts[currentLanguage].noSavedFormatsText}</p>`;
        return;
    }
    
    let html = '';
    for (const [key, format] of Object.entries(savedFormats)) {
        const date = new Date(format.timestamp).toLocaleDateString();
        html += `
            <div class="format-item" data-key="${key}" tabindex="0">
                <div>
                    <div class="format-key">${key}</div>
                    <div class="format-preview">${format.font.name} ${format.font.size}px - ${getAlignmentText(format.paragraph.alignment)} (${date})</div>
                </div>
                <button class="format-remove" data-key="${key}">×</button>
            </div>
        `;
    }
    
    savedFormatsList.innerHTML = html;
    
    // 削除ボタンのイベントリスナーを追加
    const removeButtons = savedFormatsList.querySelectorAll('.format-remove');
    removeButtons.forEach(button => {
        button.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const key = button.dataset.key;
            removeFormat(key);
        });
    });
    
    // 書式項目のイベントリスナーを追加
    const formatItems = savedFormatsList.querySelectorAll('.format-item');
    formatItems.forEach(item => {
        item.addEventListener('mouseenter', (e) => {
            e.preventDefault();
            item.classList.add('focused');
            item.focus();
        });
        
        item.addEventListener('mouseleave', (e) => {
            item.classList.remove('focused');
        });
        
        item.addEventListener('keydown', (e) => {
            if (e.key !== 'Tab' && e.key !== 'Shift' && e.key !== 'Control' && 
                e.key !== 'Alt' && e.key !== 'Meta' && e.key !== 'CapsLock' &&
                e.key !== 'Enter' && e.key !== 'Escape' && e.key !== 'ArrowUp' &&
                e.key !== 'ArrowDown' && e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') {
                e.preventDefault();
                e.stopPropagation();
                const key = e.key.toLowerCase();
                loadFormat(key);
            }
        });
    });
}

// 書式の削除
function removeFormat(key) {
    if (confirm(`書式 "${key}" を削除しますか？`)) {
        delete savedFormats[key];
        localStorage.setItem('savedFormats', JSON.stringify(savedFormats));
        updateSavedFormatsList();
        showMessage(`書式 "${key}" を削除しました`, 'success');
    }
}

// メッセージを表示
function showMessage(message, type) {
    // 既存のメッセージを削除
    const existingMessage = document.querySelector('.status-message');
    if (existingMessage) {
        existingMessage.remove();
    }
    
    // 新しいメッセージを作成
    const messageDiv = document.createElement('div');
    messageDiv.className = `status-message status-${type}`;
    messageDiv.textContent = message;
    
    // メッセージを表示
    document.body.appendChild(messageDiv);
    
    // 3秒後にメッセージを削除
    setTimeout(() => {
        if (messageDiv.parentNode) {
            messageDiv.remove();
        }
    }, 3000);
}

// 疑似クリックイベントの設定
function setupSyntheticClick() {
    console.log('=== setupSyntheticClick called ===');
    try {
        // 位置0,0での疑似クリックイベントを作成
        const syntheticClickEvent = new MouseEvent('click', {
            bubbles: true,
            cancelable: true,
            view: window,
            clientX: 0,
            clientY: 0,
            screenX: 0,
            screenY: 0,
            button: 0,
            buttons: 1,
            ctrlKey: false,
            shiftKey: false,
            altKey: false,
            metaKey: false
        });
        
        // 疑似クリックイベントを発火
        document.dispatchEvent(syntheticClickEvent);
        
        console.log('✅ Synthetic click event dispatched at position (0,0)');
        console.log('Event details:', {
            type: syntheticClickEvent.type,
            bubbles: syntheticClickEvent.bubbles,
            cancelable: syntheticClickEvent.cancelable
        });
    } catch (error) {
        console.error('❌ Synthetic click error:', error);
        console.error('Error stack:', error.stack);
    }
}

// Word APIの可用性チェック
function checkWordAPIAvailability() {
    console.log('=== Word API Availability Check ===');
    console.log('Check started at:', new Date().toISOString());
    
    // 1. Office.jsの読み込み確認
    if (typeof Office === 'undefined') {
        console.error('❌ Office.js is not loaded');
        showMessage('Office.jsが読み込まれていません', 'error');
        return false;
    }
    console.log('✅ Office.js is loaded');
    
    // 2. Office.contextの確認
    if (!Office.context) {
        console.error('❌ Office.context is not available');
        showMessage('Office.contextが利用できません', 'error');
        return false;
    }
    console.log('✅ Office.context is available');
    
    // 3. Word APIの確認
    if (typeof Word === 'undefined') {
        console.error('❌ Word API is not available');
        showMessage('Word APIが利用できません', 'error');
        return false;
    }
    console.log('✅ Word API is available');
    
    // 4. Office.context.documentの確認
    if (!Office.context.document) {
        console.error('❌ Office.context.document is not available');
        showMessage('Office.context.documentが利用できません', 'error');
        return false;
    }
    console.log('✅ Office.context.document is available');
    
    // 5. ホストアプリケーションの確認
    console.log('Host application:', Office.context.host);
    if (Office.context.host !== Office.HostType.Word) {
        console.warn('⚠️ Not running in Word host:', Office.context.host);
        showMessage('Word以外のアプリケーションで実行されています', 'error');
        return false;
    }
    console.log('✅ Running in Word host');
    
    // 6. プラットフォーム情報の確認
    console.log('Platform:', Office.context.platform);
    console.log('Office version:', Office.context.requirements);
    
    // 7. 基本的なWord API機能のテスト
    try {
        Word.run(async (context) => {
            const document = context.document;
            document.load('body');
            await context.sync();
            console.log('✅ Basic Word API test passed');
            console.log('Document body length:', document.body.text ? document.body.text.length : 0);
        }).catch(error => {
            console.error('❌ Basic Word API test failed:', error);
            showMessage('Word APIの基本テストに失敗しました', 'error');
        });
    } catch (error) {
        console.error('❌ Word API test error:', error);
        showMessage('Word APIテストでエラーが発生しました', 'error');
    }
    
    console.log('=== Word API Availability Check Complete ===');
    console.log('Check completed at:', new Date().toISOString());
    return true;
}

// フォントサイズ調整
function adjustFontSize(key) {
    const step = 1;
    if (key === '+' || key === '=') {
        currentFontSize += step;
    } else if (key === '-') {
        currentFontSize = Math.max(1, currentFontSize - step);
    } else {
        return;
    }
    
    updateFontSizeDisplay();
    applyCurrentFormat();
}

// 行間調整
function adjustLineSpacing(key) {
    const step = 0.1;
    if (key === '+' || key === '=') {
        currentLineSpacing += step;
    } else if (key === '-') {
        currentLineSpacing = Math.max(0.1, currentLineSpacing - step);
    } else {
        return;
    }
    
    updateLineSpacingDisplay();
    applyCurrentFormat();
}

// フォントサイズ表示更新
function updateFontSizeDisplay() {
    const display = document.getElementById('font-size-display');
    if (display) {
        display.textContent = `${currentFontSize}px`;
    }
}

// 行間表示更新
function updateLineSpacingDisplay() {
    const display = document.getElementById('line-spacing-display');
    if (display) {
        display.textContent = currentLineSpacing.toFixed(1);
    }
}

// 現在の書式を適用
function applyCurrentFormat() {
    if (!currentFormat) return;
    
    Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            const font = selection.font;
            const paragraph = selection.paragraphs.getFirst();
            
            // 現在の書式を更新
            currentFormat.font.size = currentFontSize;
            currentFormat.paragraph.lineSpacing = currentLineSpacing;
            
            // 書式を適用
            font.name = currentFormat.font.name;
            font.size = currentFormat.font.size;
            font.bold = currentFormat.font.bold;
            font.italic = currentFormat.font.italic;
            font.color = currentFormat.font.color;
            font.underline = currentFormat.font.underline;
            font.highlightColor = currentFormat.font.highlightColor;
            
            paragraph.alignment = currentFormat.paragraph.alignment;
            paragraph.leftIndent = currentFormat.paragraph.leftIndent;
            paragraph.rightIndent = currentFormat.paragraph.rightIndent;
            paragraph.lineSpacing = currentFormat.paragraph.lineSpacing;
            paragraph.spaceAfter = currentFormat.paragraph.spaceAfter;
            paragraph.spaceBefore = currentFormat.paragraph.spaceBefore;
            
            await context.sync();
            
        } catch (error) {
            console.error('書式適用エラー:', error);
        }
    }).catch(error => {
        console.error('Word.run エラー:', error);
    });
}

// ホイールイベント処理
function handleFontWheel(event) {
    event.preventDefault();
    const delta = event.deltaY > 0 ? -1 : 1;
    currentFontSize = Math.max(1, currentFontSize + delta);
    updateFontSizeDisplay();
    applyCurrentFormat();
}

function handleLineSpacingWheel(event) {
    event.preventDefault();
    const delta = event.deltaY > 0 ? -0.1 : 0.1;
    currentLineSpacing = Math.max(0.1, currentLineSpacing + delta);
    updateLineSpacingDisplay();
    applyCurrentFormat();
}

// 幅切り替え
function toggleWidth() {
    isWideMode = !isWideMode;
    const app = document.getElementById('app');
    const button = document.getElementById('width-toggle');
    
    if (isWideMode) {
        app.classList.remove('narrow');
        app.classList.add('wide');
        button.textContent = '幅: 300px';
    } else {
        app.classList.remove('wide');
        app.classList.add('narrow');
        button.textContent = '幅: 100px';
    }
}

// グローバル関数として公開
window.removeFormat = removeFormat;

// デバッグ用: 手動初期化
window.manualInit = function() {
    console.log('Manual initialization triggered');
    window.appInitialized = false;
    initializeApp();
};