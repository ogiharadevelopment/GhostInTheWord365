/* global Office */

// グローバル変数
let currentFormat = null;
let savedFormats = {};
let currentLanguage = 'ja';
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
        keyGuideText: 'SAVE領域でキーを押すと書式を保存、LOAD領域でキーを押すと書式を適用します',
        formatSaved: '書式を保存しました',
        formatApplied: '書式を適用しました',
        formatNotFound: '保存された書式が見つかりません',
        noTextSelected: 'テキストが選択されていません',
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
        keyGuideText: 'Press key in SAVE area to save format, press key in LOAD area to apply format',
        formatSaved: 'Format saved',
        formatApplied: 'Format applied',
        formatNotFound: 'Saved format not found',
        noTextSelected: 'No text selected',
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
        document.addEventListener("DOMContentLoaded", initializeApp);
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
        
        console.log('✅ App initialization completed successfully');
    } catch (error) {
        console.error('❌ App initialization error:', error);
        console.error('Error stack:', error.stack);
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
    
        // SAVE/LOAD領域のイベント
        const saveArea = document.getElementById('save-area');
        const loadArea = document.getElementById('load-area');
        
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
        
        if (loadArea) {
            console.log('✅ Load area found');
            loadArea.addEventListener('mouseenter', (e) => {
                e.preventDefault();
                selectArea('load');
                // フォーカスを確実に取得
                setTimeout(() => {
                    loadArea.focus();
                    loadArea.click();
                }, 10);
            });
            console.log('✅ Load area mouseenter event added');
        } else {
            console.error('❌ Load area not found');
        }
    
        // フォーカスイベント
        if (saveArea) {
            saveArea.addEventListener('focus', () => selectArea('save'));
            console.log('✅ Save area focus event added');
        }
        if (loadArea) {
            loadArea.addEventListener('focus', () => selectArea('load'));
            console.log('✅ Load area focus event added');
        }
        
        // キーボードイベント
        if (saveArea) {
            saveArea.addEventListener('keydown', handleKeyPress);
            console.log('✅ Save area keydown event added');
        }
        if (loadArea) {
            loadArea.addEventListener('keydown', handleKeyPress);
            console.log('✅ Load area keydown event added');
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
        if (loadArea) {
            loadArea.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                loadArea.focus();
            });
            console.log('✅ Load area click event added');
        }
        
        // マウスリーブイベント（フォーカスを維持）
        if (saveArea) {
            saveArea.addEventListener('mouseleave', () => {
                // フォーカスを維持
            });
            console.log('✅ Save area mouseleave event added');
        }
        if (loadArea) {
            loadArea.addEventListener('mouseleave', () => {
                // フォーカスを維持
            });
            console.log('✅ Load area mouseleave event added');
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
    
    document.getElementById('app-title').textContent = t.appTitle;
    document.getElementById('current-format-title').textContent = t.currentFormatTitle;
    document.getElementById('no-selection-text').textContent = t.noSelectionText;
    document.getElementById('save-label').textContent = t.saveLabel;
    document.getElementById('save-instruction').textContent = t.saveInstruction;
    document.getElementById('load-label').textContent = t.loadLabel;
    document.getElementById('load-instruction').textContent = t.loadInstruction;
    document.getElementById('saved-formats-title').textContent = t.savedFormatsTitle;
    document.getElementById('no-saved-formats-text').textContent = t.noSavedFormatsText;
    document.getElementById('key-guide-title').textContent = t.keyGuideTitle;
    document.getElementById('key-guide-text').textContent = t.keyGuideText;
    document.getElementById('lang-ja').textContent = t.japanese;
    document.getElementById('lang-en').textContent = t.english;
}

// 領域の選択
function selectArea(area) {
    selectedArea = area;
    
    // 視覚的フィードバック
    document.querySelectorAll('.action-area').forEach(el => el.classList.remove('selected'));
    document.getElementById(`${area}-area`).classList.add('selected');
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
    const area = event.currentTarget.id.replace('-area', '');
    
    console.log(`Key pressed: ${key} in ${area} area`);
    
    if (area === 'save') {
        saveFormat(key);
    } else if (area === 'load') {
        loadFormat(key);
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

// 書式の適用
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
            
            // テキストが選択されていない場合は、カーソル位置に書式を適用
            if (!selection.text || selection.text.trim() === '') {
                // カーソル位置に書式を適用（新しいテキスト入力用）
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
                
                showMessage(`${key}: ${texts[currentLanguage].formatApplied} (カーソル位置)`, 'success');
                return;
            }
            
            // 選択されたテキストに書式を適用
            const font = selection.font;
            const paragraph = selection.paragraphs.getFirst();
            const format = savedFormats[key];
            
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
            
            showMessage(`${key}: ${texts[currentLanguage].formatApplied}`, 'success');
            
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
            <div class="format-item" data-key="${key}">
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
        
        console.log('Synthetic click event dispatched at position (0,0)');
    } catch (error) {
        console.error('Synthetic click error:', error);
    }
}

// Word APIの可用性チェック
function checkWordAPIAvailability() {
    console.log('=== Word API Availability Check ===');
    
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
    return true;
}

// グローバル関数として公開
window.removeFormat = removeFormat;