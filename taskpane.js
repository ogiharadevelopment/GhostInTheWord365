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
    if (info.host === Office.HostType.Word) {
        document.addEventListener("DOMContentLoaded", initializeApp);
    }
});

// アプリケーションの初期化
function initializeApp() {
    // 言語設定の読み込み
    loadLanguage();
    
    // UIの初期化
    updateUI();
    
    // イベントリスナーの設定
    setupEventListeners();
    
    // 保存された書式の読み込み
    loadSavedFormats();
    
    // 選択変更の監視
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    
    // 初期表示
    updateCurrentFormat();
    
    // 疑似クリックイベントの設定
    setupSyntheticClick();
}

// イベントリスナーの設定
function setupEventListeners() {
    // 言語切り替え
    document.getElementById('lang-ja').addEventListener('click', () => setLanguage('ja'));
    document.getElementById('lang-en').addEventListener('click', () => setLanguage('en'));
    
    // SAVE/LOAD領域のイベント
    const saveArea = document.getElementById('save-area');
    const loadArea = document.getElementById('load-area');
    
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
    loadArea.addEventListener('mouseenter', (e) => {
        e.preventDefault();
        selectArea('load');
        // フォーカスを確実に取得
        setTimeout(() => {
            loadArea.focus();
            loadArea.click();
        }, 10);
    });
    
    // フォーカスイベント
    saveArea.addEventListener('focus', () => selectArea('save'));
    loadArea.addEventListener('focus', () => selectArea('load'));
    
    // キーボードイベント
    saveArea.addEventListener('keydown', handleKeyPress);
    loadArea.addEventListener('keydown', handleKeyPress);
    
    // クリックイベント（フォーカス用）
    saveArea.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        saveArea.focus();
    });
    loadArea.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        loadArea.focus();
    });
    
    // マウスリーブイベント（フォーカスを維持）
    saveArea.addEventListener('mouseleave', () => {
        // フォーカスを維持
    });
    loadArea.addEventListener('mouseleave', () => {
        // フォーカスを維持
    });
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
    updateCurrentFormat();
}

// 現在の書式を更新
function updateCurrentFormat() {
    Word.run(async (context) => {
        try {
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
}

// グローバル関数として公開
window.removeFormat = removeFormat;